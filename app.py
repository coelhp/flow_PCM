# -*- coding: utf-8 -*-
"""
programacao_semanal.py
=======================
App Streamlit para automação da PROGRAMAÇÃO SEMANAL de manutenção
(nivelamento de capacidade por técnico/dia) a partir de dois arquivos
carregados pelo PCM:

    Base A - Disponibilidade (HH líquido por técnico/dia da semana)
    Base B - IW37N (lista de OPERAÇÕES do SAP PM — uma ordem pode ter
             várias operações, uma por linha)

Fluxo do app:
    1. Upload (ou colar) das duas bases -> validação de colunas.
    2. Dashboard de capacidade (HH disponível x HH demandado).
    3. Aba "Mesa de Programação":
       3.1 Triagem — o PCM marca quais ORDENS entram no escopo da
           semana; ao marcar uma Ordem, TODAS as suas Operações entram
           junto automaticamente (agrupamento por Ordem).
       3.2 Escopo Fechado — mostra só as Ordens/Operações selecionadas.
       3.3 Mesa de Atribuição — aloca Colaborador x Operação x Dia x
           Horas. Operações com Executantes > 1 podem receber vários
           colaboradores distintos, cada um debitado a Duração Normal
           cheia (não dividida) da sua disponibilidade individual.
    4. Aba "Cronograma (Gantt)" — visão Plotly do nivelamento por
       colaborador/dia.
    5. Aba "Exportar" — layout fixo (.xlsx/.csv) para o Power BI.

Como rodar:
    pip install -r requirements.txt
    streamlit run programacao_semanal.py
"""
from __future__ import annotations

import io
import unicodedata
from datetime import time, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="PCM | Programação Semanal",
    page_icon="🗓️",
    layout="wide",
)

# =============================================================================
# 0) CONSTANTES — layout das bases de entrada e da saída
# =============================================================================

DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# Mapeia cada dia da semana -> nome exato da coluna de C.H. na Base A
# (mantido fiel ao layout informado, incluindo a inconsistência de
# pontuação entre os dias — "C.H Segunda" vs "C.H. Quarta")
COL_CH_POR_DIA = {
    "Segunda": "C.H Segunda",
    "Terça": "C.H Terça",
    "Quarta": "C.H. Quarta",
    "Quinta": "C.H. Quinta",
    "Sexta": "C.H Sexta",
    "Sábado": "C.H Sábado",
    "Domingo": "C.H Domingo",
}

COLS_DISPONIBILIDADE = [
    "Centro Trabalho", "Colaborador", "Matrícula", "Turno",
    *COL_CH_POR_DIA.values(),
]

# "Plano" incluída antes de "Operação"; "Operação" segue existindo no
# relatório (referência), mas a alocação agora é feita por ORDEM inteira.
COLS_IW37N = [
    "Ordem", "Plano", "Operação", "Tipo de Ordem", "Local de Instalação",
    "Denominação do loc. instalação", "Texto Breve", "Prioridade",
    "Duração Normal", "Executantes (Nº de Pessoas)", "Data de entrada",
    "Data-base fim", "Centro de Trabalho", "Status Usuário",
]

# Layout de exportação — uma linha por Ordem x Colaborador (não mais por
# Operação, já que a alocação passou a ser no nível da Ordem).
COLS_SAIDA = [
    "Colaborador", "Matrícula", "Dia_Semana", "Data_Programada", "Ordem",
    "Texto Breve", "Equipamento", "Local de Instalação",
    "Horas Programadas", "Reprogramada",
]

# Turno NÃO é mais escolhido na Mesa de Atribuição — vem da Disponibilidade
# (um Turno por colaborador) e é usado só internamente para ancorar o Gantt.
ALOCACOES_COLS = [
    "Centro Trabalho", "Colaborador", "Matrícula", "Ordem",
    "Dia_Semana", "Horas Programadas", "Reprogramada",
]

# Turnos padrão (o usuário pode redefinir na sidebar)
TURNOS_PADRAO = {
    "A": (time(6, 0), time(14, 0)),
    "B": (time(14, 0), time(22, 0)),
    "C": (time(22, 0), time(6, 0)),
    "ADM": (time(7, 0), time(17, 0)),
}

# Colunas de identificador que devem ser lidas SEMPRE como texto — evita que
# o pandas infira número e apague zeros à esquerda (ex.: Operação "0010"
# virando 10, Matrícula "00123" virando 123).
ID_COLS_AS_TEXT = ["Ordem", "Operação", "Plano", "Matrícula", "Centro Trabalho", "Centro de Trabalho", "Turno"]


# =============================================================================
# 1) LEITURA / VALIDAÇÃO DE ARQUIVOS
# =============================================================================

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nomes de coluna (acentos NFC, remove BOM, colapsa espaços)."""
    def _clean(col):
        col = unicodedata.normalize("NFC", str(col))
        col = col.replace("\ufeff", "")
        return " ".join(col.split())

    df = df.copy()
    df.columns = [_clean(c) for c in df.columns]
    return df


def _parse_tabular(file_bytes: bytes, filename: str, dtype_map: dict) -> pd.DataFrame:
    buffer = io.BytesIO(file_bytes)
    if filename.lower().endswith(".csv"):
        return pd.read_csv(buffer, dtype=dtype_map)
    return pd.read_excel(buffer, dtype=dtype_map)


# Cacheado pelo CONTEÚDO do arquivo (hash dos bytes) — a reexecução do script
# a cada interação (ex.: escolher um nome numa célula) NÃO relê/reprocessa o
# Excel se o arquivo não mudou. Sem isso, cada edição reprocessava um export
# do SAP inteiro (potencialmente dezenas de milhares de linhas) do zero.
_parse_tabular_cached = st.cache_data(show_spinner="Lendo arquivo...")(_parse_tabular)


def read_uploaded(file) -> pd.DataFrame:
    dtype_map = {c: str for c in ID_COLS_AS_TEXT}
    return _parse_tabular_cached(file.getvalue(), file.name, dtype_map)


@st.cache_data(show_spinner=False)
def _read_pasted_cached(text: str, dtype_map: dict) -> pd.DataFrame:
    sep = "\t" if "\t" in text.splitlines()[0] else ";"
    return pd.read_csv(io.StringIO(text), sep=sep, dtype=dtype_map)


def _parse_hora(texto: str, default: time) -> tuple[time, bool]:
    """Converte texto 'HH:MM' digitado manualmente em datetime.time.
    Devolve (horário, ok) — se o texto for inválido, cai no default e
    ok=False (para exibirmos um aviso sem travar o app)."""
    texto = (texto or "").strip()
    try:
        partes = texto.split(":")
        if len(partes) not in (1, 2):
            raise ValueError
        h = int(partes[0])
        m = int(partes[1]) if len(partes) > 1 else 0
        if not (0 <= h <= 23 and 0 <= m <= 59):
            raise ValueError
        return time(hour=h, minute=m), True
    except (ValueError, IndexError):
        return default, False


def read_pasted(text: str) -> pd.DataFrame:
    """Lê dados colados (Excel copia como TSV — separado por tabulação)."""
    dtype_map = {c: str for c in ID_COLS_AS_TEXT}
    return _read_pasted_cached(text, dtype_map)


def load_base(label: str, key_prefix: str) -> pd.DataFrame | None:
    modo = st.radio(
        f"{label} — forma de entrada", ["Upload de arquivo", "Colar dados"],
        horizontal=True, key=f"{key_prefix}_modo",
    )
    df = None
    if modo == "Upload de arquivo":
        up = st.file_uploader(f"{label} (.xlsx ou .csv)", type=["xlsx", "csv"], key=f"{key_prefix}_upload")
        if up is not None:
            df = read_uploaded(up)
    else:
        txt = st.text_area(f"Cole aqui os dados de {label} (com cabeçalho)", height=150, key=f"{key_prefix}_paste")
        if txt.strip():
            df = read_pasted(txt)
    if df is not None:
        df = normalize_columns(df)
    return df


def validate_columns(df: pd.DataFrame, required: list[str], label: str) -> bool:
    faltando = [c for c in required if c not in df.columns]
    if faltando:
        st.warning(
            f"⚠️ A base **{label}** está sem a(s) coluna(s) esperada(s): {faltando}.\n\n"
            f"Colunas encontradas: {list(df.columns)}"
        )
        return False
    return True


# =============================================================================
# 2) PROCESSAMENTO — HH demandado, disponibilidade longa, saldo, backlog
# =============================================================================

@st.cache_data(show_spinner="Calculando HH das operações...")
def compute_hh_operacao(iw37n: pd.DataFrame) -> pd.DataFrame:
    df = iw37n.copy()
    df["Duração Normal"] = pd.to_numeric(df["Duração Normal"], errors="coerce").fillna(0)
    df["Executantes (Nº de Pessoas)"] = pd.to_numeric(df["Executantes (Nº de Pessoas)"], errors="coerce").fillna(1)
    df["HH_Operacao"] = df["Duração Normal"] * df["Executantes (Nº de Pessoas)"]
    df["Ordem"] = df["Ordem"].astype(str)
    df["Operação"] = df["Operação"].astype(str)
    df["Plano"] = df["Plano"].fillna("").astype(str).str.strip()
    df["C/S Plano"] = np.where(df["Plano"].isin(["", "0", "0.0"]), "Sem Plano", "Com Plano")
    return df


@st.cache_data(show_spinner="Processando disponibilidade...")
def disponibilidade_longa(disp: pd.DataFrame) -> pd.DataFrame:
    df = disp.copy()
    df["Matrícula"] = df["Matrícula"].astype(str)
    linhas = []
    for dia, col in COL_CH_POR_DIA.items():
        if col not in df.columns:
            continue
        sub = df[["Centro Trabalho", "Colaborador", "Matrícula", "Turno", col]].copy()
        sub = sub.rename(columns={col: "Saldo Inicial"})
        sub["Dia_Semana"] = dia
        sub["Saldo Inicial"] = pd.to_numeric(sub["Saldo Inicial"], errors="coerce").fillna(0)
        linhas.append(sub)
    return pd.concat(linhas, ignore_index=True) if linhas else pd.DataFrame(
        columns=["Centro Trabalho", "Colaborador", "Matrícula", "Turno", "Dia_Semana", "Saldo Inicial"]
    )


def calcular_saldo(disp_longa: pd.DataFrame, alocacoes: pd.DataFrame) -> pd.DataFrame:
    if alocacoes.empty:
        alocado = pd.DataFrame(columns=["Matrícula", "Dia_Semana", "Horas Alocadas"])
    else:
        alocado = (
            alocacoes.groupby(["Matrícula", "Dia_Semana"], as_index=False)["Horas Programadas"]
            .sum()
            .rename(columns={"Horas Programadas": "Horas Alocadas"})
        )
    out = disp_longa.merge(alocado, on=["Matrícula", "Dia_Semana"], how="left")
    out["Horas Alocadas"] = out["Horas Alocadas"].fillna(0)
    out["Saldo Restante"] = out["Saldo Inicial"] - out["Horas Alocadas"]
    return out


def build_ordens_base(escopo_centro: pd.DataFrame) -> pd.DataFrame:
    """Pré-lista uma linha por 'vaga de executante' de cada Ordem do
    escopo. Se a Ordem exigir N executantes (maior valor de 'Executantes
    (Nº de Pessoas)' entre suas Operações), ela aparece N vezes — uma para
    cada executante ser atribuído. Horas Programadas já vem preenchida com
    a soma da Duração Normal de todas as operações da Ordem (cada
    executante é debitado o total cheio, pois trabalham em paralelo)."""
    agg = escopo_centro.groupby("Ordem", as_index=False).agg(
        Plano=("Plano", "first"),
        CS_Plano=("C/S Plano", "first"),
        Texto_Breve=("Texto Breve", "first"),
        Qtd_Operacoes=("Operação", "nunique"),
        Duracao_Total=("Duração Normal", "sum"),
        Executantes_Necessarios=("Executantes (Nº de Pessoas)", "max"),
    )
    agg = agg.rename(columns={
        "CS_Plano": "C/S Plano", "Texto_Breve": "Texto Breve",
        "Qtd_Operacoes": "Qtd. Operações", "Duracao_Total": "Horas Programadas",
        "Executantes_Necessarios": "Executantes Necessários",
    })

    linhas = []
    for _, ordem_row in agg.iterrows():
        n_exec = int(ordem_row["Executantes Necessários"]) if ordem_row["Executantes Necessários"] >= 1 else 1
        for slot in range(1, n_exec + 1):
            linhas.append({
                "Ordem": ordem_row["Ordem"],
                "Plano": ordem_row["Plano"],
                "C/S Plano": ordem_row["C/S Plano"],
                "Texto Breve": ordem_row["Texto Breve"],
                "Qtd. Operações": ordem_row["Qtd. Operações"],
                "Executante": f"{slot}/{n_exec}",
                "Colaborador": "",
                "Matrícula": "",
                "Dia_Semana": DIAS_SEMANA[0],
                "Horas Programadas": float(ordem_row["Horas Programadas"]),
                "Reprogramada": False,
            })
    cols = ["Ordem", "Plano", "C/S Plano", "Texto Breve", "Qtd. Operações", "Executante",
            "Colaborador", "Matrícula", "Dia_Semana", "Horas Programadas", "Reprogramada"]
    return pd.DataFrame(linhas, columns=cols)


def reconciliar_ordens(ordens_base: pd.DataFrame, alocacoes_centro: pd.DataFrame) -> pd.DataFrame:
    """Preenche as linhas pré-listadas (por vaga de executante da Ordem)
    com o que já foi salvo anteriormente para este Centro de Trabalho,
    casando por Ordem na ordem de lançamento — preserva o trabalho do
    planejador ao trocar de Centro e voltar."""
    if alocacoes_centro.empty:
        return ordens_base
    ordens_base = ordens_base.copy()
    for ordem, grupo in alocacoes_centro.groupby("Ordem"):
        idx_slots = ordens_base.index[ordens_base["Ordem"] == ordem].tolist()
        for slot_idx, (_, aloc) in zip(idx_slots, grupo.iterrows()):
            ordens_base.loc[slot_idx, "Colaborador"] = aloc["Colaborador"]
            ordens_base.loc[slot_idx, "Matrícula"] = aloc["Matrícula"]
            ordens_base.loc[slot_idx, "Dia_Semana"] = aloc["Dia_Semana"]
            ordens_base.loc[slot_idx, "Horas Programadas"] = aloc["Horas Programadas"]
            ordens_base.loc[slot_idx, "Reprogramada"] = aloc["Reprogramada"]
    return ordens_base


@st.cache_data(show_spinner="Gerando arquivo Excel...")
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Programacao_Semanal")
    return buf.getvalue()


def _fingerprint(df: pd.DataFrame) -> int:
    """'Assinatura' leve do conteúdo de um DataFrame — usada para detectar
    se st.session_state['alocacoes'] mudou desde o último clique em
    Calcular Saldo / Gerar Cronograma / Gerar Arquivo, sem precisar
    recalcular nada (só compara o fingerprint salvo com o atual)."""
    if df is None or df.empty:
        return 0
    return int(pd.util.hash_pandas_object(df, index=True).sum())


def build_backlog(iw37n: pd.DataFrame) -> pd.DataFrame:
    """Agrega as operações por ORDEM — base da tela de triagem."""
    agg = iw37n.groupby("Ordem", as_index=False).agg(
        Centro_de_Trabalho=("Centro de Trabalho", "first"),
        Tipo_de_Ordem=("Tipo de Ordem", "first"),
        Prioridade=("Prioridade", "first"),
        Plano=("Plano", "first"),
        CS_Plano=("C/S Plano", "first"),
        Qtd_Operacoes=("Operação", "nunique"),
        HH_Total=("HH_Operacao", "sum"),
        Data_de_Entrada=("Data de entrada", "min"),
        Data_Base_Fim=("Data-base fim", "min"),
    )
    agg = agg.rename(columns={
        "Centro_de_Trabalho": "Centro de Trabalho",
        "Tipo_de_Ordem": "Tipo de Ordem",
        "CS_Plano": "C/S Plano",
        "Qtd_Operacoes": "Qtd. Operações",
        "HH_Total": "HH Total",
        "Data_de_Entrada": "Data de Entrada",
        "Data_Base_Fim": "Data Base Fim",
    })
    return agg.sort_values(["Centro de Trabalho", "Ordem"]).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def build_saida_interna(alocacoes: pd.DataFrame, iw37n: pd.DataFrame, semana_inicio, colaborador_turno_map: dict) -> pd.DataFrame:
    """Como build_saida, mas mantém a coluna 'Turno' (uso interno — Gantt),
    derivada do Colaborador via colaborador_turno_map (Turno vem da
    Disponibilidade, não é mais escolhido na Mesa de Atribuição). O layout
    de exportação oficial (COLS_SAIDA) não inclui Turno."""
    if alocacoes.empty:
        return pd.DataFrame(columns=COLS_SAIDA + ["Turno"])

    df = alocacoes.copy()
    dia_para_offset = {d: i for i, d in enumerate(DIAS_SEMANA)}
    df["Data_Programada"] = df["Dia_Semana"].map(
        lambda d: pd.Timestamp(semana_inicio) + timedelta(days=dia_para_offset.get(d, 0))
    )
    df["Turno"] = df["Colaborador"].map(colaborador_turno_map).fillna("A")

    # Ordem pode ter várias Operações — usa a 1ª como referência p/ Texto/Equipamento/Local
    ref = iw37n[["Ordem", "Texto Breve", "Denominação do loc. instalação", "Local de Instalação"]].drop_duplicates("Ordem")
    ref = ref.rename(columns={"Denominação do loc. instalação": "Equipamento"})

    saida = df.merge(ref, on="Ordem", how="left")
    return saida[COLS_SAIDA + ["Turno"]]


def build_saida(alocacoes: pd.DataFrame, iw37n: pd.DataFrame, semana_inicio, colaborador_turno_map: dict) -> pd.DataFrame:
    """Monta o DataFrame consolidado no layout FIXO de saída (sem Turno —
    Turno é um detalhe operacional interno, não faz parte do layout
    combinado de exportação)."""
    interna = build_saida_interna(alocacoes, iw37n, semana_inicio, colaborador_turno_map)
    if interna.empty:
        return pd.DataFrame(columns=COLS_SAIDA)
    return interna[COLS_SAIDA]


@st.cache_data(show_spinner=False)
def build_gantt_df(saida_interna: pd.DataFrame, turnos_cfg: dict) -> pd.DataFrame:
    """A partir do layout de saída (com Turno), monta Início/Fim
    sequenciais por Colaborador x Dia x Turno (empilhando as Ordens a
    partir do horário de início do turno) para o Gantt."""
    if saida_interna.empty:
        return saida_interna.assign(Início=pd.Series(dtype="datetime64[ns]"), Fim=pd.Series(dtype="datetime64[ns]"), Rótulo="")

    df = saida_interna.sort_values(["Colaborador", "Data_Programada", "Turno", "Ordem"]).reset_index(drop=True)
    cursor: dict[tuple, pd.Timestamp] = {}
    inicios, fins = [], []
    turno_default = turnos_cfg.get("A", (time(8, 0), time(16, 0)))
    for _, row in df.iterrows():
        turno = row.get("Turno", "A")
        inicio_turno, _ = turnos_cfg.get(turno, turno_default)
        key = (row["Colaborador"], row["Data_Programada"], turno)
        base = cursor.get(
            key,
            pd.Timestamp(row["Data_Programada"]) + pd.Timedelta(hours=inicio_turno.hour, minutes=inicio_turno.minute),
        )
        horas = float(row["Horas Programadas"]) if not pd.isna(row["Horas Programadas"]) else 0.0
        inicio = base
        fim = inicio + pd.Timedelta(hours=max(horas, 0.25))  # evita barra de duração zero
        cursor[key] = fim
        inicios.append(inicio)
        fins.append(fim)

    df["Início"] = inicios
    df["Fim"] = fins
    df["Rótulo"] = "OS " + df["Ordem"].astype(str)
    return df


@st.fragment
def render_mesa_atribuicao(escopo: pd.DataFrame, disp_longa: pd.DataFrame) -> None:
    """Mesa de Atribuição isolada como FRAGMENTO do Streamlit: editar uma
    célula aqui dentro (Colaborador, Dia_Semana, Horas...) reexecuta SÓ
    esta função — não a página inteira (sidebar, Triagem, abas Gantt/
    Exportar). É isso que dá a agilidade pedida: sem isso, CADA edição
    reexecutava o script do zero, e cliques rápidos em sequência podiam
    ser perdidos/revertidos enquanto o rerun anterior ainda processava."""
    if escopo.empty:
        st.info("Nenhuma Ordem no escopo fechado ainda.")
        return

    centros = sorted(escopo["Centro de Trabalho"].dropna().unique())
    centro_sel = st.selectbox("Centro de Trabalho", centros, key="centro_sel_mesa")

    tecnicos_centro = disp_longa[disp_longa["Centro Trabalho"] == centro_sel][
        ["Colaborador", "Matrícula", "Turno"]
    ].drop_duplicates()
    escopo_centro = escopo[escopo["Centro de Trabalho"] == centro_sel].copy()

    if tecnicos_centro.empty:
        st.warning("Nenhum técnico cadastrado na Base A para este Centro de Trabalho.")
    if escopo_centro.empty:
        st.warning("Nenhuma Operação do escopo fechado pertence a este Centro de Trabalho.")

    st.write(
        "**Aloque as Ordens** — cada vaga de executante já aparece listada abaixo com as Horas "
        "Programadas pré-preenchidas (soma da Duração Normal de todas as Operações da Ordem); "
        "basta selecionar o **Colaborador** (o Turno é o cadastrado na Disponibilidade dele — "
        "ajuste Dia/Horas se necessário). Se a Ordem exigir 2 executantes, ela aparece 2 vezes."
    )

    editor_state_key = f"ordens_editor_data_{centro_sel}"
    if editor_state_key not in st.session_state:
        ordens_base = build_ordens_base(escopo_centro)
        master = st.session_state["alocacoes"]
        alocacoes_centro_atual = master[master["Centro Trabalho"] == centro_sel]
        st.session_state[editor_state_key] = reconciliar_ordens(ordens_base, alocacoes_centro_atual)

    mapa_matricula = dict(zip(tecnicos_centro["Colaborador"], tecnicos_centro["Matrícula"]))

    edited = st.data_editor(
        st.session_state[editor_state_key],
        num_rows="fixed",
        use_container_width=True,
        key=f"editor_{centro_sel}",
        column_order=["Ordem", "Plano", "C/S Plano", "Texto Breve", "Qtd. Operações", "Executante",
                       "Colaborador", "Dia_Semana", "Horas Programadas", "Reprogramada"],
        column_config={
            "Ordem": st.column_config.TextColumn("Ordem", disabled=True),
            "Plano": st.column_config.TextColumn("Plano", disabled=True),
            "C/S Plano": st.column_config.TextColumn("C/S Plano", disabled=True),
            "Texto Breve": st.column_config.TextColumn("Texto Breve", disabled=True, width="large"),
            "Qtd. Operações": st.column_config.NumberColumn("Qtd. Operações", disabled=True),
            "Executante": st.column_config.TextColumn(
                "Executante", disabled=True, help="Posição da vaga dentro da Ordem (ex.: 1/2 = 1º de 2 necessários).",
            ),
            "Colaborador": st.column_config.SelectboxColumn(
                "Colaborador", options=[""] + tecnicos_centro["Colaborador"].tolist(),
                help="Único campo obrigatório — selecione quem vai executar esta vaga.",
            ),
            "Dia_Semana": st.column_config.SelectboxColumn("Dia_Semana", options=DIAS_SEMANA),
            "Horas Programadas": st.column_config.NumberColumn(
                "Horas Programadas", min_value=0.0, step=0.5,
                help="Pré-preenchida com a soma da Duração Normal das Operações da Ordem — ajuste se necessário.",
            ),
            "Reprogramada": st.column_config.CheckboxColumn("Reprogramada", default=False),
        },
    )

    st.session_state[editor_state_key] = edited

    edited_completo = edited.copy()
    edited_completo["Matrícula"] = edited_completo["Colaborador"].map(mapa_matricula)
    edited_completo["Centro Trabalho"] = centro_sel
    edited_validas = edited_completo[edited_completo["Colaborador"].astype(str).str.strip() != ""].copy()

    master = st.session_state["alocacoes"]
    st.session_state["alocacoes"] = pd.concat(
        [master[master["Centro Trabalho"] != centro_sel],
         edited_validas[ALOCACOES_COLS] if not edited_validas.empty else pd.DataFrame(columns=ALOCACOES_COLS)],
        ignore_index=True,
    )

    st.divider()

    st.write("**Saldo de horas e alertas de sobrecarga:**")
    calcular_click = st.button("🔄 Calcular Saldo", key="btn_calcular_saldo")

    if calcular_click:
        st.session_state["saldo_resultado"] = calcular_saldo(disp_longa, st.session_state["alocacoes"])
        st.session_state["saldo_fp"] = _fingerprint(st.session_state["alocacoes"])

    if "saldo_resultado" in st.session_state:
        fp_atual = _fingerprint(st.session_state["alocacoes"])
        if st.session_state.get("saldo_fp") != fp_atual:
            st.caption("⚠️ A alocação mudou desde o último cálculo — clique em **Calcular Saldo** para atualizar.")

        saldo_centro = st.session_state["saldo_resultado"]
        saldo_centro = saldo_centro[saldo_centro["Centro Trabalho"] == centro_sel]
        pivot_saldo = saldo_centro.pivot_table(
            index=["Colaborador", "Matrícula"], columns="Dia_Semana", values="Saldo Restante", aggfunc="sum"
        ).reindex(columns=DIAS_SEMANA)

        def _highlight_negativo(v):
            if pd.isna(v):
                return ""
            return "background-color:#ffcccc; color:#900" if v < 0 else ""

        st.dataframe(pivot_saldo.style.map(_highlight_negativo).format("{:.1f}"), use_container_width=True)

        if (pivot_saldo < 0).any().any():
            st.error("⚠️ Há técnico(s) com saldo negativo (capacidade estourada) neste Centro de Trabalho.")

        overload = saldo_centro[saldo_centro["Saldo Restante"] < 0]
        if not overload.empty:
            linhas_overload = "\n".join(
                f"- **{row['Colaborador']}** ({row['Dia_Semana']}): saldo {row['Saldo Restante']:.1f} h"
                for _, row in overload.iterrows()
            )
            st.warning(f"🔴 Sobrecarga:\n\n{linhas_overload}")
    else:
        st.info("Clique em **Calcular Saldo** para ver o saldo de horas e alertas de sobrecarga.")


# =============================================================================
# 3) ESTADO
# =============================================================================

if "alocacoes" not in st.session_state:
    st.session_state["alocacoes"] = pd.DataFrame(columns=ALOCACOES_COLS)
if "ordens_selecionadas" not in st.session_state:
    st.session_state["ordens_selecionadas"] = set()


# =============================================================================
# UI — Upload e validação
# =============================================================================

st.title("🗓️ Programação Semanal de Manutenção (PCM)")
st.caption(
    "Nivelamento de capacidade por técnico e atribuição diária de operações (SAP PM / IW37N), "
    "com exportação consolidada para o Power BI."
)

with st.sidebar:
    st.header("📥 Bases de entrada")
    with st.expander("Base A — Disponibilidade", expanded=True):
        disp_raw = load_base("Disponibilidade", "disp")
    with st.expander("Base B — IW37N (Operações)", expanded=True):
        iw37n_raw = load_base("IW37N", "iw37n")

    st.divider()
    semana_inicio = st.date_input(
        "Segunda-feira da semana de programação",
        help="Usada para calcular a Data_Programada de cada dia da semana no arquivo final e no Gantt.",
    )

    st.divider()
    st.header("🕐 Turnos")
    st.caption("Digite os horários manualmente (formato HH:MM) — usados para ancorar o início dos blocos no Cronograma (Gantt).")
    turnos_cfg: dict[str, tuple] = {}
    for turno in ["A", "B", "C", "ADM"]:
        default_ini, default_fim = TURNOS_PADRAO[turno]
        col_ini, col_fim = st.columns(2)
        txt_ini = col_ini.text_input(f"Turno {turno} — Início", value=default_ini.strftime("%H:%M"), key=f"turno_{turno}_ini_txt")
        txt_fim = col_fim.text_input(f"Turno {turno} — Fim", value=default_fim.strftime("%H:%M"), key=f"turno_{turno}_fim_txt")
        ini, ini_ok = _parse_hora(txt_ini, default_ini)
        fim, fim_ok = _parse_hora(txt_fim, default_fim)
        if not ini_ok:
            col_ini.caption("⚠️ HH:MM inválido — usando padrão")
        if not fim_ok:
            col_fim.caption("⚠️ HH:MM inválido — usando padrão")
        turnos_cfg[turno] = (ini, fim)

if disp_raw is None or iw37n_raw is None:
    st.info("👈 Envie (ou cole) as duas bases na barra lateral — Disponibilidade e IW37N — para começar.")
    st.stop()

disp_ok = validate_columns(disp_raw, COLS_DISPONIBILIDADE, "Disponibilidade")
iw37n_ok = validate_columns(iw37n_raw, COLS_IW37N, "IW37N")
if not (disp_ok and iw37n_ok):
    st.stop()

iw37n = compute_hh_operacao(iw37n_raw)
disp_longa = disponibilidade_longa(disp_raw)
colaborador_turno_map = disp_longa.drop_duplicates("Colaborador").set_index("Colaborador")["Turno"].to_dict()

# =============================================================================
# UI — Dashboard de capacidade
# =============================================================================

hh_disponivel = disp_longa["Saldo Inicial"].sum()
hh_demandado = iw37n["HH_Operacao"].sum()
pct_carregamento = (hh_demandado / hh_disponivel * 100) if hh_disponivel else 0

c1, c2, c3 = st.columns(3)
c1.metric("HH Total Disponível (semana)", f"{hh_disponivel:,.1f} h")
c2.metric("HH Total Demandado (todo o backlog)", f"{hh_demandado:,.1f} h")
c3.metric("% Carregamento Geral da Fábrica", f"{pct_carregamento:,.1f}%",
          delta=None if pct_carregamento <= 100 else "Sobrecarga", delta_color="inverse")

st.divider()

tab_mesa, tab_gantt, tab_export = st.tabs(["🧰 Mesa de Programação", "📅 Cronograma (Gantt)", "📤 Exportar"])

# =============================================================================
# ABA 1 — Mesa de Programação (Triagem -> Escopo Fechado -> Atribuição)
# =============================================================================
with tab_mesa:

    # -------------------------------------------------------------------
    # 1) TRIAGEM — seleção de Ordens (agrupa automaticamente as Operações)
    # -------------------------------------------------------------------
    st.subheader("1️⃣ Triagem — selecione as Ordens que entram na semana")
    st.caption(
        "Marque as Ordens abaixo e clique em **Carregar Ordens** para trazê-las (com todas as "
        "suas Operações) para o escopo fechado. Nada é aplicado automaticamente ao marcar — assim "
        "você pode marcar várias Ordens seguidas sem perder cliques."
    )

    backlog = build_backlog(iw37n)
    backlog.insert(0, "Incluir na Semana", backlog["Ordem"].isin(st.session_state["ordens_selecionadas"]))

    backlog_editado = st.data_editor(
        backlog,
        hide_index=True,
        use_container_width=True,
        height=280,
        key="backlog_editor",
        disabled=[c for c in backlog.columns if c != "Incluir na Semana"],
        column_config={
            "Incluir na Semana": st.column_config.CheckboxColumn("Incluir na Semana"),
            "HH Total": st.column_config.NumberColumn("HH Total", format="%.1f"),
        },
    )

    carregar_ordens_click = st.button("📥 Carregar Ordens", key="btn_carregar_ordens")
    if carregar_ordens_click:
        st.session_state["ordens_selecionadas"] = set(
            backlog_editado.loc[backlog_editado["Incluir na Semana"], "Ordem"]
        )
        # Força reconstrução das mesas de atribuição (por centro) com o novo escopo,
        # sem perder as alocações já feitas (que continuam em st.session_state["alocacoes"]).
        for k in list(st.session_state.keys()):
            if k.startswith("ordens_editor_data_"):
                del st.session_state[k]
        st.success(f"{len(st.session_state['ordens_selecionadas'])} Ordem(ns) carregada(s) no escopo da semana.")

    selecao_atual = set(backlog_editado.loc[backlog_editado["Incluir na Semana"], "Ordem"])
    if selecao_atual != st.session_state["ordens_selecionadas"]:
        st.caption("⚠️ Há marcações não carregadas — clique em **Carregar Ordens** para aplicá-las.")

    st.divider()

    # -------------------------------------------------------------------
    # 2) ESCOPO FECHADO — só Ordens/Operações selecionadas
    # -------------------------------------------------------------------
    st.subheader("2️⃣ Escopo Fechado da Semana")

    escopo = iw37n[iw37n["Ordem"].isin(st.session_state["ordens_selecionadas"])].copy()

    if escopo.empty:
        st.info("Nenhuma Ordem selecionada na Triagem acima ainda.")
    else:
        st.caption("Referência — detalhe das Operações de cada Ordem selecionada (a alocação abaixo é feita por Ordem inteira).")
        st.dataframe(
            escopo[["Ordem", "Plano", "C/S Plano", "Operação", "Tipo de Ordem", "Texto Breve", "Prioridade",
                     "Duração Normal", "Executantes (Nº de Pessoas)", "HH_Operacao",
                     "Centro de Trabalho", "Status Usuário"]],
            use_container_width=True, height=220,
        )

    st.divider()

    # -------------------------------------------------------------------
    # 3) MESA DE ATRIBUIÇÃO — Colaborador x Ordem x Dia x Horas
    # -------------------------------------------------------------------
    st.subheader("3️⃣ Mesa de Atribuição")
    render_mesa_atribuicao(escopo, disp_longa)

# =============================================================================
# ABA 2 — Cronograma (Gantt)
# =============================================================================
with tab_gantt:
    st.subheader("📅 Cronograma de Alocação — Colaborador x Dia x Turno")

    gerar_gantt_click = st.button("📅 Gerar Cronograma", key="btn_gerar_gantt")

    if gerar_gantt_click:
        st.session_state["gantt_saida_interna"] = build_saida_interna(st.session_state["alocacoes"], iw37n, semana_inicio, colaborador_turno_map)
        st.session_state["gantt_fp"] = _fingerprint(st.session_state["alocacoes"])

    if "gantt_saida_interna" in st.session_state and not st.session_state["gantt_saida_interna"].empty:
        fp_atual = _fingerprint(st.session_state["alocacoes"])
        if st.session_state.get("gantt_fp") != fp_atual:
            st.caption("⚠️ A alocação mudou desde a última geração — clique em **Gerar Cronograma** para atualizar.")

        gantt_df = build_gantt_df(st.session_state["gantt_saida_interna"], turnos_cfg)
        fig = px.timeline(
            gantt_df,
            x_start="Início", x_end="Fim", y="Colaborador",
            color="Turno", text="Rótulo",
            hover_data=["Texto Breve", "Horas Programadas", "Reprogramada", "Dia_Semana"],
            title="Nivelamento de Carga — Semana de " + pd.Timestamp(semana_inicio).strftime("%d/%m/%Y"),
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_traces(textposition="inside")
        fig.update_layout(height=max(350, 60 * gantt_df["Colaborador"].nunique()))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Clique em **Gerar Cronograma** para visualizar o nivelamento de carga (Colaborador x Dia x Turno).")

# =============================================================================
# ABA 3 — Exportar
# =============================================================================
with tab_export:
    st.subheader("📤 Exportar programação consolidada")

    gerar_export_click = st.button("📤 Gerar Arquivo de Programação", key="btn_gerar_export")

    if gerar_export_click:
        st.session_state["saida_final"] = build_saida(st.session_state["alocacoes"], iw37n, semana_inicio, colaborador_turno_map)
        st.session_state["export_fp"] = _fingerprint(st.session_state["alocacoes"])

    if "saida_final" in st.session_state and not st.session_state["saida_final"].empty:
        fp_atual = _fingerprint(st.session_state["alocacoes"])
        if st.session_state.get("export_fp") != fp_atual:
            st.caption("⚠️ A alocação mudou desde a última geração — clique em **Gerar Arquivo de Programação** para atualizar.")

        saida = st.session_state["saida_final"]
        st.dataframe(saida, use_container_width=True, height=300)

        colx, coly = st.columns(2)
        with colx:
            st.download_button(
                "⬇️ Baixar programação (.xlsx)",
                to_excel_bytes(saida),
                file_name="programacao_semanal.xlsx",
            )
        with coly:
            st.download_button(
                "⬇️ Baixar programação (.csv)",
                saida.to_csv(index=False, sep=";").encode("utf-8-sig"),
                file_name="programacao_semanal.csv",
            )
    else:
        st.info("Clique em **Gerar Arquivo de Programação** para montar o consolidado e liberar o download.")
