# -*- coding: utf-8 -*-
"""
programacao_semanal.py
=======================
App Streamlit para automação da PROGRAMAÇÃO SEMANAL de manutenção
(nivelamento de capacidade por técnico/dia) a partir de dois arquivos
carregados pelo PCM:

    Base A - Disponibilidade (HH líquido por técnico/dia da semana, com Turno)
    Base B - IW37N (lista de OPERAÇÕES do SAP PM, com Plano)

Fluxo do app:
    0. Sidebar: upload/colar das bases, SEMANA de programação (selecionável,
       sempre começa na segunda-feira; padrão = semana seguinte à atual) e
       horários dos turnos A/B/C/ADM (digitados em HH:MM).
    1. Aba "Mesa de Programação":
       1.1 Triagem: marca as Ordens da semana e clica em "Carregar Ordens".
       1.2 Escopo Fechado: referência das Operações das Ordens carregadas.
       1.3 Mesa de Atribuição: Colaborador x Ordem x Dia x Horas, uma linha
           por vaga de executante; "Aplicar Atribuições" grava; "Calcular
           Saldo" mostra saldo/sobrecarga.
    2. Aba "Cronograma (Gantt)": botão "Gerar Cronograma".
    3. Aba "Exportar": botão "Gerar Arquivo de Programação" (.xlsx/.csv).

Como rodar:
    pip install streamlit pandas numpy openpyxl plotly
    streamlit run programacao_semanal.py
"""
from __future__ import annotations

import io
import tempfile
import unicodedata
from datetime import date, time, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from streamlit.errors import StreamlitAPIException

st.set_page_config(
    page_title="PCM | Programação Semanal",
    page_icon="🗓️",
    layout="wide",
)

# =============================================================================
# 0) CONSTANTES: layout das bases de entrada e da saída
# =============================================================================

DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

# Mapeia cada dia da semana -> nome exato da coluna de C.H. na Base A
# (mantido fiel ao layout informado, incluindo a inconsistência de
# pontuação entre os dias: "C.H Segunda" vs "C.H. Quarta")
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

# "Plano" antes de "Operação"; "Operação" segue como referência, mas a
# alocação é feita pela ORDEM.
COLS_IW37N = [
    "Ordem", "Plano", "Operação", "Tipo de Ordem", "Local de Instalação",
    "Denominação do loc. instalação", "Texto Breve", "Prioridade",
    "Duração Normal", "Executantes (Nº de Pessoas)", "Data de entrada",
    "Data-base fim", "Centro de Trabalho", "Status Usuário",
]

# Layout de exportação: uma linha por Ordem x Colaborador.
COLS_SAIDA = [
    "Colaborador", "Matrícula", "Dia_Semana", "Data_Programada", "Ordem",
    "Texto Breve", "Equipamento", "Local de Instalação",
    "Horas Programadas", "Reprogramada",
]

# Turno NÃO é escolhido na Mesa de Atribuição: vem da Disponibilidade
# (um Turno por colaborador) e é usado só internamente para ancorar o Gantt.
ALOCACOES_COLS = [
    "Centro Trabalho", "Colaborador", "Matrícula", "Ordem",
    "Dia_Semana", "Horas Programadas", "Reprogramada",
]

# Colunas de identificador lidas SEMPRE como texto: evita que o pandas
# infira número e apague zeros à esquerda (ex.: Operação "0010" -> 10).
ID_COLS_AS_TEXT = ["Ordem", "Operação", "Plano", "Matrícula", "Centro Trabalho", "Centro de Trabalho", "Turno"]

# Janela do seletor de semana (a partir da semana atual)
SEMANAS_PASSADAS = 104   # ~2 anos para trás, ...       
SEMANAS_FUTURAS = 26     # ~6 meses à frente            

# Turnos padrão (o usuário redefine na sidebar)
TURNOS_PADRAO = {
    "A": (time(7, 0), time(15, 20)),
    "B": (time(15, 20), time(23, 28)),
    "C": (time(23, 28), time(7, 0)),
    "ADM": (time(7, 0), time(17, 0)),
}


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


# Cacheado pelo CONTEÚDO do arquivo: a reexecução do script a cada interação
# não relê/reprocessa o Excel se o arquivo não mudou.
_parse_tabular_cached = st.cache_data(show_spinner="Lendo arquivo...")(_parse_tabular)


def read_uploaded(file) -> pd.DataFrame:
    dtype_map = {c: str for c in ID_COLS_AS_TEXT}
    return _parse_tabular_cached(file.getvalue(), file.name, dtype_map)


def _parse_hora(texto: str, default: time) -> tuple[time, bool]:
    """Converte texto 'HH:MM' digitado manualmente em datetime.time.
    Devolve (horário, ok): se inválido, cai no default e ok=False."""
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


def semanas_disponiveis(hoje: date, passadas: int = SEMANAS_PASSADAS, futuras: int = SEMANAS_FUTURAS) -> list[date]:
    """Segundas-feiras de (passadas) semanas atrás até (futuras) semanas à frente."""
    segunda_atual = hoje - timedelta(days=hoje.weekday())
    return [segunda_atual + timedelta(weeks=i) for i in range(-passadas, futuras + 1)]


def rotulo_semana(segunda: date) -> str:
    domingo = segunda + timedelta(days=6)
    iso = segunda.isocalendar()
    return f"Semana {iso.week:02d}/{iso.year} ({segunda:%d/%m} – {domingo:%d/%m/%Y})"


@st.cache_data(show_spinner=False)
def _read_pasted_cached(text: str, dtype_map: dict) -> pd.DataFrame:
    sep = "\t" if "\t" in text.splitlines()[0] else ";"
    return pd.read_csv(io.StringIO(text), sep=sep, dtype=dtype_map)


def read_pasted(text: str) -> pd.DataFrame:
    """Lê dados colados (Excel copia como TSV: separado por tabulação)."""
    dtype_map = {c: str for c in ID_COLS_AS_TEXT}
    return _read_pasted_cached(text, dtype_map)


def load_base(label: str, key_prefix: str) -> pd.DataFrame | None:
    modo = st.radio(
        f"{label}: forma de entrada", ["Upload de arquivo", "Colar dados"],
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
# 2) PROCESSAMENTO: HH demandado, disponibilidade longa, saldo, backlog
# =============================================================================

@st.cache_data(show_spinner="Calculando HH das operações...")
def compute_hh_operacao(iw37n: pd.DataFrame) -> pd.DataFrame:
    df = iw37n.copy()
    df["Duração Normal"] = pd.to_numeric(df["Duração Normal"], errors="coerce").fillna(0)
    df["Executantes (Nº de Pessoas)"] = pd.to_numeric(df["Executantes (Nº de Pessoas)"], errors="coerce").fillna(1)
    df["HH_Operacao"] = df["Duração Normal"] * df["Executantes (Nº de Pessoas)"]
    df["Ordem"] = df["Ordem"].astype(str)
    df["Operação"] = df["Operação"].astype(str)
    # fillna ANTES do astype: .astype(str) não converte NaN em "nan" de forma confiável
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
    """Pré-lista uma linha por 'vaga de executante' de cada Ordem do escopo.
    Se a Ordem exigir N executantes (maior valor de 'Executantes (Nº de
    Pessoas)' entre suas Operações), ela aparece N vezes. Horas Programadas
    já vem com a soma da Duração Normal das operações da Ordem (cada
    executante é debitado o total cheio: trabalham em paralelo)."""
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
    """Preenche as linhas pré-listadas com o que já foi salvo para este
    Centro de Trabalho, casando por Ordem na ordem de lançamento."""
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
    """'Assinatura' leve do conteúdo de um DataFrame: detecta se a alocação
    mudou desde o último clique em Aplicar/Calcular/Gerar, sem recalcular."""
    if df is None or df.empty:
        return 0
    return int(pd.util.hash_pandas_object(df, index=True).sum())


def build_backlog(iw37n: pd.DataFrame) -> pd.DataFrame:
    """Agrega as operações por ORDEM: base da tela de triagem."""
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
    """Como build_saida, mas mantém 'Turno' (uso interno: Gantt), derivado
    do Colaborador via colaborador_turno_map (vem da Disponibilidade)."""
    if alocacoes.empty:
        return pd.DataFrame(columns=COLS_SAIDA + ["Turno"])

    df = alocacoes.copy()
    dia_para_offset = {d: i for i, d in enumerate(DIAS_SEMANA)}
    df["Data_Programada"] = df["Dia_Semana"].map(
        lambda d: pd.Timestamp(semana_inicio) + timedelta(days=dia_para_offset.get(d, 0))
    )
    df["Turno"] = df["Colaborador"].map(colaborador_turno_map).fillna("A")

    # Ordem pode ter várias Operações: usa a 1ª como referência p/ Texto/Equipamento/Local
    ref = iw37n[["Ordem", "Texto Breve", "Denominação do loc. instalação", "Local de Instalação"]].drop_duplicates("Ordem")
    ref = ref.rename(columns={"Denominação do loc. instalação": "Equipamento"})

    saida = df.merge(ref, on="Ordem", how="left")
    return saida[COLS_SAIDA + ["Turno"]]


def build_saida(alocacoes: pd.DataFrame, iw37n: pd.DataFrame, semana_inicio, colaborador_turno_map: dict) -> pd.DataFrame:
    """Layout FIXO de saída (sem Turno)."""
    interna = build_saida_interna(alocacoes, iw37n, semana_inicio, colaborador_turno_map)
    if interna.empty:
        return pd.DataFrame(columns=COLS_SAIDA)
    return interna[COLS_SAIDA]


@st.cache_data(show_spinner=False)
def build_gantt_df(saida_interna: pd.DataFrame, turnos_cfg: dict) -> pd.DataFrame:
    """Monta Início/Fim sequenciais por Colaborador x Dia x Turno, a partir
    do horário de início do turno do colaborador."""
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


# =============================================================================
# UI — MATRIZ VISUAL DE NIVELAMENTO (Colaborador x Dia da Semana)
# Só camada de interface: usa as funções de negócio já existentes
# (build_ordens_base, reconciliar_ordens, calcular_saldo, ...) sem alterá-las.
# =============================================================================

# Componente HTML/JS (arrastar e soltar + seleção por clique/teclado). É gravado
# em uma pasta temporária na inicialização e registrado via st.components.v1.
_MATRIZ_HTML = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--fg:#0f172a;--muted:#334155;--bg:#ffffff;--panel:#f1f5f9;--line:#94a3b8;--sel:#1d4ed8;
--ok-bg:#d7f5dc;--ok-fg:#064e1b;--ok-bd:#15803d;--bad-bg:#fdd9d9;--bad-fg:#7f1010;--bad-bd:#b91c1c;--off-bg:#e5e7eb;--off-fg:#1f2937;--off-bd:#6b7280}
*{box-sizing:border-box}
body{margin:0;padding:6px;font:16px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:var(--fg);background:transparent}
#root{background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:10px}
button{font:inherit;color:inherit;cursor:pointer}
.btn{min-height:42px;padding:6px 14px;border:2px solid var(--sel);background:#fff;border-radius:8px;font-weight:700}
:focus-visible{outline:3px solid #f59e0b;outline-offset:2px}
.toolbar{display:flex;gap:10px 16px;align-items:center;flex-wrap:wrap;margin-bottom:6px}
.legend{display:flex;gap:8px;flex-wrap:wrap;font-size:14px}
.lg{padding:3px 10px;border-radius:999px;font-weight:700;border:2px solid}
.lg.ok{background:var(--ok-bg);color:var(--ok-fg);border-color:var(--ok-bd)}
.lg.bad{background:var(--bad-bg);color:var(--bad-fg);border-color:var(--bad-bd)}
.lg.off{background:var(--off-bg);color:var(--off-fg);border-color:var(--off-bd)}
.msg{min-height:26px;margin:2px 0 8px;font-size:15px;font-weight:600;color:var(--muted)}
.layout{display:flex;gap:12px;align-items:flex-start;flex-wrap:wrap}
#panel{flex:0 0 310px;max-width:100%;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:8px;max-height:640px;overflow:auto}
#panel.hidden{display:none}
#panel h3{margin:2px 4px 8px;font-size:16px}
.empty{margin:8px 4px;font-weight:600}
.main{flex:1 1 560px;min-width:0}
.matrix-wrap{overflow:auto;max-height:660px;border:1px solid var(--line);border-radius:10px}
table{border-collapse:separate;border-spacing:0;width:100%}
th,td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);padding:6px;vertical-align:top}
thead th{position:sticky;top:0;background:#e2e8f0;z-index:2;text-align:center;min-width:132px}
.col-h{position:sticky;left:0;z-index:3;background:#e2e8f0;min-width:158px;text-align:left}
thead .col-h{z-index:4}
.sub{font-size:13px;font-weight:500;color:var(--muted)}
.cell{min-height:92px;min-width:132px}
.cell.ok{background:var(--ok-bg);color:var(--ok-fg);box-shadow:inset 0 0 0 2px var(--ok-bd)}
.cell.bad{background:var(--bad-bg);color:var(--bad-fg);box-shadow:inset 0 0 0 3px var(--bad-bd)}
.cell.off{background:var(--off-bg);color:var(--off-fg);box-shadow:inset 0 0 0 2px var(--off-bd)}
.cell.drop{outline:4px dashed var(--sel);outline-offset:-4px}
.saldo{font-weight:800;font-size:16px}
.card{background:#fff;border:2px solid var(--line);border-radius:10px;padding:8px;margin:0 0 8px;cursor:grab}
.card.selected,.chip.selected{border-color:var(--sel);box-shadow:0 0 0 3px #bfdbfe}
.card.dragging,.chip.dragging{opacity:.5}
.card-top{display:flex;justify-content:space-between;gap:6px;align-items:center}
.card-bottom{display:flex;justify-content:space-between;gap:6px;align-items:center;margin-top:6px}
.meta{font-size:13px;color:var(--muted)}
.txt{font-size:15px;font-weight:600;margin:2px 0}
.hh{font-weight:800}
.badge{font-size:12px;font-weight:800;color:#fff;padding:2px 8px;border-radius:999px;white-space:nowrap}
.p-alta{background:#b91c1c}.p-media{background:#92400e}.p-baixa{background:#166534}.p-na{background:#374151}
.rp{min-height:34px;padding:2px 8px;border:2px solid var(--line);background:#fff;border-radius:8px;font-weight:700}
.rp.on{background:#fde68a;border-color:#b45309;color:#451a03}
.chip{display:flex;align-items:center;gap:4px;background:#fff;color:var(--fg);border:2px solid var(--line);border-radius:8px;padding:3px 4px 3px 6px;margin-top:5px;cursor:grab}
.chip-t{flex:1;font-size:13px;font-weight:700}
.rm{min-height:32px;min-width:32px;border:2px solid var(--bad-bd);background:#fff;color:var(--bad-fg);border-radius:8px;font-weight:800}
</style>
</head>
<body>
<div id="root"></div>
<script>
(function () {
  "use strict";
  var args = {cards: [], colabs: [], dias: [], cells: {}};
  var selectedKey = null, dragKey = null, panelHidden = false, msgText = "", counter = 0;
  var root = document.getElementById("root");

  function post(m) { window.parent.postMessage(m, "*"); }
  function send(ev) {
    ev.id = Date.now() + "-" + (++counter);
    post({isStreamlitMessage: true, type: "streamlit:setComponentValue", value: ev, dataType: "json"});
  }
  function fit() {
    post({isStreamlitMessage: true, type: "streamlit:setFrameHeight", height: document.documentElement.scrollHeight + 6});
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c];
    });
  }
  function fmt(n) { return (Math.round(Number(n) * 10) / 10).toFixed(1).replace(".", ","); }
  function cardByKey(k) { return (args.cards || []).filter(function (c) { return c.key === k; })[0]; }
  function cellOf(colab, dia) { return (args.cells || {})[colab + "||" + dia]; }
  function setMsg(t) {
    msgText = t;
    var m = root.querySelector("#msg");
    if (m) { m.textContent = t; }
  }

  function chipHtml(c) {
    return '<div class="chip' + (c.key === selectedKey ? " selected" : "") + '" draggable="true" tabindex="0" data-key="' + esc(c.key) +
      '" aria-label="Ordem ' + esc(c.ordem) + ' vaga ' + esc(c.vaga) + ', ' + fmt(c.hh) + ' horas. Enter seleciona para mover.">' +
      '<span class="chip-t">OS ' + esc(c.ordem) + " · " + esc(c.vaga) + " · " + fmt(c.hh) + "h</span>" +
      '<button class="rp' + (c.reprog ? " on" : "") + '" data-act="reprog" data-key="' + esc(c.key) + '" aria-pressed="' + (!!c.reprog) +
      '" title="Reprogramada" aria-label="Alternar reprogramada">&#8635;</button>' +
      '<button class="rm" data-act="remove" data-key="' + esc(c.key) + '" aria-label="Remover da célula" title="Remover">&times;</button></div>';
  }

  function cardHtml(c) {
    var sel = c.key === selectedKey;
    return '<div class="card' + (sel ? " selected" : "") + '" draggable="true" tabindex="0" role="button" aria-pressed="' + sel +
      '" data-key="' + esc(c.key) + '" aria-label="Ordem ' + esc(c.ordem) + ', vaga ' + esc(c.vaga) + ', ' + fmt(c.hh) + ' horas, ' + esc(c.prioridade) +
      '. Enter seleciona; depois escolha a célula do colaborador e dia.">' +
      '<div class="card-top"><strong>OS ' + esc(c.ordem) + '</strong><span class="badge p-' + esc(c.nivel) + '">' + esc(c.prioridade) + "</span></div>" +
      '<div class="meta">Vaga ' + esc(c.vaga) + " · Operações: " + esc(c.operacoes) + "</div>" +
      '<div class="txt">' + esc(c.texto) + "</div>" +
      '<div class="meta">&#128205; ' + esc(c.local) + "</div>" +
      '<div class="card-bottom"><span class="hh">&#9201; ' + fmt(c.hh) + " HH</span>" +
      '<button class="rp' + (c.reprog ? " on" : "") + '" data-act="reprog" data-key="' + esc(c.key) + '" aria-pressed="' + (!!c.reprog) +
      '">&#8635; Reprogramada</button></div></div>';
  }

  function cellHtml(colab, dia) {
    var cell = cellOf(colab, dia) || {disp: 0, alocado: 0, saldo: 0, status: "off"};
    var icon = cell.status === "bad" ? "▲" : (cell.status === "off" ? "—" : "✔");
    var label = cell.status === "bad" ? "Excedido " + fmt(-cell.saldo) + " h" : (cell.status === "off" ? "Sem escala" : "Saldo " + fmt(cell.saldo) + " h");
    var placed = (args.cards || []).filter(function (c) { return c.colaborador === colab && c.dia === dia; });
    return '<td class="cell ' + esc(cell.status) + '" tabindex="0" role="gridcell" data-colab="' + esc(colab) + '" data-dia="' + esc(dia) +
      '" aria-label="' + esc(colab) + ", " + esc(dia) + ": " + esc(label) + ". " + fmt(cell.alocado) + " de " + fmt(cell.disp) + ' horas alocadas. Enter coloca o card selecionado aqui.">' +
      '<div class="saldo"><span aria-hidden="true">' + icon + "</span> " + esc(label) + "</div>" +
      '<div class="sub">' + fmt(cell.alocado) + " / " + fmt(cell.disp) + " h</div>" + placed.map(chipHtml).join("") + "</td>";
  }

  function focusDesc() {
    var a = document.activeElement;
    if (!a || !a.getAttribute) { return null; }
    if (a.id === "toggle") { return {t: "toggle"}; }
    if (a.classList && (a.classList.contains("card") || a.classList.contains("chip"))) { return {t: "k", k: a.getAttribute("data-key"), c: a.className.split(" ")[0]}; }
    if (a.classList && a.classList.contains("cell")) { return {t: "c", co: a.getAttribute("data-colab"), d: a.getAttribute("data-dia")}; }
    return null;
  }
  function restoreFocus(d) {
    if (!d) { return; }
    var els = root.querySelectorAll("#toggle,.card,.chip,.cell"), i, el;
    for (i = 0; i < els.length; i++) {
      el = els[i];
      if ((d.t === "toggle" && el.id === "toggle") ||
          (d.t === "k" && el.getAttribute("data-key") === d.k && el.className.split(" ")[0] === d.c) ||
          (d.t === "c" && el.getAttribute("data-colab") === d.co && el.getAttribute("data-dia") === d.d)) { el.focus(); return; }
    }
  }

  function render() {
    var fd = focusDesc();
    var cards = args.cards || [], dias = args.dias || [], colabs = args.colabs || [];
    var free = cards.filter(function (c) { return !c.colaborador; });
    var head = '<tr><th class="col-h" scope="col">Colaborador</th>' + dias.map(function (d) { return '<th scope="col">' + esc(d) + "</th>"; }).join("") + "</tr>";
    var body = colabs.map(function (co) {
      return '<tr><th class="col-h" scope="row">' + esc(co.nome) + '<div class="sub">Mat. ' + esc(co.matricula) + " · Turno " + esc(co.turno) + "</div></th>" +
        dias.map(function (d) { return cellHtml(co.nome, d); }).join("") + "</tr>";
    }).join("");
    var panel = '<aside id="panel" class="' + (panelHidden ? "hidden" : "") + '" aria-label="Backlog ativo da semana"><h3>Backlog ativo (' + free.length + " pendentes)</h3>" +
      (free.length ? free.map(cardHtml).join("") : '<p class="empty">✔ Todas as vagas já foram alocadas.</p>') + "</aside>";
    var toolbar = '<div class="toolbar"><button id="toggle" class="btn" aria-expanded="' + (!panelHidden) + '">' + (panelHidden ? "&#9654; Mostrar backlog" : "&#9664; Ocultar backlog") + "</button>" +
      '<div class="legend" aria-label="Legenda"><span class="lg ok">✔ Verde: saldo positivo</span><span class="lg bad">▲ Vermelho: capacidade excedida</span><span class="lg off">— Cinza: sem escala</span></div></div>';
    root.innerHTML = toolbar + '<div id="msg" class="msg" role="status" aria-live="polite">' + esc(msgText) + "</div>" +
      '<div class="layout">' + panel + '<div class="main"><div class="matrix-wrap"><table role="grid" aria-label="Matriz de nivelamento: colaboradores por dia da semana"><thead>' +
      head + "</thead><tbody>" + body + "</tbody></table></div></div></div>";
    restoreFocus(fd);
    fit();
  }

  function toggleSelect(k) {
    selectedKey = (selectedKey === k) ? null : k;
    var c = cardByKey(selectedKey);
    msgText = c ? "Selecionado: OS " + c.ordem + " vaga " + c.vaga + ". Clique (ou Enter) na célula do colaborador e dia desejados. Esc cancela." : "Seleção cancelada.";
    render();
  }
  function placeSelected(cellEl) {
    if (!selectedKey) { setMsg("Selecione primeiro um card (clique ou Enter) e depois a célula de destino."); return; }
    send({type: "place", key: selectedKey, colaborador: cellEl.getAttribute("data-colab"), dia: cellEl.getAttribute("data-dia")});
    selectedKey = null;
    msgText = "Atribuindo…";
    render();
  }
  function preview(cellEl) {
    var k = dragKey || selectedKey;
    if (!k || !cellEl) { return; }
    var c = cardByKey(k), info = cellOf(cellEl.getAttribute("data-colab"), cellEl.getAttribute("data-dia"));
    if (!c || !info) { return; }
    var same = c.colaborador === cellEl.getAttribute("data-colab") && c.dia === cellEl.getAttribute("data-dia");
    var apos = same ? info.saldo : info.saldo - c.hh;
    setMsg("Se soltar OS " + c.ordem + " (" + fmt(c.hh) + " HH) em " + cellEl.getAttribute("data-colab") + " / " + cellEl.getAttribute("data-dia") +
      ": saldo ficaria " + fmt(apos) + " h" + (apos < 0 ? " — ALERTA: excede a capacidade." : "."));
  }
  function clearDrop() {
    var d = root.querySelectorAll(".drop"), i;
    for (i = 0; i < d.length; i++) { d[i].classList.remove("drop"); }
  }

  root.addEventListener("click", function (e) {
    var t = e.target;
    if (!t || !t.closest) { return; }
    var act = t.closest("[data-act]");
    if (act) {
      var k = act.getAttribute("data-key");
      send({type: act.getAttribute("data-act") === "remove" ? "unplace" : "toggle_reprog", key: k});
      return;
    }
    if (t.closest("#toggle")) { panelHidden = !panelHidden; render(); return; }
    var cardEl = t.closest(".card, .chip");
    if (cardEl) { toggleSelect(cardEl.getAttribute("data-key")); return; }
    var cellEl = t.closest(".cell");
    if (cellEl) { placeSelected(cellEl); }
  });

  root.addEventListener("keydown", function (e) {
    var t = e.target;
    if (e.key === "Escape" && selectedKey) { selectedKey = null; msgText = "Seleção cancelada."; render(); return; }
    if ((e.key !== "Enter" && e.key !== " ") || !t || t.tagName === "BUTTON" || !t.closest) { return; }
    var cardEl = t.closest(".card, .chip");
    if (cardEl) { e.preventDefault(); toggleSelect(cardEl.getAttribute("data-key")); return; }
    var cellEl = t.closest(".cell");
    if (cellEl && t === cellEl) { e.preventDefault(); placeSelected(cellEl); }
  });

  root.addEventListener("mouseover", function (e) {
    if (selectedKey && e.target && e.target.closest) { preview(e.target.closest(".cell")); }
  });

  root.addEventListener("dragstart", function (e) {
    var el = e.target && e.target.closest ? e.target.closest(".card, .chip") : null;
    if (!el) { return; }
    dragKey = el.getAttribute("data-key");
    if (e.dataTransfer) { e.dataTransfer.setData("text/plain", dragKey); e.dataTransfer.effectAllowed = "move"; }
    el.classList.add("dragging");
  });
  root.addEventListener("dragend", function () {
    dragKey = null;
    clearDrop();
    var d = root.querySelectorAll(".dragging"), i;
    for (i = 0; i < d.length; i++) { d[i].classList.remove("dragging"); }
  });
  root.addEventListener("dragenter", function (e) {
    var cell = e.target && e.target.closest ? e.target.closest(".cell") : null;
    if (cell) { preview(cell); }
  });
  root.addEventListener("dragover", function (e) {
    if (!e.target || !e.target.closest) { return; }
    var cell = e.target.closest(".cell"), pan = e.target.closest("#panel");
    if (cell || pan) {
      e.preventDefault();
      if (e.dataTransfer) { e.dataTransfer.dropEffect = "move"; }
      if (cell) { cell.classList.add("drop"); }
    }
  });
  root.addEventListener("dragleave", function (e) {
    var cell = e.target && e.target.closest ? e.target.closest(".cell") : null;
    if (cell) { cell.classList.remove("drop"); }
  });
  root.addEventListener("drop", function (e) {
    if (!e.target || !e.target.closest) { return; }
    var cell = e.target.closest(".cell"), pan = e.target.closest("#panel");
    if (!cell && !pan) { return; }
    e.preventDefault();
    var key = (e.dataTransfer && e.dataTransfer.getData("text/plain")) || dragKey;
    clearDrop();
    dragKey = null;
    if (!key) { return; }
    var c = cardByKey(key);
    if (cell) {
      send({type: "place", key: key, colaborador: cell.getAttribute("data-colab"), dia: cell.getAttribute("data-dia")});
    } else if (c && c.colaborador) {
      send({type: "unplace", key: key});
    }
    selectedKey = null;
  });

  window.addEventListener("message", function (e) {
    var d = e.data;
    if (d && d.type === "streamlit:render") {
      args = d.args || args;
      if (selectedKey && !cardByKey(selectedKey)) { selectedKey = null; }
      msgText = msgText === "Atribuindo…" ? "" : msgText;
      render();
    }
  });
  window.addEventListener("resize", fit);
  post({isStreamlitMessage: true, type: "streamlit:componentReady", apiVersion: 1});
  render();
})();
</script>
</body>
</html>
'''

def _registrar_componente_matriz():
    pasta = Path(tempfile.gettempdir()) / "pcm_matriz_nivelamento"
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / "index.html"
    if not destino.exists() or destino.read_text(encoding="utf-8") != _MATRIZ_HTML:
        destino.write_text(_MATRIZ_HTML, encoding="utf-8")
    return components.declare_component("matriz_nivelamento", path=str(pasta))


_matriz_nivelamento = _registrar_componente_matriz()


def prioridade_info(valor) -> tuple[str, str]:
    """(rótulo, nível) para o indicador visual de prioridade do card."""
    txt = str(valor).strip().lower()
    if txt in ("1", "1.0") or "emerg" in txt:
        return "P1 · Emergencial", "alta"
    if txt in ("2", "2.0") or "urgen" in txt:
        return "P2 · Urgente", "media"
    if txt in ("3", "3.0") or "oportun" in txt:
        return "P3 · Oportunidade", "baixa"
    if txt in ("", "nan", "none"):
        return "Sem prioridade", "na"
    return f"Prioridade {str(valor).strip()}", "na"


def _chave_vaga(ordem, executante) -> str:
    return f"{ordem}|{executante}"


def montar_cards(work: pd.DataFrame, escopo_centro: pd.DataFrame) -> list[dict]:
    """Uma 'card' por vaga de executante (linha da tabela de trabalho)."""
    info: dict[str, dict] = {}
    for ordem, g in escopo_centro.groupby("Ordem"):
        nums = pd.to_numeric(g["Prioridade"], errors="coerce").dropna()
        prio = int(nums.min()) if not nums.empty else g["Prioridade"].iloc[0]
        local = g["Local de Instalação"].iloc[0]
        info[str(ordem)] = {
            "operacoes": ", ".join(sorted(g["Operação"].astype(str).unique())),
            "local": "" if pd.isna(local) else str(local),
            "prio": prio,
        }
    cards = []
    for _, r in work.iterrows():
        i = info.get(str(r["Ordem"]), {"operacoes": "", "local": "", "prio": ""})
        rotulo, nivel = prioridade_info(i["prio"])
        colab = str(r["Colaborador"]).strip()
        cards.append({
            "key": _chave_vaga(r["Ordem"], r["Executante"]),
            "ordem": str(r["Ordem"]),
            "vaga": str(r["Executante"]),
            "operacoes": i["operacoes"],
            "texto": str(r["Texto Breve"]),
            "local": i["local"],
            "hh": float(r["Horas Programadas"]),
            "prioridade": rotulo,
            "nivel": nivel,
            "reprog": bool(r["Reprogramada"]),
            "colaborador": colab,
            "dia": str(r["Dia_Semana"]) if colab else "",
        })
    return cards


def montar_celulas(disp_longa: pd.DataFrame, alocacoes_total: pd.DataFrame) -> dict:
    """Saldo por célula (Colaborador x Dia) usando a MESMA fórmula do backend
    (calcular_saldo). status: ok (saldo >= 0) | bad (capacidade excedida) | off (sem escala)."""
    saldo = calcular_saldo(disp_longa, alocacoes_total)
    cells = {}
    for _, r in saldo.iterrows():
        disp = float(r["Saldo Inicial"])
        alocado = float(r["Horas Alocadas"])
        sal = float(r["Saldo Restante"])
        if sal < -1e-9:
            status = "bad"
        elif disp <= 0 and alocado <= 0:
            status = "off"
        else:
            status = "ok"
        cells[f"{r['Colaborador']}||{r['Dia_Semana']}"] = {
            "disp": round(disp, 2), "alocado": round(alocado, 2), "saldo": round(sal, 2), "status": status,
        }
    return cells


def aplicar_evento(work: pd.DataFrame, evento: dict, mapa_matricula: dict) -> tuple[pd.DataFrame, str | None]:
    """Aplica UM evento da interface (arrastar/soltar, clique ou formulário) à
    tabela de trabalho das vagas. Não toca em st.session_state['alocacoes']
    (isso continua sendo feito só pelo botão Aplicar Atribuições)."""
    if not isinstance(evento, dict):
        return work, None
    tipo, key = evento.get("type"), evento.get("key")
    chaves = work["Ordem"].astype(str) + "|" + work["Executante"].astype(str)
    idx = work.index[chaves == key]
    if len(idx) != 1:
        return work, "Vaga não encontrada (a lista foi recarregada?)."
    i = idx[0]
    work = work.copy()
    rotulo = f"OS {work.loc[i, 'Ordem']} vaga {work.loc[i, 'Executante']}"

    if tipo in ("place", "definir"):
        colab = str(evento.get("colaborador", "")).strip()
        dia = evento.get("dia")
        if tipo == "definir" and colab == "":
            tipo = "unplace"
        else:
            if colab not in mapa_matricula or dia not in DIAS_SEMANA:
                return work, "Destino inválido."
            work.loc[i, ["Colaborador", "Matrícula", "Dia_Semana"]] = [colab, mapa_matricula[colab], dia]
            if tipo == "definir":
                work.loc[i, "Horas Programadas"] = max(float(evento.get("horas", work.loc[i, "Horas Programadas"])), 0.0)
                work.loc[i, "Reprogramada"] = bool(evento.get("reprogramada", False))
            return work, f"{rotulo} → {colab} / {dia}"
    if tipo == "unplace":
        work.loc[i, ["Colaborador", "Matrícula", "Dia_Semana"]] = ["", "", DIAS_SEMANA[0]]
        if "horas" in evento:
            work.loc[i, "Horas Programadas"] = max(float(evento["horas"]), 0.0)
            work.loc[i, "Reprogramada"] = bool(evento.get("reprogramada", False))
        return work, f"{rotulo} devolvida ao backlog"
    if tipo == "toggle_reprog":
        work.loc[i, "Reprogramada"] = not bool(work.loc[i, "Reprogramada"])
        return work, f"{rotulo}: reprogramada = {bool(work.loc[i, 'Reprogramada'])}"
    return work, None


def _rerun_fragmento() -> None:
    """Reexecuta só o fragmento; se não estiver numa reexecução de fragmento
    (ex.: execução completa da página), cai para o rerun normal."""
    try:
        st.rerun(scope="fragment")
    except StreamlitAPIException:
        st.rerun()


@st.fragment
def render_mesa_atribuicao(escopo: pd.DataFrame, disp_longa: pd.DataFrame) -> None:
    """Mesa de Atribuição como FRAGMENTO: interagir aqui reexecuta só esta função."""
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
        return

    # Tabela de trabalho (uma linha por vaga de executante) — mesma regra de antes
    editor_state_key = f"ordens_editor_data_{centro_sel}"
    if editor_state_key not in st.session_state:
        ordens_base = build_ordens_base(escopo_centro)
        master = st.session_state["alocacoes"]
        alocacoes_centro_atual = master[master["Centro Trabalho"] == centro_sel]
        st.session_state[editor_state_key] = reconciliar_ordens(ordens_base, alocacoes_centro_atual)

    mapa_matricula = dict(zip(tecnicos_centro["Colaborador"], tecnicos_centro["Matrícula"]))
    chave_comp = f"matriz_{centro_sel}"
    chave_ultimo = f"matriz_ultimo_{centro_sel}"
    chave_msg = f"matriz_msg_{centro_sel}"

    def _processar(evento) -> bool:
        """Aplica o evento uma única vez (deduplica pelo id)."""
        if not isinstance(evento, dict) or evento.get("id") in (None, st.session_state.get(chave_ultimo)):
            return False
        st.session_state[chave_ultimo] = evento["id"]
        nova, msg = aplicar_evento(st.session_state[editor_state_key], evento, mapa_matricula)
        st.session_state[editor_state_key] = nova
        st.session_state[chave_msg] = msg
        return True

    _processar(st.session_state.get(chave_comp))   # evento enviado pela matriz na interação anterior

    work = st.session_state[editor_state_key]
    atribuidas = work[work["Colaborador"].astype(str).str.strip() != ""]
    master = st.session_state["alocacoes"]
    alocacoes_total = pd.concat(
        [master[master["Centro Trabalho"] != centro_sel],
         atribuidas.assign(**{"Centro Trabalho": centro_sel})[ALOCACOES_COLS]],
        ignore_index=True,
    )
    cells = montar_celulas(disp_longa, alocacoes_total)
    nomes = set(tecnicos_centro["Colaborador"])
    em_alerta = sum(1 for k, v in cells.items() if k.split("||")[0] in nomes and v["status"] == "bad")

    # ---- Painel de situação (visual) ----
    total_vagas = len(work)
    k1, k2, k3 = st.columns(3)
    k1.metric("Vagas alocadas", f"{len(atribuidas)} / {total_vagas}")
    k2.metric("Vagas pendentes", total_vagas - len(atribuidas))
    k3.metric("Células em alerta", em_alerta, delta="capacidade excedida" if em_alerta else None, delta_color="inverse")
    st.progress(len(atribuidas) / total_vagas if total_vagas else 0.0)

    st.caption(
        "Arraste um card do backlog até a célula do colaborador e dia — ou clique no card (ou Enter) e depois "
        "na célula. Verde = saldo positivo · Vermelho = capacidade excedida · Cinza = sem escala. "
        "Clique em ↻ para marcar **Reprogramada** e em × para devolver ao backlog."
    )

    retorno = _matriz_nivelamento(
        cards=montar_cards(work, escopo_centro),
        colabs=[{"nome": str(r["Colaborador"]), "matricula": str(r["Matrícula"]), "turno": str(r["Turno"])}
                for _, r in tecnicos_centro.iterrows()],
        dias=DIAS_SEMANA,
        cells=cells,
        key=chave_comp,
        default=None,
    )
    if _processar(retorno):          # rede de segurança: evento recebido só agora
        _rerun_fragmento()

    if st.session_state.get(chave_msg):
        st.caption(f"✅ {st.session_state[chave_msg]}")

    # ---- Alternativa por seleção (sem arrastar) — campos nativos, bons para leitores de tela ----
    with st.expander("⌨️ Atribuição por seleção (alternativa ao arrastar) e ajuste de horas", expanded=False):
        rotulos = {
            _chave_vaga(r["Ordem"], r["Executante"]): f"OS {r['Ordem']} · vaga {r['Executante']} — {r['Texto Breve']}"
            for _, r in work.iterrows()
        }
        chave_sel = st.selectbox("Vaga", list(rotulos), format_func=rotulos.get, key=f"sel_vaga_{centro_sel}")
        atual = work[(work["Ordem"].astype(str) + "|" + work["Executante"].astype(str)) == chave_sel].iloc[0]
        versao = abs(hash((str(atual["Colaborador"]), str(atual["Dia_Semana"]),
                           float(atual["Horas Programadas"]), bool(atual["Reprogramada"]))))
        sufixo = f"{centro_sel}_{chave_sel}_{versao}"

        opcoes_colab = [""] + list(tecnicos_centro["Colaborador"])
        c1, c2, c3 = st.columns(3)
        colab_sel = c1.selectbox(
            "Colaborador", opcoes_colab,
            index=opcoes_colab.index(atual["Colaborador"]) if atual["Colaborador"] in opcoes_colab else 0,
            key=f"nat_colab_{sufixo}",
        )
        dia_sel = c2.selectbox(
            "Dia da semana", DIAS_SEMANA,
            index=DIAS_SEMANA.index(atual["Dia_Semana"]) if atual["Dia_Semana"] in DIAS_SEMANA else 0,
            key=f"nat_dia_{sufixo}",
        )
        horas_sel = c3.number_input(
            "Horas programadas", min_value=0.0, step=0.5, value=float(atual["Horas Programadas"]), key=f"nat_horas_{sufixo}",
        )
        reprog_sel = st.checkbox("Reprogramada", value=bool(atual["Reprogramada"]), key=f"nat_reprog_{sufixo}")
        if st.button("✅ Salvar vaga (colaborador em branco = devolve ao backlog)", key=f"nat_salvar_{sufixo}"):
            nova, msg = aplicar_evento(
                work,
                {"type": "definir", "key": chave_sel, "colaborador": colab_sel, "dia": dia_sel,
                 "horas": horas_sel, "reprogramada": reprog_sel},
                mapa_matricula,
            )
            st.session_state[editor_state_key] = nova
            st.session_state[chave_msg] = msg
            _rerun_fragmento()

    # ---- Aplicar (grava em alocações; Saldo/Gantt/Exportar leem de lá) — regra inalterada ----
    aplicar_click = st.button("💾 Aplicar Atribuições", key=f"btn_aplicar_{centro_sel}", type="primary")

    if aplicar_click:
        edited = st.session_state[editor_state_key]
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
        st.session_state[f"aplicado_fp_{centro_sel}"] = _fingerprint(edited)
        st.success("Atribuições aplicadas!")

    if st.session_state.get(f"aplicado_fp_{centro_sel}") != _fingerprint(st.session_state[editor_state_key]):
        st.caption(
            "⚠️ Há atribuições ainda não aplicadas — clique em **Aplicar Atribuições** "
            "para salvá-las (elas só entram no Saldo/Cronograma/Exportação depois disso)."
        )

    # ---- Saldo consolidado (mesmo cálculo de antes, sob demanda) ----
    with st.expander("📊 Saldo consolidado e alertas de sobrecarga (após Aplicar)", expanded=False):
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
# UI: Sidebar, upload e validação
# =============================================================================

st.title("🗓️ Programação Semanal de Manutenção (PCM)")
st.caption(
    "Nivelamento de capacidade por técnico e atribuição de Ordens (SAP PM / IW37N), "
    "com exportação consolidada para o Power BI."
)

st.markdown(
    """
    <style>
    /* Acessibilidade: alvos de clique maiores, texto legível e foco bem visível */
    div.stButton > button { min-height: 2.9rem; padding: .5rem 1.1rem; font-size: 1rem; font-weight: 600; }
    div[data-testid="stMetricValue"] { font-size: 1.9rem; }
    button:focus-visible, [role="tab"]:focus-visible { outline: 3px solid #f59e0b !important; outline-offset: 2px; }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("📥 Bases de entrada")
    with st.expander("Base A: Disponibilidade", expanded=True):
        disp_raw = load_base("Disponibilidade", "disp")
    with st.expander("Base B: IW37N (Operações)", expanded=True):
        iw37n_raw = load_base("IW37N", "iw37n")

    st.divider()
    hoje = date.today()
    semanas = semanas_disponiveis(hoje)
    proxima_semana = hoje - timedelta(days=hoje.weekday()) + timedelta(weeks=1)
    semana_inicio = st.selectbox(
        "Semana de programação",
        options=semanas,
        index=semanas.index(proxima_semana),
        format_func=rotulo_semana,
        key="semana_programacao",
        help="A semana sempre começa na segunda-feira. Por padrão vem a semana seguinte à atual.",
    )

    st.divider()
    st.header("🕐 Turnos")
    st.caption("Digite os horários manualmente (formato HH:MM): usados para ancorar o início dos blocos no Cronograma (Gantt).")
    turnos_cfg: dict[str, tuple] = {}
    for turno in ["A", "B", "C", "ADM"]:
        default_ini, default_fim = TURNOS_PADRAO[turno]
        col_ini, col_fim = st.columns(2)
        txt_ini = col_ini.text_input(f"Turno {turno}: Início", value=default_ini.strftime("%H:%M"), key=f"turno_{turno}_ini_txt")
        txt_fim = col_fim.text_input(f"Turno {turno}: Fim", value=default_fim.strftime("%H:%M"), key=f"turno_{turno}_fim_txt")
        ini, ini_ok = _parse_hora(txt_ini, default_ini)
        fim, fim_ok = _parse_hora(txt_fim, default_fim)
        if not ini_ok:
            col_ini.caption("⚠️ HH:MM inválido: usando padrão")
        if not fim_ok:
            col_fim.caption("⚠️ HH:MM inválido: usando padrão")
        turnos_cfg[turno] = (ini, fim)

if disp_raw is None or iw37n_raw is None:
    st.info("👈 Envie (ou cole) as duas bases na barra lateral: Disponibilidade e IW37N: para começar.")
    st.stop()

disp_ok = validate_columns(disp_raw, COLS_DISPONIBILIDADE, "Disponibilidade")
iw37n_ok = validate_columns(iw37n_raw, COLS_IW37N, "IW37N")
if not (disp_ok and iw37n_ok):
    st.stop()

iw37n = compute_hh_operacao(iw37n_raw)
disp_longa = disponibilidade_longa(disp_raw)
colaborador_turno_map = disp_longa.drop_duplicates("Colaborador").set_index("Colaborador")["Turno"].to_dict()

# =============================================================================
# UI: Dashboard de capacidade
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
# ABA 1: Mesa de Programação (Triagem -> Escopo Fechado -> Atribuição)
# =============================================================================
with tab_mesa:

    # -------------------------------------------------------------------
    # 1) TRIAGEM: seleção de Ordens
    # -------------------------------------------------------------------
    st.subheader("1️⃣ Triagem: selecione as Ordens que entram na semana")
    st.caption(
        "Marque as Ordens abaixo e clique em **Carregar Ordens** para trazê-las (com todas as "
        "suas Operações) para o escopo fechado. Nada é aplicado automaticamente ao marcar: assim "
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
        # sem perder as alocações já feitas (continuam em st.session_state["alocacoes"]).
        for k in list(st.session_state.keys()):
            if k.startswith("ordens_editor_data_"):
                del st.session_state[k]
        st.success(f"{len(st.session_state['ordens_selecionadas'])} Ordem(ns) carregada(s) no escopo da semana.")

    selecao_atual = set(backlog_editado.loc[backlog_editado["Incluir na Semana"], "Ordem"])
    if selecao_atual != st.session_state["ordens_selecionadas"]:
        st.caption("⚠️ Há marcações não carregadas: clique em **Carregar Ordens** para aplicá-las.")

    st.divider()

    # -------------------------------------------------------------------
    # 2) ESCOPO FECHADO: só Ordens/Operações selecionadas
    # -------------------------------------------------------------------
    st.subheader("2️⃣ Escopo Fechado da Semana")

    escopo = iw37n[iw37n["Ordem"].isin(st.session_state["ordens_selecionadas"])].copy()

    if escopo.empty:
        st.info("Nenhuma Ordem selecionada na Triagem acima ainda.")
    else:
        st.caption("Referência: detalhe das Operações de cada Ordem selecionada (a alocação abaixo é feita por Ordem).")
        st.dataframe(
            escopo[["Ordem", "Plano", "C/S Plano", "Operação", "Tipo de Ordem", "Texto Breve", "Prioridade",
                     "Duração Normal", "Executantes (Nº de Pessoas)", "HH_Operacao",
                     "Centro de Trabalho", "Status Usuário"]],
            use_container_width=True, height=220,
        )

    st.divider()

    # -------------------------------------------------------------------
    # 3) MESA DE ATRIBUIÇÃO: Colaborador x Ordem x Dia x Horas
    # -------------------------------------------------------------------
    st.subheader("3️⃣ Mesa de Atribuição")
    render_mesa_atribuicao(escopo, disp_longa)

# =============================================================================
# ABA 2: Cronograma (Gantt)
# =============================================================================
with tab_gantt:
    st.subheader("📅 Cronograma de Alocação: Colaborador x Dia x Turno")

    gerar_gantt_click = st.button("📅 Gerar Cronograma", key="btn_gerar_gantt")

    if gerar_gantt_click:
        st.session_state["gantt_saida_interna"] = build_saida_interna(st.session_state["alocacoes"], iw37n, semana_inicio, colaborador_turno_map)
        st.session_state["gantt_fp"] = _fingerprint(st.session_state["alocacoes"])

    if "gantt_saida_interna" in st.session_state and not st.session_state["gantt_saida_interna"].empty:
        fp_atual = _fingerprint(st.session_state["alocacoes"])
        if st.session_state.get("gantt_fp") != fp_atual:
            st.caption("⚠️ A alocação mudou desde a última geração: clique em **Gerar Cronograma** para atualizar.")

        gantt_df = build_gantt_df(st.session_state["gantt_saida_interna"], turnos_cfg)
        fig = px.timeline(
            gantt_df,
            x_start="Início", x_end="Fim", y="Colaborador",
            color="Turno", text="Rótulo",
            hover_data=["Texto Breve", "Horas Programadas", "Reprogramada", "Dia_Semana"],
            title="Nivelamento de Carga: Semana de " + pd.Timestamp(semana_inicio).strftime("%d/%m/%Y"),
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_traces(textposition="inside")
        fig.update_layout(height=max(350, 60 * gantt_df["Colaborador"].nunique()))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Clique em **Gerar Cronograma** para visualizar o nivelamento de carga (Colaborador x Dia x Turno).")

# =============================================================================
# ABA 3: Exportar
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
            st.caption("⚠️ A alocação mudou desde a última geração: clique em **Gerar Arquivo de Programação** para atualizar.")

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
