# Programação Semanal | PCM (SAP PM)

App Streamlit para nivelamento de capacidade e atribuição semanal de
manutentores a ordens de manutenção, a partir de duas exportações do SAP
(Disponibilidade e IW37N), com Gantt e exportação consolidada para o Power BI.

Arquivo principal: `programacao_semanal.py` · `streamlit run programacao_semanal.py`

---

## 1. Bases de entrada

### 1.1 Disponibilidade (HH líquido por técnico/dia)

| Coluna | Observação |
|---|---|
| `Centro Trabalho` | mesmo texto usado em `Centro de Trabalho` do IW37N |
| `Colaborador` | nome do manutentor |
| `Matrícula` | tratada como texto (evita perder zeros à esquerda) |
| `Turno` | **A / B / C / ADM** - turno fixo do colaborador; usado para ancorar o Gantt |
| `C.H Segunda`, `C.H Terça`, `C.H. Quarta`, `C.H. Quinta`, `C.H Sexta`, `C.H Sábado`, `C.H Domingo` | HH disponível por dia (0 nos dias sem escala) |

### 1.2 IW37N (Operações do SAP PM)

| Coluna | Observação |
|---|---|
| `Ordem` | chave da ordem de manutenção |
| `Plano` | código do Plano de Manutenção, se houver - em branco/0 = corretiva avulsa. Gera automaticamente `C/S Plano` ("Com Plano"/"Sem Plano") |
| `Operação` | número da operação dentro da ordem (ex. `0010`) — mantida como referência; a alocação é feita pela **Ordem**, não por operação |
| `Tipo de Ordem`, `Local de Instalação`, `Denominação do loc. instalação`, `Texto Breve`, `Prioridade` | descritivos |
| `Duração Normal` | horas da operação - somada por Ordem para sugerir "Horas Programadas" |
| `Executantes (Nº de Pessoas)` | usa-se o **maior** valor entre as operações da ordem para saber quantas vagas de executante gerar |
| `Data de entrada`, `Data-base fim`, `Centro de Trabalho`, `Status Usuário` | descritivos/filtros |

Colunas de identificador (`Ordem`, `Operação`, `Plano`, `Matrícula`, `Centro
Trabalho`/`Centro de Trabalho`, `Turno`) são sempre lidas como texto — o
pandas, por padrão, apagaria zeros à esquerda (ex. `"0010"` → `10`).

---

## 2. Fluxo do app

```
1️⃣ Triagem            → marca as Ordens da semana (botão "📥 Carregar Ordens")
2️⃣ Escopo Fechado     → referência só-leitura das Operações das Ordens marcadas
3️⃣ Mesa de Atribuição → aloca Colaborador × Ordem × Dia × Horas (botão "💾 Aplicar Atribuições")
                          → 🔄 Calcular Saldo (saldo de horas + sobrecarga)
📅 Cronograma          → 📅 Gerar Cronograma (Gantt Plotly, ancorado no Turno)
📤 Exportar            → 📤 Gerar Arquivo de Programação (.xlsx / .csv)
```

**Toda ação pesada é sob demanda (botão).** Nada recalcula automaticamente a
cada edição ou decisão tomada depois de diagnosticar perda de edições em
sequência (ver seção 4).

### 2.1 Triagem → Escopo Fechado

A tabela de backlog é agrupada por Ordem (`Qtd. Operações`, `HH Total`,
`Plano`/`C/S Plano`). Marcar o checkbox **não** aplica nada sozinho — só ao
clicar **Carregar Ordens**, o que evita perder marcações em cliques
sucessivos e limpa as mesas de atribuição por centro para reconstrução com o
novo escopo.

### 2.2 Mesa de Atribuição

- Isolada com `@st.fragment` — editar uma célula reexecuta só esta função,
  não a página inteira (sidebar, Triagem, outras abas).
- Uma linha por **vaga de executante** da Ordem: se `max(Executantes)` entre
  as operações da ordem for 2, a ordem aparece 2 vezes.
- `Horas Programadas` vem pré-preenchida com a soma da `Duração Normal` das
  operações da ordem — cada executante é debitado o total cheio (trabalho em
  paralelo, não dividido).
- `Matrícula` é **oculta** da tabela (via `column_order`) — calculada nos
  bastidores a partir do Colaborador escolhido, só no clique de Aplicar.
- `Turno` **não é escolhido aqui** — vem da Disponibilidade do colaborador.

### 2.3 Cronograma (Gantt)

`plotly.express.timeline`, eixo Y = Colaborador, cor = Turno. O horário de
início de cada bloco é ancorado no horário de início do Turno do
colaborador (definido manualmente na sidebar, formato `HH:MM` — texto livre,
não seletor), empilhando sequencialmente quando há mais de uma ordem no
mesmo dia/turno.

### 2.4 Exportar

Layout de saída fixo, uma linha por Ordem × Colaborador:

```
Colaborador | Matrícula | Dia_Semana | Data_Programada | Ordem |
Texto Breve | Equipamento | Local de Instalação | Horas Programadas | Reprogramada
```

---

## 3. Sidebar

- Upload (ou colar) das duas bases.
- Data de referência ("hoje" para cálculo de dias).
- Segunda-feira da semana de programação (base para `Data_Programada`).
- Horários dos turnos **A, B, C, ADM** - início/fim digitados manualmente em
  texto (`HH:MM`), não seletor; formato inválido cai no padrão com aviso.

---

## 4. Decisões e aprendizados técnicos

- **Antipadrão de perda de edição (`st.data_editor`):** reescrever o
  resultado editado de volta como `data` do próprio editor a cada rerun faz
  o Streamlit achar que "os dados mudaram por fora" e resetar o
  rastreamento de edições em andamento - perdendo a próxima ação em
  sequência. Correção: a base do editor só é escrita na criação e no clique
  do botão de aplicar; nunca a cada edição.
- **`st.fragment`** isola a Mesa de Atribuição de reruns da página inteira —
  essencial depois que Saldo/Gantt/Exportar já eram sob demanda (sem esse
  isolamento, cada edição ainda reprocessava sidebar, Triagem e outras
  abas).
- **Cache (`st.cache_data`)** na leitura de arquivo (por hash de bytes),
  nas transformações (`compute_hh_operacao`, `disponibilidade_longa`) e na
  geração do Excel - sem isso, cada rerun reparseava a planilha inteira do
  zero (medido: ~2,45s para 30 mil linhas).
- **Normalização de colunas** (NFC, remoção de BOM, espaços) necessária
  para lidar com inconsistências de export do SAP GUI.
- **`pd.read_excel(..., dtype=str)`** por coluna evita perda de zeros à
  esquerda em `Ordem`/`Operação`/`Plano`/`Matrícula`.
- **`.astype(str)` não converte `NaN` em `"nan"`** de forma confiável nesta
  versão do pandas quando a coluna é mista — usar `.fillna("").astype(str)`.

---

## 5. Como rodar

```bash
pip install streamlit pandas numpy openpyxl plotly
streamlit run programacao_semanal.py
```
