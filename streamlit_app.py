"""
Codificador IFec — versao Streamlit com novo sistema visual
(glass, off-white quente, navy modernizado + amber, Geist).

Mantém toda a logica de backend do app original. So o front foi reescrito.
"""

import os
import time
from io import BytesIO
from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st


# =============================================================================
# CONFIGURACAO DA PAGINA
# =============================================================================
st.set_page_config(
    page_title="Codificador IFec",
    page_icon="logo_ifec.png",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =============================================================================
# DESIGN SYSTEM — tokens via CSS custom properties + estilizacao do Streamlit
# =============================================================================
DESIGN_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&family=Geist+Mono:wght@400;500;600&display=swap');

:root {
  --ink-900:#0A1628; --ink-800:#142540; --ink-700:#1F3461; --ink-600:#2C4A82;
  --ink-500:#486AAD; --ink-400:#7A93C7; --ink-300:#B5C4DD; --ink-200:#DBE3EF;
  --ink-100:#ECF0F7; --ink-50:#F5F7FB;
  --amber-700:#B57514; --amber-600:#D08C1F; --amber-500:#E0A02C;
  --amber-400:#ECBA5C; --amber-200:#F6DFA9; --amber-100:#FBEFD2;
  --paper:#F4F2EC; --paper-2:#EEEBE2; --bone:#FAF8F2; --white:#FFFFFF;
  --line:rgba(10,22,40,0.08); --line-2:rgba(10,22,40,0.14); --hairline:rgba(10,22,40,0.04);
  --success:#2F7D5C; --success-bg:#DFF1E7; --warning:#B47514; --warning-bg:#FBEFD2;
  --danger:#B33A3A; --danger-bg:#F6DCDC;
  --glass-bg:rgba(255,255,255,0.62); --glass-edge:rgba(255,255,255,0.85);
  --glass-shadow:
    0 1px 0 rgba(255,255,255,0.6) inset,
    0 0 0 1px rgba(255,255,255,0.35) inset,
    0 1px 1px rgba(10,22,40,0.04),
    0 12px 32px -12px rgba(10,22,40,0.18);
  --font-sans:"Geist","Söhne",-apple-system,BlinkMacSystemFont,system-ui,sans-serif;
  --font-mono:"Geist Mono","JetBrains Mono",ui-monospace,Menlo,monospace;
}

/* ===== Reset / fundo ===== */
html, body, [class*="css"] {
  font-family: var(--font-sans) !important;
  -webkit-font-smoothing: antialiased;
  letter-spacing: -0.005em;
  color: var(--ink-900);
}
.stApp {
  background:
    radial-gradient(1200px 600px at 20% -10%, rgba(224,160,44,0.10), transparent 60%),
    radial-gradient(900px 700px at 110% 110%, rgba(72,106,173,0.10), transparent 60%),
    linear-gradient(180deg, #F4F2EC 0%, #EEEBE2 100%) !important;
}
.stApp::before {
  content: "";
  position: fixed; inset: 0;
  background-image:
    linear-gradient(rgba(10,22,40,0.025) 1px, transparent 1px),
    linear-gradient(90deg, rgba(10,22,40,0.025) 1px, transparent 1px);
  background-size: 56px 56px;
  mask-image: radial-gradient(900px 700px at 50% 30%, black 30%, transparent 80%);
  pointer-events: none; z-index: 0;
}
header[data-testid="stHeader"] { background: transparent; height: 0; }
.block-container {
  padding-top: 1.4rem !important;
  padding-bottom: 3rem !important;
  max-width: 1200px;
  position: relative; z-index: 1;
}

/* ===== Sidebar ===== */
[data-testid="stSidebar"] > div:first-child {
  background: linear-gradient(180deg, rgba(255,255,255,0.85) 0%, rgba(255,255,255,0.55) 100%);
  border-right: 1px solid var(--line);
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  padding: 1.4rem 1rem 1rem 1.2rem;
}
[data-testid="stSidebar"] * { color: var(--ink-900); }
[data-testid="stSidebar"] img { max-width: 150px; }

[data-testid="stSidebar"] .stButton button {
  width: 100%;
  background: transparent !important;
  color: var(--ink-800) !important;
  border: 1px solid transparent !important;
  border-radius: 14px !important;
  padding: 0.6rem 0.8rem !important;
  font-weight: 500 !important;
  font-size: 13px !important;
  letter-spacing: 0 !important;
  text-transform: none !important;
  text-align: left !important;
  transition: background .15s, border-color .15s !important;
}
[data-testid="stSidebar"] .stButton button:hover {
  background: rgba(255,255,255,0.7) !important;
}
[data-testid="stSidebar"] .stButton button:focus { box-shadow: none !important; }

[data-testid="stSidebar"] .nav-active .stButton button,
[data-testid="stSidebar"] .nav-active .stButton button:hover {
  background: var(--ink-900) !important;
  color: white !important;
  border-color: var(--ink-800) !important;
  box-shadow: 0 6px 20px -10px rgba(10,22,40,0.5) !important;
}

.sb-label {
  font-family: var(--font-mono);
  font-size: 10px;
  letter-spacing: 0.14em;
  text-transform: uppercase;
  color: var(--ink-500);
  padding: 12px 4px 6px;
}
.sb-status {
  display: flex; align-items: center; gap: 8px;
  font-family: var(--font-mono);
  font-size: 11px;
  color: var(--ink-700);
  padding: 8px 6px;
}
.sb-status .pulse {
  width: 7px; height: 7px; border-radius: 50%;
  background: #34A36B;
  box-shadow: 0 0 0 4px rgba(52,163,107,0.18);
}
.sb-status.off .pulse { background: var(--danger); box-shadow: 0 0 0 4px rgba(179,58,58,0.18); }

/* ===== Hero ===== */
.eyebrow {
  font-family: var(--font-mono); font-size: 10px;
  letter-spacing: 0.16em; text-transform: uppercase;
  color: var(--ink-500); display: inline-block;
}
.hero-title {
  font-size: 30px; font-weight: 600; letter-spacing: -0.022em;
  line-height: 1.08; color: var(--ink-900); margin: 6px 0 4px;
}
.hero-sub { color: var(--ink-700); font-size: 14px; max-width: 580px; line-height: 1.5;}

/* ===== Glass card ===== */
.glass {
  background: var(--glass-bg);
  backdrop-filter: blur(22px) saturate(120%);
  -webkit-backdrop-filter: blur(22px) saturate(120%);
  border: 1px solid var(--glass-edge);
  border-radius: 20px;
  box-shadow: var(--glass-shadow);
  padding: 22px;
  margin-bottom: 14px;
  position: relative;
}

/* ===== Stepper ===== */
.stepper {
  display: flex; align-items: center;
  padding: 8px;
  border-radius: 999px;
  background: rgba(255,255,255,0.6);
  border: 1px solid var(--line);
  width: fit-content;
  margin-bottom: 18px;
  flex-wrap: wrap;
  gap: 4px;
}
.stepper .step {
  display: flex; align-items: center; gap: 10px;
  padding: 6px 14px;
  border-radius: 999px;
  font-size: 12px;
  color: var(--ink-600);
  font-weight: 500;
}
.stepper .step .num {
  width: 22px; height: 22px;
  border-radius: 50%;
  background: var(--ink-100);
  display: inline-flex; align-items: center; justify-content: center;
  font-family: var(--font-mono); font-size: 11px; color: var(--ink-700);
}
.stepper .step.active { background: var(--ink-900); color: white; }
.stepper .step.active .num { background: var(--amber-500); color: var(--ink-900); }
.stepper .step.done .num { background: var(--success); color: white; }
.stepper .step.done { color: var(--ink-800); }

/* ===== KPI ===== */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 14px;
  margin-bottom: 14px;
}
.kpi {
  padding: 16px 18px;
  background: var(--glass-bg);
  border: 1px solid var(--glass-edge);
  border-radius: 18px;
  box-shadow: var(--glass-shadow);
  backdrop-filter: blur(22px);
  -webkit-backdrop-filter: blur(22px);
}
.kpi .label {
  font-family: var(--font-mono); font-size: 10px;
  text-transform: uppercase; letter-spacing: 0.14em;
  color: var(--ink-500);
}
.kpi .value {
  font-size: 28px; font-weight: 600; letter-spacing: -0.022em;
  font-variant-numeric: tabular-nums; color: var(--ink-900);
  margin-top: 4px; display: flex; align-items: baseline; gap: 6px;
}
.kpi .value .unit { font-size: 13px; color: var(--ink-500); font-weight: 500; }
.kpi .delta {
  font-size: 11px; font-family: var(--font-mono);
  display: inline-flex; align-items: center; gap: 4px;
  padding: 2px 6px; border-radius: 6px; margin-top: 6px;
}
.kpi .delta.up { color: var(--success); background: var(--success-bg); }
.kpi .delta.down { color: var(--danger); background: var(--danger-bg); }
.kpi .delta.flat { color: var(--ink-700); background: var(--ink-100); }

/* ===== Inputs ===== */
.stTextInput input, .stNumberInput input,
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div {
  background: rgba(255,255,255,0.85) !important;
  border-radius: 12px !important;
  border: 1px solid var(--line-2) !important;
  font-family: var(--font-sans) !important;
  font-size: 13px !important;
  color: var(--ink-900) !important;
}
.stTextArea textarea {
  background: rgba(255,255,255,0.85) !important;
  border-radius: 12px !important;
  border: 1px solid var(--line-2) !important;
  font-family: var(--font-sans) !important;
  font-size: 13px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
  border-color: var(--ink-700) !important;
  box-shadow: 0 0 0 3px rgba(31,52,97,0.12) !important;
}
[data-testid="stWidgetLabel"] p {
  font-size: 12px !important;
  font-weight: 500 !important;
  color: var(--ink-700) !important;
}

/* ===== File uploader ===== */
[data-testid="stFileUploader"] section {
  background: radial-gradient(400px 200px at 50% 0%, rgba(255,255,255,0.7), rgba(255,255,255,0.3)) !important;
  border-radius: 18px !important;
  border: 1.5px dashed var(--line-2) !important;
  padding: 28px 18px !important;
}
[data-testid="stFileUploader"] section:hover { border-color: var(--ink-700) !important; }
[data-testid="stFileUploader"] button {
  background: var(--ink-900) !important;
  color: white !important;
  border: none !important;
  border-radius: 12px !important;
  padding: 8px 16px !important;
  font-weight: 500 !important;
  font-family: var(--font-sans) !important;
  font-size: 13px !important;
}

/* ===== Buttons ===== */
.stButton button, .stDownloadButton button {
  background: rgba(255,255,255,0.7) !important;
  color: var(--ink-900) !important;
  border: 1px solid var(--line-2) !important;
  border-radius: 12px !important;
  padding: 8px 16px !important;
  font-weight: 500 !important;
  font-family: var(--font-sans) !important;
  font-size: 13px !important;
  letter-spacing: 0 !important;
  text-transform: none !important;
  transition: background .15s, border-color .15s !important;
}
.stButton button:hover, .stDownloadButton button:hover {
  background: white !important;
}
.stButton button[kind="primary"] {
  background: var(--ink-900) !important;
  color: white !important;
  border-color: var(--ink-900) !important;
  box-shadow: 0 1px 0 rgba(255,255,255,0.18) inset, 0 8px 22px -12px rgba(10,22,40,0.5) !important;
}
.stButton button[kind="primary"]:hover { background: var(--ink-800) !important; }
.stDownloadButton button {
  background: var(--amber-500) !important;
  color: var(--ink-900) !important;
  border-color: var(--amber-600) !important;
  font-weight: 600 !important;
  box-shadow: 0 1px 0 rgba(255,255,255,0.4) inset, 0 8px 22px -12px rgba(208,140,31,0.5) !important;
}
.stDownloadButton button:hover { background: var(--amber-400) !important; }

/* ===== Checkbox / Radio ===== */
.stCheckbox label, .stRadio label { font-size: 13px !important; color: var(--ink-800) !important; }

/* ===== Divider ===== */
hr, [data-testid="stDivider"] { border-color: var(--line) !important; opacity: 1; }

/* ===== Progress bar ===== */
.stProgress > div > div > div { background: var(--ink-900) !important; }
.stProgress > div > div { background: var(--ink-100) !important; }

/* ===== Expander ===== */
[data-testid="stExpander"] {
  background: rgba(255,255,255,0.5) !important;
  border: 1px solid var(--line) !important;
  border-radius: 14px !important;
  backdrop-filter: blur(8px);
}
[data-testid="stExpander"] summary { font-size: 13px !important; font-weight: 500 !important; }

/* ===== Dataframe ===== */
[data-testid="stDataFrame"] {
  border-radius: 14px;
  overflow: hidden;
  border: 1px solid var(--line);
}

/* ===== Alerts ===== */
[data-testid="stAlertContainer"] {
  border-radius: 14px !important;
  border: 1px solid var(--line) !important;
  font-size: 13px !important;
}

/* ===== Misc badges via spans ===== */
.badge {
  display: inline-flex; align-items: center; gap: 4px;
  padding: 2px 8px; border-radius: 999px;
  font-size: 11px; font-weight: 500;
  background: var(--ink-100); color: var(--ink-700);
  border: 1px solid var(--line);
}
.badge.amber { background: var(--amber-100); color: var(--amber-700); border-color: var(--amber-200); }
.badge.success { background: var(--success-bg); color: var(--success); border-color: rgba(47,125,92,0.18); }
.badge.outline { background: white; border-color: var(--line-2); color: var(--ink-700); }
.mono { font-family: var(--font-mono); }

/* footer fade */
footer { visibility: hidden; }
</style>
"""
st.markdown(DESIGN_CSS, unsafe_allow_html=True)


# =============================================================================
# BACKEND HELPERS  (preservados — assumindo os mesmos modulos do app original)
# =============================================================================
def _secret_to_env() -> None:
    if os.getenv("OPENAI_API_KEY"):
        return
    try:
        key = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        key = ""
    if key:
        os.environ["OPENAI_API_KEY"] = key


@st.cache_resource(show_spinner=False)
def _get_codificador():
    _secret_to_env()
    from codificador import CodificadorIA
    return CodificadorIA()


@st.cache_data(show_spinner=False)
def _read_uploaded_file(name: str, data: bytes) -> dict[str, pd.DataFrame]:
    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        return {"Planilha": pd.read_csv(bio)}
    xl = pd.ExcelFile(bio)
    return {sheet: xl.parse(sheet) for sheet in xl.sheet_names}


def _to_excel(sheets: dict[str, pd.DataFrame]) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31] or "Planilha"
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return out.getvalue()


@st.cache_data(show_spinner=False)
def _read_tabulation_file(name: str, data: bytes, two_line_header: bool) -> pd.DataFrame:
    bio = BytesIO(data)
    if name.lower().endswith(".csv"):
        return pd.read_csv(bio)
    if two_line_header:
        from tabulador import set_header
        raw = pd.read_excel(bio, header=None, sheet_name=0)
        return set_header(raw)
    return pd.read_excel(bio, sheet_name=0)


def _sanitize_export_df(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    if isinstance(clean.columns, pd.MultiIndex):
        clean.columns = ["_".join(str(c) for c in col).strip("_") for col in clean.columns]
    seen: dict[str, int] = {}
    cols = []
    for col in clean.columns:
        name = str(col)
        if name in seen:
            seen[name] += 1
            cols.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            cols.append(name)
    clean.columns = cols
    for col in clean.columns:
        if isinstance(clean[col], pd.DataFrame):
            clean[col] = clean[col].iloc[:, 0]
        try:
            clean[col] = clean[col].astype(object)
        except Exception:
            pass
    return clean


def _build_tab_excel(df, perguntas, titulo):
    from tabulador import exportar_excel
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = tmp.name
    try:
        exportar_excel(_sanitize_export_df(df), perguntas, saida=path,
                       titulo=titulo or "Pesquisa IFec RJ", total_respostas=len(df))
        return Path(path).read_bytes()
    finally:
        try: Path(path).unlink(missing_ok=True)
        except Exception: pass


def _build_tab_ppt(df, perguntas, titulo, subtitulo, periodo):
    from gerador_ppt import gerar_ppt
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        path = tmp.name
    try:
        gerar_ppt(_sanitize_export_df(df), perguntas, saida=path,
                  titulo=titulo or "Pesquisa IFec RJ",
                  subtitulo=subtitulo or "", periodo=periodo or "")
        return Path(path).read_bytes()
    finally:
        try: Path(path).unlink(missing_ok=True)
        except Exception: pass


def _parse_list(raw: str) -> list[str]:
    items = []
    for part in raw.replace("\n", ",").split(","):
        val = part.strip()
        if val and val not in items:
            items.append(val)
    return items


def _load_taxonomies():
    _secret_to_env()
    from codificador import MODOS_RESPOSTA, TIPOS_PERGUNTA
    return TIPOS_PERGUNTA, MODOS_RESPOSTA


def _run_coding(sheets, configs, global_context, previous_categories):
    codificador = _get_codificador()
    result_sheets = {name: df.copy() for name, df in sheets.items()}
    selected = [name for name, cfg in configs.items() if cfg["selected"]]
    progress = st.progress(0, text="Preparando codificacao...")

    for sheet_idx, sheet_name in enumerate(selected):
        cfg = configs[sheet_name]
        df = result_sheets[sheet_name]
        respostas = df[cfg["input_col"]].astype(str).tolist()
        cats_prev = previous_categories.get(sheet_name, [])
        cats_imp = cats_prev or cfg["categories"]
        context = cfg["context"] if cfg["type_key"] == "livre" else global_context

        def on_progress(i_local, total_local, resposta, categoria):
            pct = (sheet_idx + ((i_local + 1) / max(total_local, 1))) / max(len(selected), 1)
            progress.progress(min(pct, 1.0),
                              text=f"{sheet_name}: {i_local + 1}/{total_local}")

        coded = codificador.codificar_lote_modo(
            respostas, tipo=cfg["type_key"], modo=cfg["mode_key"],
            contexto_custom=context,
            categorias_imputacao=cats_imp,
            categorias_anteriores=cats_prev,
            callback_progresso=on_progress,
        )
        if "imputado" in coded:
            df[cfg["imputed_col"]] = coded["imputado"]
            df[cfg["new_col"]] = coded["novo"]
        else:
            df[cfg["output_col"]] = coded["resultado"]
        result_sheets[sheet_name] = df

    progress.progress(1.0, text="Codificacao concluida.")
    return result_sheets


# =============================================================================
# UI HELPERS
# =============================================================================
def _ss_get(k, default=None): return st.session_state.get(k, default)
def _ss_set(k, v): st.session_state[k] = v


def render_stepper(current: int) -> None:
    steps = [
        (1, "Importar base"),
        (2, "Configurar perguntas"),
        (3, "Codificar com IA"),
        (4, "Revisar resultado"),
    ]
    html = ['<div class="stepper">']
    for i, (n, label) in enumerate(steps):
        cls = "active" if n == current else ("done" if n < current else "")
        num = "✓" if n < current else str(n)
        html.append(
            f'<div class="step {cls}"><span class="num">{num}</span>{label}</div>'
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def render_hero(eyebrow: str, title: str, sub: str) -> None:
    st.markdown(
        f'<div class="eyebrow">{eyebrow}</div>'
        f'<div class="hero-title">{title}</div>'
        f'<div class="hero-sub">{sub}</div>',
        unsafe_allow_html=True,
    )


def render_kpis(items: list[dict]) -> None:
    """items: [{label, value, unit?, delta?, tone?}]"""
    cards = []
    for it in items:
        unit = f'<span class="unit">{it.get("unit","")}</span>' if it.get("unit") else ""
        delta = ""
        if it.get("delta"):
            tone = it.get("tone", "flat")
            arrow = "↑" if tone == "up" else "↓" if tone == "down" else "→"
            delta = f'<div class="delta {tone}">{arrow} {it["delta"]}</div>'
        cards.append(
            f'<div class="kpi">'
            f'<div class="label">{it["label"]}</div>'
            f'<div class="value">{it["value"]} {unit}</div>'
            f'{delta}'
            f'</div>'
        )
    st.markdown(f'<div class="kpi-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def glass_open(): st.markdown('<div class="glass">', unsafe_allow_html=True)
def glass_close(): st.markdown('</div>', unsafe_allow_html=True)


# =============================================================================
# CODIFICADOR — wizard
# =============================================================================
def _render_codificador(api_ok: bool) -> None:
    if not api_ok:
        st.warning("Configure OPENAI_API_KEY em Secrets para habilitar a codificacao.")

    step = _ss_get("cod_step", 1)
    if step == 4 and "result_sheets" in st.session_state:
        render_stepper(4)
        _step4_resultado()
        return

    render_stepper(step)
    if step == 1: _step1_upload()
    elif step == 2: _step2_perguntas()
    elif step == 3: _step3_executar()
    else:
        _ss_set("cod_step", 1); _step1_upload()


# ---- Etapa 1 -------------------------------------------------------------
def _step1_upload() -> None:
    col_main, col_side = st.columns([1.5, 1])

    with col_main:
        glass_open()
        st.markdown(
            '<div class="eyebrow">Etapa 1 · Base principal</div>'
            '<div style="font-size:18px; font-weight:600; margin: 4px 0 14px;">'
            'Importe a base que sera codificada</div>',
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Arraste a planilha ou clique para selecionar",
            type=["xlsx", "csv"],
            help="Aceitamos .xlsx e .csv ate 200 MB. Cada aba vira uma pergunta.",
            key="cod_upload_main",
            label_visibility="collapsed",
        )
        st.markdown(
            '<div style="font-size:12px; color:var(--ink-500); text-align:center; margin-top:6px;">'
            'Aceitamos <span class="mono">.xlsx</span> e <span class="mono">.csv</span> · ate 200 MB</div>',
            unsafe_allow_html=True,
        )

        if uploaded is not None:
            _ss_set("cod_uploaded_name", uploaded.name)
            _ss_set("cod_uploaded_bytes", uploaded.getvalue())
            st.markdown(
                f'<div style="margin-top:14px; padding:12px 14px; background:white; '
                f'border:1px solid var(--line); border-radius:12px; '
                f'display:flex; align-items:center; gap:12px;">'
                f'<span class="badge success">✓ Lido</span>'
                f'<span style="font-size:13px; font-weight:500;">{uploaded.name}</span>'
                f'<span style="margin-left:auto;" class="mono badge outline">'
                f'{len(uploaded.getvalue())/1024:.0f} KB</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown('<div style="height: 16px;"></div>', unsafe_allow_html=True)
        st.markdown('<hr style="margin: 10px 0 14px;">', unsafe_allow_html=True)

        recorrente = st.checkbox(
            "Pesquisa recorrente — usar categorias de uma rodada anterior",
            value=_ss_get("cod_recorrente", False),
            key="cod_recorrente",
        )
        previous_file = None
        if recorrente:
            previous_file = st.file_uploader(
                "Importe o dicionario da pesquisa anterior",
                type=["xlsx", "csv"],
                key="cod_upload_prev",
            )
            if previous_file is not None:
                _ss_set("cod_prev_name", previous_file.name)
                _ss_set("cod_prev_bytes", previous_file.getvalue())
        else:
            for k in ("cod_prev_name", "cod_prev_bytes", "cod_prev_categories"):
                st.session_state.pop(k, None)
        glass_close()

    with col_side:
        glass_open()
        st.markdown('<div class="eyebrow">O que esperar</div>', unsafe_allow_html=True)
        for n, t, d in [
            (1, "Importe", "Base principal e — opcionalmente — dicionario de pesquisa anterior."),
            (2, "Configure", "Para cada aba, indique a coluna de resposta, tipo e modo."),
            (3, "Codifique", "A IA categoriza com base no contexto e na biblioteca."),
            (4, "Refine e exporte", "Ajuste categorias, mescle duplicadas e baixe a planilha."),
        ]:
            st.markdown(
                f'<div style="display:flex; gap:12px; margin-top:12px;">'
                f'<div style="width:26px; height:26px; border-radius:8px; background:var(--ink-900); '
                f'color:white; display:flex; align-items:center; justify-content:center; '
                f'font-size:12px; font-weight:600; font-family:var(--font-mono); flex-shrink:0;">{n}</div>'
                f'<div><div style="font-size:13px; font-weight:600;">{t}</div>'
                f'<div style="font-size:12px; color:var(--ink-500); line-height:1.45;">{d}</div></div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        glass_close()

        glass_open()
        st.markdown('<div class="eyebrow">Modelo · GPT</div>', unsafe_allow_html=True)
        for k, v in [
            ("Custo estimado", "R$ 0,12 / 1k"),
            ("Tempo medio", "~40s / 100 resp."),
            ("Armazenamento", "local · IFec RJ"),
        ]:
            st.markdown(
                f'<div style="display:flex; justify-content:space-between; '
                f'align-items:center; padding:6px 0; border-top:1px solid var(--hairline); font-size:13px;">'
                f'<span>{k}</span><span class="mono" style="font-weight:600;">{v}</span></div>',
                unsafe_allow_html=True,
            )
        glass_close()

    # footer
    _, c2 = st.columns([3, 1])
    with c2:
        pode = "cod_uploaded_bytes" in st.session_state
        if st.button("Avancar →", type="primary", disabled=not pode,
                     use_container_width=True, key="cod_next_1"):
            _ss_set("cod_step", 2); _ss_set("cod_q_idx", 0); st.rerun()


# ---- Etapa 2 -------------------------------------------------------------
def _step2_perguntas() -> None:
    try:
        sheets = _read_uploaded_file(_ss_get("cod_uploaded_name"),
                                     _ss_get("cod_uploaded_bytes"))
    except Exception as exc:
        st.error(f"Nao foi possivel ler o arquivo: {exc}")
        if st.button("Voltar", key="back_err"):
            _ss_set("cod_step", 1); st.rerun()
        return

    previous_categories: dict[str, list[str]] = {}
    if _ss_get("cod_prev_bytes") is not None:
        with st.expander("Pesquisa anterior — confirme as colunas de categoria", expanded=False):
            prev_sheets = _read_uploaded_file(_ss_get("cod_prev_name"),
                                              _ss_get("cod_prev_bytes"))
            for prev_name, prev_df in prev_sheets.items():
                if prev_df.empty: continue
                default_col = prev_df.columns[-1]
                category_col = st.selectbox(
                    f"Coluna de categorias em '{prev_name}'",
                    list(prev_df.columns),
                    index=list(prev_df.columns).index(default_col),
                    key=f"prev_col_{prev_name}",
                )
                cats = (prev_df[category_col].dropna().astype(str).map(str.strip)
                        .loc[lambda s: s.ne("")].drop_duplicates().tolist())
                previous_categories[prev_name] = cats
        _ss_set("cod_prev_categories", previous_categories)
    else:
        previous_categories = _ss_get("cod_prev_categories", {}) or {}

    tipos_pergunta, modos_resposta = _load_taxonomies()
    type_labels = {data["label"]: key for key, data in tipos_pergunta.items()}
    mode_labels = {data["label"]: key for key, data in modos_resposta.items()}

    sheet_names = list(sheets.keys())
    total_q = len(sheet_names)
    q_idx = max(0, min(_ss_get("cod_q_idx", 0), total_q - 1))
    sheet_name = sheet_names[q_idx]
    df = sheets[sheet_name]

    # Layout de duas colunas: lista + detalhe
    col_list, col_detail = st.columns([1, 2.6])

    with col_list:
        glass_open()
        st.markdown(
            '<div style="display:flex; justify-content:space-between; align-items:center;">'
            '<span class="eyebrow">Perguntas</span>'
            f'<span class="mono" style="font-size:11px; color:var(--ink-500);">{q_idx+1}/{total_q}</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        for i, name in enumerate(sheet_names):
            cfg_n = _ss_get(f"cod_cfg_{name}")
            done = bool(cfg_n)
            active = i == q_idx
            bg = "var(--ink-900)" if active else "transparent"
            color = "white" if active else "var(--ink-900)"
            num_bg = "var(--amber-500)" if active else ("var(--success)" if done else "var(--ink-100)")
            num_color = "var(--ink-900)" if active else ("white" if done else "var(--ink-700)")
            label_col = "rgba(255,255,255,0.7)" if active else "var(--ink-500)"
            st.markdown(
                f'<div style="display:flex; gap:10px; padding:10px; border-radius:12px; '
                f'background:{bg}; color:{color}; margin-top:6px;">'
                f'<div style="width:22px; height:22px; border-radius:6px; background:{num_bg}; '
                f'color:{num_color}; display:flex; align-items:center; justify-content:center; '
                f'font-family:var(--font-mono); font-size:11px; font-weight:600; flex-shrink:0;">{"✓" if done and not active else i+1}</div>'
                f'<div><div style="font-size:12.5px; font-weight:500; line-height:1.3;">{name}</div>'
                f'<div style="font-size:11px; color:{label_col};">{len(sheets[name])} respostas</div></div></div>',
                unsafe_allow_html=True,
            )
        glass_close()

        glass_open()
        st.markdown('<span class="eyebrow">Contexto da pesquisa</span>', unsafe_allow_html=True)
        gctx = st.text_area(
            "Descreva o objetivo da pesquisa e os criterios",
            value=_ss_get("cod_global_context", ""),
            height=110,
            key="cod_global_context_input",
            label_visibility="collapsed",
        )
        _ss_set("cod_global_context", gctx)
        st.caption("Aplicado a todas as perguntas marcadas como “livre”.")
        glass_close()

    with col_detail:
        glass_open()
        cfg_prev = _ss_get(f"cod_cfg_{sheet_name}", {}) or {}
        cols = list(df.columns)

        # Header da pergunta
        nav_l, nav_c, nav_r = st.columns([1, 8, 1])
        with nav_l:
            if st.button("‹", key="q_prev", disabled=q_idx == 0, use_container_width=True):
                _ss_set("cod_q_idx", q_idx - 1); st.rerun()
        with nav_c:
            st.markdown(
                f'<div style="text-align:center;">'
                f'<span class="badge amber mono">P{q_idx+1} · {sheet_name}</span>'
                f'<div style="font-size:18px; font-weight:600; margin-top:6px;">{sheet_name}</div>'
                f'<div style="font-size:12px; color:var(--ink-500);">{len(df)} respostas · {len(cols)} colunas</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        with nav_r:
            if st.button("›", key="q_next", disabled=q_idx >= total_q - 1, use_container_width=True):
                _ss_set("cod_q_idx", q_idx + 1); st.rerun()

        st.markdown('<div style="height: 12px;"></div>', unsafe_allow_html=True)

        selected = st.checkbox(
            "Codificar esta aba",
            value=cfg_prev.get("selected", True),
            key=f"sel_{sheet_name}",
        )

        c1, c2 = st.columns(2)
        with c1:
            input_col = st.selectbox(
                "Coluna de resposta",
                cols,
                index=cols.index(cfg_prev["input_col"]) if cfg_prev.get("input_col") in cols else 0,
                key=f"in_{sheet_name}",
            )
        with c2:
            output_col = st.text_input(
                "Coluna de saida",
                value=cfg_prev.get("output_col", "codigo_ia"),
                key=f"out_{sheet_name}",
            )

        c3, c4 = st.columns(2)
        with c3:
            type_options = list(type_labels.keys())
            default_type_label = next(k for k, v in type_labels.items() if v == "livre")
            type_label = st.selectbox(
                "Tipo de pergunta", type_options,
                index=type_options.index(cfg_prev.get("type_label", default_type_label))
                      if cfg_prev.get("type_label", default_type_label) in type_options
                      else type_options.index(default_type_label),
                key=f"type_{sheet_name}",
            )
        with c4:
            mode_options = list(mode_labels.keys())
            mode_pref = cfg_prev.get("mode_label", mode_options[0])
            mode_label = st.selectbox(
                "Modo de resposta", mode_options,
                index=mode_options.index(mode_pref) if mode_pref in mode_options else 0,
                key=f"mode_{sheet_name}",
            )
        mode_key = mode_labels[mode_label]

        aba_anterior = ""
        if previous_categories:
            opcoes_aba_ant = ["(nenhuma)"] + list(previous_categories.keys())
            aba_pref = cfg_prev.get("aba_anterior", "(nenhuma)")
            idx_default = opcoes_aba_ant.index(aba_pref) if aba_pref in opcoes_aba_ant else 0
            aba_anterior = st.selectbox(
                "Aba da pesquisa anterior", opcoes_aba_ant,
                index=idx_default, key=f"aba_ant_{sheet_name}",
            )

        imputed_col = cfg_prev.get("imputed_col", "col_imputado")
        new_col = cfg_prev.get("new_col", "col_nova")
        categories: list[str] = cfg_prev.get("categories", [])
        if "semi" in mode_key:
            sa, sb = st.columns(2)
            with sa: imputed_col = st.text_input("Coluna de imputacao", value=imputed_col, key=f"imp_{sheet_name}")
            with sb: new_col = st.text_input("Coluna de novas categorias", value=new_col, key=f"new_{sheet_name}")
            categories = _parse_list(st.text_area(
                "Categorias pre-definidas",
                value=", ".join(categories),
                placeholder="Uma categoria por linha ou separadas por virgula.",
                key=f"cats_{sheet_name}",
            ))

        custom_context = st.text_area(
            "Contexto especifico desta pergunta",
            value=cfg_prev.get("context", ""),
            placeholder="Use quando o tipo for Personalizado.",
            key=f"ctx_{sheet_name}", height=70,
        )

        with st.expander("Previa das primeiras respostas"):
            st.dataframe(df[[input_col]].head(10), use_container_width=True, hide_index=True)

        glass_close()

        # Salva config
        _ss_set(f"cod_cfg_{sheet_name}", {
            "selected": selected, "input_col": input_col,
            "output_col": output_col.strip() or "codigo_ia",
            "type_label": type_label, "type_key": type_labels[type_label],
            "mode_label": mode_label, "mode_key": mode_key,
            "categories": categories,
            "imputed_col": imputed_col.strip() or "col_imputado",
            "new_col": new_col.strip() or "col_nova",
            "context": custom_context.strip(),
            "aba_anterior": aba_anterior,
        })

    # footer
    bt1, _, bt3 = st.columns([1, 4, 1])
    with bt1:
        if st.button("← Voltar", key="cod_back_2", use_container_width=True):
            _ss_set("cod_step", 1); st.rerun()
    with bt3:
        ok = all(_ss_get(f"cod_cfg_{n}") is not None for n in sheet_names)
        if st.button("Iniciar codificacao →", type="primary",
                     disabled=not ok, use_container_width=True, key="cod_next_2"):
            _ss_set("cod_step", 3); st.rerun()


# ---- Etapa 3 -------------------------------------------------------------
def _step3_executar() -> None:
    glass_open()
    st.markdown(
        '<div class="eyebrow">Etapa 3 · Execucao</div>'
        '<div style="font-size:20px; font-weight:600; margin:6px 0 14px;">'
        'Codificando respostas com IA</div>',
        unsafe_allow_html=True,
    )

    try:
        sheets = _read_uploaded_file(_ss_get("cod_uploaded_name"),
                                     _ss_get("cod_uploaded_bytes"))
    except Exception as exc:
        st.error(f"Erro ao reler o arquivo: {exc}")
        glass_close()
        return

    configs, previous_for_run = {}, {}
    previous_categories = _ss_get("cod_prev_categories", {}) or {}
    for sheet_name in sheets.keys():
        cfg = _ss_get(f"cod_cfg_{sheet_name}")
        if not cfg: continue
        configs[sheet_name] = {
            "selected": cfg["selected"], "input_col": cfg["input_col"],
            "output_col": cfg["output_col"], "type_key": cfg["type_key"],
            "mode_key": cfg["mode_key"], "categories": cfg["categories"],
            "imputed_col": cfg["imputed_col"], "new_col": cfg["new_col"],
            "context": cfg["context"],
        }
        aba_ant = cfg.get("aba_anterior", "")
        if aba_ant and aba_ant != "(nenhuma)" and aba_ant in previous_categories:
            previous_for_run[sheet_name] = previous_categories[aba_ant]

    global_context = _ss_get("cod_global_context", "")

    with st.spinner("Codificando respostas..."):
        try:
            result_sheets = _run_coding(sheets, configs, global_context.strip(), previous_for_run)
        except Exception as exc:
            st.error(f"Erro durante a codificacao: {exc}")
            glass_close()
            if st.button("Voltar para configuracao", key="back_to_cfg"):
                _ss_set("cod_step", 2); st.rerun()
            return

    _ss_set("result_sheets", result_sheets)
    st.success("Codificacao finalizada com sucesso!")
    glass_close()

    if st.button("Ver resultado →", type="primary", use_container_width=True, key="cod_next_3"):
        _ss_set("cod_step", 4); st.rerun()


# ---- Etapa 4 -------------------------------------------------------------
def _step4_resultado() -> None:
    result_sheets: dict[str, pd.DataFrame] = _ss_get("result_sheets", {})
    if not result_sheets:
        st.info("Nenhum resultado disponivel ainda.")
        return

    # KPIs gerais
    total_resp = sum(len(df) for df in result_sheets.values())
    total_cats = 0
    for name, df in result_sheets.items():
        cfg = _ss_get(f"cod_cfg_{name}", {}) or {}
        out_col = cfg.get("output_col", "codigo_ia")
        if out_col in df.columns:
            cats = df[out_col].dropna().astype(str).str.split(";").explode().map(str.strip)
            total_cats += cats[cats.ne("")].nunique()

    render_kpis([
        {"label": "Respostas codificadas", "value": f"{total_resp:,}".replace(",", "."), "delta": "100% de cobertura", "tone": "up"},
        {"label": "Categorias finais", "value": str(total_cats), "delta": "consolidadas", "tone": "flat"},
        {"label": "Abas processadas", "value": str(len(result_sheets)), "delta": f"{len(result_sheets)} de {len(result_sheets)}", "tone": "up"},
        {"label": "Status", "value": "✓", "delta": "concluido", "tone": "up"},
    ])

    glass_open()
    st.markdown(
        '<div class="eyebrow">Frequencia de categorias</div>'
        '<div style="font-size:18px; font-weight:600; margin:4px 0 12px;">'
        'Selecione uma aba e explore as respostas por categoria</div>',
        unsafe_allow_html=True,
    )
    preview_sheet = st.selectbox(
        "Aba para analisar", list(result_sheets.keys()),
        key="preview_sheet", label_visibility="collapsed",
    )
    df_res = result_sheets[preview_sheet]
    cfg = _ss_get(f"cod_cfg_{preview_sheet}", {}) or {}
    out_col = cfg.get("output_col", "codigo_ia")
    if out_col not in df_res.columns:
        for c in [cfg.get("imputed_col"), cfg.get("new_col")]:
            if c and c in df_res.columns:
                out_col = c; break

    freq = pd.Series(dtype=int)
    if out_col in df_res.columns:
        s = df_res[out_col].dropna().astype(str).map(str.strip)
        s = s.str.split(";").explode().map(str.strip)
        s = s[s.ne("")]
        freq = s.value_counts()

    left, right = st.columns(2)
    cat_selecionada = None
    with left:
        st.markdown(
            '<div style="font-size:13px; font-weight:600; margin-bottom:8px;">Categorias</div>',
            unsafe_allow_html=True,
        )
        if not freq.empty:
            cat_df = freq.reset_index()
            cat_df.columns = ["Categoria", "Respostas"]
            cat_df["%"] = (cat_df["Respostas"] / cat_df["Respostas"].sum() * 100).round(1)
            cat_selecionada = st.radio(
                "Categoria",
                cat_df["Categoria"].tolist(),
                key=f"cat_sel_{preview_sheet}",
                label_visibility="collapsed",
            )
            st.dataframe(cat_df, use_container_width=True, hide_index=True, height=320)
        else:
            st.caption("Nenhuma categoria detectada.")

    with right:
        st.markdown(
            f'<div style="font-size:13px; font-weight:600; margin-bottom:8px;">'
            f'Respostas{" · " + str(cat_selecionada) if cat_selecionada else ""}</div>',
            unsafe_allow_html=True,
        )
        if cat_selecionada and out_col in df_res.columns:
            input_col = cfg.get("input_col")
            cols_show = [c for c in [input_col, out_col] if c and c in df_res.columns]
            mask = df_res[out_col].astype(str).str.contains(
                str(cat_selecionada), case=False, na=False, regex=False)
            st.dataframe(
                df_res.loc[mask, cols_show].head(200),
                use_container_width=True, hide_index=True, height=320,
            )
        else:
            st.caption("Selecione uma categoria a esquerda.")
    glass_close()

    # Export bar
    glass_open()
    st.markdown(
        '<div class="eyebrow">Exportacao</div>'
        '<div style="font-size:18px; font-weight:600; margin:4px 0 4px;">'
        'Pronto para baixar a planilha codificada</div>'
        '<div style="font-size:13px; color:var(--ink-700); margin-bottom:14px;">'
        'Inclui as colunas <span class="mono">codigo_ia</span> em todas as abas codificadas.'
        '</div>',
        unsafe_allow_html=True,
    )
    col_d1, col_d2, col_d3 = st.columns([1,1,1])
    with col_d3:
        st.download_button(
            "↓ Baixar planilha codificada",
            data=_to_excel(result_sheets),
            file_name="base_codificada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    glass_close()

    a, b = st.columns(2)
    with a:
        if st.button("← Reconfigurar perguntas", use_container_width=True, key="reconfig_btn"):
            _ss_set("cod_step", 2); st.rerun()
    with b:
        if st.button("Iniciar nova codificacao", use_container_width=True, key="reset_btn"):
            for k in list(st.session_state.keys()):
                if k.startswith("cod_") or k == "result_sheets":
                    st.session_state.pop(k, None)
            _ss_set("cod_step", 1); st.rerun()


# =============================================================================
# TABULADOR
# =============================================================================
def _render_tabulador() -> None:
    fonte_options = ["Arquivo enviado"]
    if "result_sheets" in st.session_state:
        fonte_options.insert(0, "Resultado codificado")
    fonte = st.radio("Fonte da base", fonte_options, horizontal=True)

    df_tab: pd.DataFrame | None = None
    source_name = ""

    if fonte == "Resultado codificado":
        sheet_name = st.selectbox(
            "Aba codificada para tabular",
            list(st.session_state["result_sheets"].keys()),
            key="tab_result_sheet",
        )
        df_tab = st.session_state["result_sheets"][sheet_name].copy()
        source_name = sheet_name
    else:
        upload = st.file_uploader("Importe a base", type=["xlsx", "csv"], key="tab_upload")
        if upload is None:
            st.info("Envie uma planilha para iniciar a tabulacao.")
            return
        two = st.checkbox("Cabecalho em duas linhas (SurveyMonkey/TabIFec)", value=True)
        try:
            df_tab = _read_tabulation_file(upload.name, upload.getvalue(), two)
        except Exception as exc:
            st.error(f"Nao foi possivel preparar a base: {exc}")
            return
        source_name = Path(upload.name).stem

    render_kpis([
        {"label": "Base", "value": source_name[:18] or "—"},
        {"label": "Respondentes", "value": f"{len(df_tab):,}".replace(",", ".")},
        {"label": "Colunas", "value": str(len(df_tab.columns))},
        {"label": "Status", "value": "pronto", "delta": "para detectar", "tone": "flat"},
    ])

    if st.button("Detectar perguntas", type="primary", key="tab_detect"):
        from tabulador import detectar_perguntas
        try:
            st.session_state["tab_questions"] = detectar_perguntas(df_tab)
            st.session_state["tab_source_name"] = source_name
        except Exception as exc:
            st.error(f"Erro ao detectar perguntas: {exc}")

    perguntas_detectadas = st.session_state.get("tab_questions", [])
    if not perguntas_detectadas:
        with st.expander("Previa da base", expanded=True):
            st.dataframe(df_tab.head(30), use_container_width=True, hide_index=True)
        return

    st.markdown(
        f'<span class="badge success">✓ {len(perguntas_detectadas)} pergunta(s) detectada(s)</span>',
        unsafe_allow_html=True,
    )

    glass_open()
    st.markdown('<div class="eyebrow">Metadados do relatorio</div>', unsafe_allow_html=True)
    titulo = st.text_input("Titulo do relatorio", value="Pesquisa IFec RJ", key="tab_title")
    a, b = st.columns(2)
    with a: subtitulo = st.text_input("Subtitulo", value="", key="tab_subtitle")
    with b: periodo = st.text_input("Periodo", value="", key="tab_period")
    glass_close()

    from tabulador import TIPOS_LABEL, tabular_pergunta
    tipo_keys = list(TIPOS_LABEL.keys())
    tipo_labels = [f"{key} · {TIPOS_LABEL[key]}" for key in tipo_keys]
    perguntas_config: list[dict] = []

    for idx, p in enumerate(perguntas_detectadas):
        label = f"{p.get('num', f'P{idx+1:02d}')} · {p.get('pergunta', '')}"
        with st.expander(label, expanded=idx < 2):
            top_a, top_b, top_c = st.columns([1, 2, 4])
            ativo = top_a.checkbox("Ativa", value=p.get("ativo", True), key=f"tab_active_{idx}")
            tipo_atual = p.get("tipo", "ABERTA")
            tipo_index = tipo_keys.index(tipo_atual) if tipo_atual in tipo_keys else tipo_keys.index("ABERTA")
            tipo_label = top_b.selectbox("Tipo", tipo_labels, index=tipo_index, key=f"tab_type_{idx}")
            texto = top_c.text_input("Pergunta", value=str(p.get("pergunta", "")), key=f"tab_question_{idx}")
            nota = st.text_area("Nota", value=str(p.get("nota", "")), height=60, key=f"tab_note_{idx}")
            cfg = dict(p)
            cfg["ativo"] = ativo
            cfg["tipo"] = tipo_keys[tipo_labels.index(tipo_label)]
            cfg["pergunta"] = texto.strip() or p.get("pergunta", "")
            cfg["nota"] = nota.strip()
            perguntas_config.append(cfg)
            colunas = cfg.get("colunas", [])
            st.caption("Colunas: " + ", ".join(str(c) for c in colunas[:6]) + ("..." if len(colunas) > 6 else ""))
            if st.checkbox("Previa da tabulacao", key=f"tab_preview_{idx}"):
                try:
                    st.dataframe(tabular_pergunta(df_tab, cfg), use_container_width=True, hide_index=True)
                except Exception as exc:
                    st.warning(f"Nao foi possivel tabular esta pergunta: {exc}")

    ativas = [p for p in perguntas_config if p.get("ativo") and p.get("tipo") != "IGNORAR"]

    glass_open()
    st.markdown(
        f'<div class="eyebrow">Exportacao</div>'
        f'<div style="font-size:14px; margin-top:4px;">'
        f'<strong>{len(ativas)}</strong> pergunta(s) ativa(s) para exportacao</div>',
        unsafe_allow_html=True,
    )
    gx, gp = st.columns(2)
    with gx:
        if st.button("Gerar Excel de tabulacao", disabled=not ativas, use_container_width=True, key="tab_gen_xlsx"):
            with st.spinner("Gerando Excel..."):
                try:
                    st.session_state["tab_excel_bytes"] = _build_tab_excel(df_tab, ativas, titulo)
                except Exception as exc:
                    st.error(f"Erro ao gerar Excel: {exc}")
        if "tab_excel_bytes" in st.session_state:
            st.download_button("↓ Baixar Excel",
                data=st.session_state["tab_excel_bytes"],
                file_name=f"Tabulacao_{source_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
    with gp:
        if st.button("Gerar PowerPoint", disabled=not ativas, use_container_width=True, key="tab_gen_ppt"):
            with st.spinner("Gerando PowerPoint..."):
                try:
                    st.session_state["tab_ppt_bytes"] = _build_tab_ppt(df_tab, ativas, titulo, subtitulo, periodo)
                except Exception as exc:
                    st.error(f"Erro ao gerar PowerPoint: {exc}")
        if "tab_ppt_bytes" in st.session_state:
            st.download_button("↓ Baixar PowerPoint",
                data=st.session_state["tab_ppt_bytes"],
                file_name=f"Apresentacao_{source_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
    glass_close()


# =============================================================================
# MAIN
# =============================================================================
def main() -> None:
    _secret_to_env()
    if "active_view" not in st.session_state:
        st.session_state["active_view"] = "codificador"

    # SIDEBAR
    with st.sidebar:
        try:
            st.image("logo_ifec_header.png", use_container_width=True)
        except Exception:
            st.markdown("### IFec RJ")
        st.markdown(
            '<div class="mono" style="font-size:10px; letter-spacing:0.14em; '
            'text-transform:uppercase; color:var(--ink-500); margin-top:8px;">'
            'Plataforma de pesquisas</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="sb-label">Ferramentas</div>', unsafe_allow_html=True)
        is_cod = st.session_state["active_view"] == "codificador"
        st.markdown(f'<div class="{"nav-active" if is_cod else ""}">', unsafe_allow_html=True)
        if st.button("Codificador IA", key="nav_cod"):
            st.session_state["active_view"] = "codificador"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        is_tab = st.session_state["active_view"] == "tabulador"
        st.markdown(f'<div class="{"nav-active" if is_tab else ""}">', unsafe_allow_html=True)
        if st.button("Tabulacao automatica", key="nav_tab"):
            st.session_state["active_view"] = "tabulador"; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<div style='flex:1; min-height: 4rem;'></div>", unsafe_allow_html=True)

        api_ok = bool(os.getenv("OPENAI_API_KEY"))
        try:
            if not api_ok:
                key = st.secrets.get("OPENAI_API_KEY", "")
                if key:
                    os.environ["OPENAI_API_KEY"] = key
                    api_ok = True
        except Exception:
            pass

        cls = "" if api_ok else "off"
        st.markdown(
            f'<div class="sb-status {cls}"><span class="pulse"></span>'
            f'OPENAI · {"CONECTADA" if api_ok else "DESCONECTADA"}</div>',
            unsafe_allow_html=True,
        )

    # AREA CENTRAL
    if st.session_state["active_view"] == "codificador":
        render_hero("Codificador · IA",
                    "Codifique respostas abertas em minutos.",
                    "Importe a base, configure as perguntas e deixe a IA propor categorias. Voce revisa, ajusta e exporta — tudo em um unico fluxo.")
        st.markdown('<div style="height: 16px;"></div>', unsafe_allow_html=True)
        _render_codificador(api_ok=bool(os.getenv("OPENAI_API_KEY")))
    else:
        render_hero("Tabulacao · automatica",
                    "Da base codificada para o relatorio, sem planilha intermediaria.",
                    "Detectamos perguntas, agrupamos multipla escolha e geramos PowerPoint pronto para entrega.")
        st.markdown('<div style="height: 16px;"></div>', unsafe_allow_html=True)
        _render_tabulador()


if __name__ == "__main__":
    main()
