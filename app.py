from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

try:
    from streamlit_echarts import st_echarts
except Exception:
    st_echarts = None


DATA_PATH = Path("Data/ExcelPAEsintetizado.xlsx")
LOCAL_DATA_PATH = Path(r"C:\Users\marqu\OneDrive\Escritorio\Proyecto_PAE\Data\ExcelPAEsintetizado.xlsx")

st.set_page_config(
    page_title="PAE Santa Marta",
    page_icon="PAE",
    layout="wide",
    initial_sidebar_state="expanded",
)

CSS = """
<style>
.stApp {
  background:
    linear-gradient(180deg, #edf4f7 0%, #f7fafc 38%, #ffffff 100%);
}
.main .block-container {
  max-width: 1420px;
  padding-top: 1.1rem;
}
section[data-testid="stSidebar"] {
  background: #111d2e;
}
section[data-testid="stSidebar"] * {
  color: #f7fbff;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] * {
  color: #111d2e !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
  background: #ffffff !important;
  border-color: #d7e0ea !important;
}
section[data-testid="stSidebar"] div[data-baseweb="tag"] {
  background: #e8f2fb !important;
}
section[data-testid="stSidebar"] div[data-baseweb="tag"] * {
  color: #111d2e !important;
}
section[data-testid="stSidebar"] button {
  border-radius: 6px !important;
}
.sidebar-title {
  padding: 10px 0 14px 0;
  border-bottom: 1px solid rgba(255,255,255,.14);
  margin-bottom: 14px;
}
.sidebar-title .label {
  color: #8bd9d4;
  font-size: .72rem;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: .08em;
}
.sidebar-title .title {
  color: white;
  font-size: 1.15rem;
  font-weight: 850;
  margin-top: 4px;
}
.sidebar-help {
  color: rgba(255,255,255,.72);
  font-size: .82rem;
  margin: -4px 0 10px 0;
}
.hero {
  background:
    radial-gradient(circle at 87% 15%, rgba(255,255,255,.18), transparent 24%),
    linear-gradient(115deg, #111d2e 0%, #1f5f8b 58%, #00a19a 100%);
  color: white;
  padding: 34px 36px;
  border-radius: 8px;
  margin-bottom: 16px;
  box-shadow: 0 18px 42px rgba(17,29,46,.16);
}
.hero h1 {
  margin: 0 0 8px 0;
  font-size: 2.35rem;
  line-height: 1.08;
}
.hero p {
  margin: 0;
  max-width: 920px;
  color: rgba(255,255,255,.84);
  font-size: 1.08rem;
  font-weight: 600;
}
.source {
  display: inline-block;
  margin-top: 14px;
  padding: 6px 10px;
  background: rgba(255,255,255,.12);
  border: 1px solid rgba(255,255,255,.22);
  border-radius: 999px;
  font-size: .82rem;
}
.kpis {
  display: grid;
  grid-template-columns: repeat(5, minmax(150px, 1fr));
  gap: 12px;
  margin: 14px 0 18px 0;
}
.kpi {
  background: white;
  border: 1px solid #dce3ec;
  border-radius: 8px;
  padding: 15px 16px;
  box-shadow: 0 10px 26px rgba(17,29,46,.07);
}
.kpi:hover {
  transform: translateY(-1px);
  box-shadow: 0 14px 32px rgba(17,29,46,.10);
}
.kpi .label {
  color: #667085;
  font-size: .74rem;
  text-transform: uppercase;
  font-weight: 800;
}
.kpi .value {
  color: #142033;
  font-size: 1.62rem;
  font-weight: 850;
  margin-top: 4px;
}
.kpi .note {
  color: #667085;
  font-size: .8rem;
  margin-top: 2px;
}
.question {
  color: #142033;
  font-weight: 850;
  font-size: 1.13rem;
  margin: 20px 0 8px 0;
}
.insight {
  background: white;
  border-left: 5px solid #1f5f8b;
  border-radius: 8px;
  padding: 13px 15px;
  margin: 6px 0 14px 0;
  box-shadow: 0 10px 26px rgba(17,29,46,.06);
  color: #263243;
}
.guide {
  display: grid;
  grid-template-columns: repeat(3, minmax(180px, 1fr));
  gap: 12px;
  margin: 10px 0 16px 0;
}
.guide-card {
  background: #ffffff;
  border: 1px solid #dce3ec;
  border-radius: 8px;
  padding: 13px 15px;
  box-shadow: 0 8px 22px rgba(17,29,46,.055);
}
.guide-card strong {
  display: block;
  color: #142033;
  margin-bottom: 4px;
}
.guide-card span {
  color: #667085;
  font-size: .87rem;
}
.warning {
  border-left-color: #df5c5c;
}
.ok {
  border-left-color: #00a19a;
}
.small {
  color: #667085;
  font-size: .86rem;
}
@media (max-width: 1000px) {
  .kpis { grid-template-columns: repeat(2, minmax(150px, 1fr)); }
  .guide { grid-template-columns: 1fr; }
  .hero h1 { font-size: 1.65rem; }
}
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)

COLORS = ["#1f5f8b", "#00a19a", "#f4b942", "#df5c5c", "#6d5bd0", "#536878"]


def n(value) -> str:
    if pd.isna(value):
        return "-"
    return f"{float(value):,.0f}".replace(",", ".")


def money(value) -> str:
    if pd.isna(value):
        return "-"
    return "$" + f"{float(value):,.0f}".replace(",", ".")


def pct(value) -> str:
    if pd.isna(value):
        return "-"
    return f"{float(value) * 100:.1f}%".replace(".", ",")


def clean_name(value) -> str:
    return " ".join(str(value).strip().split())


def chart(options: dict, height: str = "430px", key: str | None = None):
    if st_echarts is None:
        st.info("Falta instalar streamlit-echarts. Ejecuta el lanzador que te dejé para usar el entorno correcto.")
        return
    st_echarts(options=options, height=height, key=key)


def grid() -> dict:
    return {"left": 16, "right": 24, "top": 44, "bottom": 34, "containLabel": True}


@st.cache_data(show_spinner=False)
def load_data(path_or_file):
    raw = pd.read_excel(path_or_file, sheet_name=None)
    sheet_names = list(raw.keys())
    inversion_sheet = next((s for s in sheet_names if "General24-25" in s), sheet_names[2])

    inv_raw = raw[inversion_sheet]
    globales = pd.DataFrame(
        {
            "anio": pd.to_numeric(inv_raw.iloc[:, 0], errors="coerce").astype(int),
            "inversion": pd.to_numeric(inv_raw.iloc[:, 1], errors="coerce"),
            "aumento": pd.to_numeric(inv_raw.iloc[:, 2], errors="coerce"),
            "beneficiarios": pd.to_numeric(inv_raw.iloc[:, 3], errors="coerce"),
            "matriculados": pd.to_numeric(inv_raw.iloc[:, 4], errors="coerce"),
            "cobertura_global": pd.to_numeric(inv_raw.iloc[:, 5], errors="coerce"),
        }
    )

    cobertura_frames = []
    for sheet in ["CoberturaxIED2024", "%CoberturaxIED2025"]:
        df = raw[sheet]
        tmp = pd.DataFrame(
            {
                "anio": pd.to_numeric(df.iloc[:, 0], errors="coerce").astype(int),
                "institucion": df.iloc[:, 1].map(clean_name),
                "matriculados": pd.to_numeric(df.iloc[:, 2], errors="coerce"),
                "beneficiarios": pd.to_numeric(df.iloc[:, 3], errors="coerce"),
                "cobertura": pd.to_numeric(df.iloc[:, 4], errors="coerce"),
                "almuerzo_regular": pd.to_numeric(df.iloc[:, 5], errors="coerce").fillna(0),
                "complemento_am": pd.to_numeric(df.iloc[:, 6], errors="coerce").fillna(0),
                "complemento_pm": pd.to_numeric(df.iloc[:, 7], errors="coerce").fillna(0),
            }
        )
        cobertura_frames.append(tmp)

    cobertura = pd.concat(cobertura_frames, ignore_index=True)
    total_cobertura = cobertura[cobertura["institucion"].str.upper() == "TOTAL GENERAL"].copy()
    instituciones = cobertura[cobertura["institucion"].str.upper() != "TOTAL GENERAL"].copy()
    instituciones["brecha"] = instituciones["matriculados"] - instituciones["beneficiarios"]
    instituciones["brecha"] = instituciones["brecha"].where(instituciones["brecha"] > 0, 0)
    instituciones["estado"] = pd.cut(
        instituciones["cobertura"].fillna(-1),
        bins=[-2, 0.5, 0.7, 0.85, 1.0, 99],
        labels=["Critico", "Bajo", "Medio", "Alto", "Sobre cobertura"],
    )

    gen_raw = raw["Genero"]
    genero = pd.DataFrame(
        {
            "anio": pd.to_numeric(gen_raw.iloc[:, 0], errors="coerce").astype(int),
            "genero": gen_raw.iloc[:, 1].astype(str).str.strip().str.title().replace({"Fememino": "Femenino"}),
            "cantidad": pd.to_numeric(gen_raw.iloc[:, 2], errors="coerce"),
            "participacion": pd.to_numeric(gen_raw.iloc[:, 3], errors="coerce"),
        }
    )

    grado_raw = raw["Accesoxgrado"]
    grado = pd.DataFrame(
        {
            "anio": pd.to_numeric(grado_raw.iloc[:, 0], errors="coerce").astype(int),
            "grado": grado_raw.iloc[:, 1].astype(str).str.strip().str.replace(".0", "", regex=False),
            "beneficiarios": pd.to_numeric(grado_raw.iloc[:, 2], errors="coerce"),
            "participacion": pd.to_numeric(grado_raw.iloc[:, 3], errors="coerce"),
        }
    )
    grado["orden"] = grado["grado"].map(lambda x: int(x) if str(x).isdigit() else 99)

    return {
        "globales": globales,
        "instituciones": instituciones,
        "total_cobertura": total_cobertura,
        "genero": genero,
        "grado": grado,
        "source_global": inversion_sheet,
    }


st.sidebar.markdown(
    """
    <div class="sidebar-title">
      <div class="label">Panel de control</div>
      <div class="title">Filtros del observatorio</div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    '<div class="sidebar-help">Usa los filtros para pasar de la lectura distrital al detalle institucional.</div>',
    unsafe_allow_html=True,
)

uploaded = st.sidebar.file_uploader("Cargar base de datos", type=["xlsx"])
source = uploaded if uploaded is not None else (DATA_PATH if DATA_PATH.exists() else LOCAL_DATA_PATH)

if uploaded is None and not source.exists():
    st.error("No encontré la base de datos. Cárgala desde la barra lateral.")
    st.stop()

data = load_data(source)
globales = data["globales"]
instituciones = data["instituciones"]
total_cobertura = data["total_cobertura"]
genero = data["genero"]
grado = data["grado"]

years = sorted(globales["anio"].unique().tolist())
year = st.sidebar.selectbox("Año", years, index=len(years) - 1)

estado_opts = instituciones["estado"].dropna().astype(str).unique().tolist()
selected_estado = st.sidebar.multiselect("Estado de cobertura", estado_opts, default=estado_opts)
inst_opts = sorted(instituciones["institucion"].dropna().unique().tolist())

if "selected_inst" not in st.session_state:
    st.session_state.selected_inst = inst_opts.copy()
else:
    st.session_state.selected_inst = [
        item for item in st.session_state.selected_inst if item in inst_opts
    ]

col_clear_all, col_select_all = st.sidebar.columns(2)
if col_clear_all.button("Eliminar todas", use_container_width=True):
    st.session_state.selected_inst = []
if col_select_all.button("Seleccionar todas", use_container_width=True):
    st.session_state.selected_inst = inst_opts.copy()

selected_inst = st.sidebar.multiselect("Institución", inst_opts, key="selected_inst")

global_selected = globales[globales["anio"] == year]
if global_selected.empty:
    st.error("No hay indicadores globales para el año seleccionado.")
    st.stop()
global_year = global_selected.iloc[0]
inst_year_all = instituciones[instituciones["anio"] == year].copy()
inst_year = inst_year_all[
    inst_year_all["estado"].astype(str).isin(selected_estado)
    & inst_year_all["institucion"].isin(selected_inst)
].copy()

total_cov_selected = total_cobertura[total_cobertura["anio"] == year]
if total_cov_selected.empty:
    total_cov_year = pd.Series(
        {"almuerzo_regular": 0, "complemento_am": 0, "complemento_pm": 0}
    )
else:
    total_cov_year = total_cov_selected.iloc[0]
gender_year = genero[(genero["anio"] == year) & (genero["genero"].str.lower() != "total")].copy()
gender_total = genero[(genero["anio"] == year) & (genero["genero"].str.lower() == "total")].copy()
grade_year = grado[(grado["anio"] == year) & (grado["grado"].str.lower() != "total")].sort_values("orden")
grade_total = grado[(grado["anio"] == year) & (grado["grado"].str.lower() == "total")]

previous = globales[globales["anio"] != year].sort_values("anio").tail(1)
if not previous.empty:
    previous = previous.iloc[0]
    delta_cov = global_year["cobertura_global"] - previous["cobertura_global"]
    delta_ben = global_year["beneficiarios"] - previous["beneficiarios"]
    delta_inv = global_year["inversion"] - previous["inversion"]
else:
    previous = None
    delta_cov = np.nan
    delta_ben = np.nan
    delta_inv = np.nan

st.markdown(
    f"""
    <div class="hero">
      <h1>Observatorio PAE, Distrito de Santa Marta</h1>
      <p>Plan Decenal de Educación</p>
      <span class="source">Año seleccionado: {year}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

selected_beneficiaries = inst_year["beneficiarios"].sum() if not inst_year.empty else 0
selected_matriculados = inst_year["matriculados"].sum() if not inst_year.empty else 0
selected_coverage = (
    selected_beneficiaries / selected_matriculados if selected_matriculados else np.nan
)

st.markdown(
    f"""
    <div class="kpis">
      <div class="kpi"><div class="label">Inversión</div><div class="value">{money(global_year["inversion"])}</div><div class="note">Recursos del periodo</div></div>
      <div class="kpi"><div class="label">Beneficiarios</div><div class="value">{n(global_year["beneficiarios"])}</div><div class="note">Población atendida</div></div>
      <div class="kpi"><div class="label">Matriculados</div><div class="value">{n(global_year["matriculados"])}</div><div class="note">Población escolar</div></div>
      <div class="kpi"><div class="label">Cobertura global</div><div class="value">{pct(global_year["cobertura_global"])}</div><div class="note">Alcance del programa</div></div>
      <div class="kpi"><div class="label">Instituciones</div><div class="value">{n(inst_year_all["institucion"].nunique())}</div><div class="note">Sedes educativas analizadas</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="guide">
      <div class="guide-card"><strong>1. Resumen</strong><span>Panorama general del distrito y evolución de las instituciones seleccionadas.</span></div>
      <div class="guide-card"><strong>2. Instituciones</strong><span>Ranking de cobertura, brecha y tabla para identificar prioridades.</span></div>
      <div class="guide-card"><strong>3. Perfil</strong><span>Distribución por modalidad, género y grado para lectura poblacional.</span></div>
    </div>
    <div class="insight">
      Selección actual: <strong>{n(len(selected_inst))}</strong> instituciones,
      <strong>{n(selected_beneficiaries)}</strong> beneficiarios y cobertura consolidada de
      <strong>{pct(selected_coverage)}</strong>.
    </div>
    """,
    unsafe_allow_html=True,
)

tab1, tab2, tab3, tab4 = st.tabs(
    ["1. Resumen ejecutivo", "2. Instituciones", "3. Género y grados", "4. Datos revisables"]
)

with tab1:
    st.markdown('<div class="question">Lectura general del distrito</div>', unsafe_allow_html=True)
    if previous is not None:
        trend_text = "sube" if delta_cov > 0 else "baja"
        beneficiary_text = "aumentan" if delta_ben > 0 else "disminuyen"
        st.markdown(
            f"""
            <div class="insight">
              Entre {int(previous["anio"])} y {year}, la cobertura global <strong>{trend_text}</strong>
              de {pct(previous["cobertura_global"])} a <strong>{pct(global_year["cobertura_global"])}</strong>.
              Los beneficiarios <strong>{beneficiary_text}</strong> en {n(abs(delta_ben))}, mientras la inversión cambia en
              <strong>{money(delta_inv)}</strong>.
              <div class="small">Lectura comparativa del desempeño general del programa.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    chart(
        {
            "color": ["#1f5f8b", "#00a19a", "#f4b942"],
            "tooltip": {"trigger": "axis"},
            "legend": {"top": 0},
            "grid": grid(),
            "xAxis": {"type": "category", "data": globales["anio"].astype(str).tolist()},
            "yAxis": [
                {"type": "value", "name": "Personas"},
                {"type": "value", "name": "Cobertura", "axisLabel": {"formatter": "{value}%"}},
            ],
            "series": [
                {"name": "Beneficiarios", "type": "bar", "barMaxWidth": 48, "data": globales["beneficiarios"].round(0).tolist()},
                {"name": "Matriculados", "type": "bar", "barMaxWidth": 48, "data": globales["matriculados"].round(0).tolist()},
                {"name": "Cobertura global", "type": "line", "yAxisIndex": 1, "smooth": True, "data": (globales["cobertura_global"] * 100).round(1).tolist()},
            ],
        },
        key="global_summary",
    )

    selected_count = len(selected_inst)
    if selected_count:
        institution_trend = (
            instituciones[instituciones["institucion"].isin(selected_inst)]
            .groupby("anio", as_index=False)
            .agg(
                matriculados=("matriculados", "sum"),
                beneficiarios=("beneficiarios", "sum"),
                brecha=("brecha", "sum"),
            )
            .sort_values("anio")
        )
        institution_trend["cobertura"] = (
            institution_trend["beneficiarios"]
            / institution_trend["matriculados"].replace(0, np.nan)
        )

        if selected_count == 1:
            selected_label = selected_inst[0]
        else:
            selected_label = f"{selected_count} instituciones seleccionadas"

        st.markdown(
            '<div class="question">Cobertura de instituciones seleccionadas</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="insight">
              Esta vista sí responde al filtro de institución. Selección actual:
              <strong>{selected_label}</strong>.
            </div>
            """,
            unsafe_allow_html=True,
        )
        chart(
            {
                "color": ["#1f5f8b", "#00a19a", "#df5c5c"],
                "tooltip": {"trigger": "axis"},
                "legend": {"top": 0},
                "grid": grid(),
                "xAxis": {"type": "category", "data": institution_trend["anio"].astype(str).tolist()},
                "yAxis": [
                    {"type": "value", "name": "Personas"},
                    {"type": "value", "name": "Cobertura", "axisLabel": {"formatter": "{value}%"}},
                ],
                "series": [
                    {
                        "name": "Matriculados",
                        "type": "bar",
                        "barMaxWidth": 46,
                        "data": institution_trend["matriculados"].round(0).tolist(),
                    },
                    {
                        "name": "Beneficiarios",
                        "type": "bar",
                        "barMaxWidth": 46,
                        "data": institution_trend["beneficiarios"].round(0).tolist(),
                    },
                    {
                        "name": "Cobertura",
                        "type": "line",
                        "yAxisIndex": 1,
                        "smooth": True,
                        "data": (institution_trend["cobertura"] * 100).round(1).tolist(),
                    },
                ],
            },
            key="institution_trend",
        )
    else:
        st.markdown(
            """
            <div class="insight warning">
              Selecciona una o varias instituciones para ver su evolución específica de cobertura.
            </div>
            """,
            unsafe_allow_html=True,
        )

    modality = pd.Series(
        {
            "Almuerzo regular": total_cov_year["almuerzo_regular"],
            "Complemento AM": total_cov_year["complemento_am"],
            "Complemento PM": total_cov_year["complemento_pm"],
        }
    )
    st.markdown('<div class="question">Cómo se distribuye el tipo de atención?</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="insight ok">
          Distribución de la atención alimentaria por modalidad para el año seleccionado.
          Total registrado por modalidad: <strong>{n(modality.sum())}</strong>.
        </div>
        """,
        unsafe_allow_html=True,
    )
    chart(
        {
            "color": COLORS,
            "tooltip": {"trigger": "item"},
            "legend": {"bottom": 0},
            "series": [
                {
                    "type": "pie",
                    "radius": ["42%", "72%"],
                    "itemStyle": {"borderRadius": 4, "borderColor": "#fff", "borderWidth": 2},
                    "label": {"formatter": "{b}: {d}%"},
                    "data": [{"name": k, "value": float(v)} for k, v in modality.items()],
                }
            ],
        },
        key="modality_total",
    )

with tab2:
    st.markdown('<div class="question">Qué instituciones tienen menor cobertura?</div>', unsafe_allow_html=True)
    if inst_year.empty:
        st.markdown(
            """
            <div class="insight warning">
              No hay instituciones seleccionadas. Usa el botón <strong>Seleccionar todas</strong> o elige una institución en la barra lateral.
            </div>
            """,
            unsafe_allow_html=True,
        )

    low = inst_year.sort_values("cobertura", ascending=True).head(15).sort_values("cobertura")
    st.markdown(
        """
        <div class="insight warning">
          Ranking de instituciones con menor cobertura reportada. La brecha ayuda a identificar prioridades de atención.
        </div>
        """,
        unsafe_allow_html=True,
    )
    chart(
        {
            "color": ["#df5c5c"],
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "grid": grid(),
            "xAxis": {"type": "value", "axisLabel": {"formatter": "{value}%"}},
            "yAxis": {"type": "category", "data": low["institucion"].tolist(), "axisLabel": {"fontSize": 10}},
            "series": [{"name": "%Cobertura", "type": "bar", "data": (low["cobertura"] * 100).round(1).fillna(0).tolist()}],
        },
        height="520px",
        key="low_coverage",
    )

    st.markdown('<div class="question">Dónde está la mayor brecha institucional?</div>', unsafe_allow_html=True)
    gap = inst_year.sort_values("brecha", ascending=False).head(15).sort_values("brecha")
    chart(
        {
            "color": ["#1f5f8b"],
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "grid": grid(),
            "xAxis": {"type": "value"},
            "yAxis": {"type": "category", "data": gap["institucion"].tolist(), "axisLabel": {"fontSize": 10}},
            "series": [{"name": "Brecha", "type": "bar", "data": gap["brecha"].round(0).tolist()}],
        },
        height="520px",
        key="gap",
    )

    view_cols = ["institucion", "matriculados", "beneficiarios", "cobertura", "brecha", "estado"]
    st.dataframe(inst_year.sort_values("cobertura")[view_cols], width="stretch", hide_index=True)

with tab3:
    st.markdown('<div class="question">Género</div>', unsafe_allow_html=True)
    if not gender_total.empty:
        st.markdown(
            f"""
            <div class="insight">
              Distribución por género de la población beneficiaria registrada para {year}: <strong>{n(gender_total.iloc[0]["cantidad"])}</strong> beneficiarios.
            </div>
            """,
            unsafe_allow_html=True,
        )
    chart(
        {
            "color": ["#df5c5c", "#1f5f8b"],
            "tooltip": {"trigger": "item"},
            "legend": {"bottom": 0},
            "series": [
                {
                    "type": "pie",
                    "radius": ["42%", "72%"],
                    "label": {"formatter": "{b}: {d}%"},
                    "data": gender_year.rename(columns={"genero": "name", "cantidad": "value"}).to_dict("records"),
                }
            ],
        },
        key="gender",
    )

    st.markdown('<div class="question">Acceso por grado</div>', unsafe_allow_html=True)
    if not grade_total.empty:
        st.markdown(
            f"""
            <div class="insight ok">
              Distribución por grado de la población beneficiaria registrada para {year}: <strong>{n(grade_total.iloc[0]["beneficiarios"])}</strong> beneficiarios.
            </div>
            """,
            unsafe_allow_html=True,
        )
    chart(
        {
            "color": ["#00a19a"],
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "grid": grid(),
            "xAxis": {"type": "category", "data": grade_year["grado"].tolist(), "name": "Grado"},
            "yAxis": {"type": "value", "name": "Beneficiarios"},
            "series": [{"name": "Beneficiarios", "type": "bar", "barMaxWidth": 44, "data": grade_year["beneficiarios"].round(0).tolist()}],
        },
        key="grade",
    )

with tab4:
    st.markdown('<div class="question">Datos revisables</div>', unsafe_allow_html=True)
    st.caption("Tablas de soporte para validar los indicadores y explorar el detalle.")
    st.subheader("Indicadores globales")
    st.dataframe(globales, width="stretch", hide_index=True)
    st.subheader("Cobertura institucional filtrada")
    st.dataframe(inst_year, width="stretch", hide_index=True)
    st.subheader("Género")
    st.dataframe(genero[genero["anio"] == year], width="stretch", hide_index=True)
    st.subheader("Acceso por grado")
    st.dataframe(grado[grado["anio"] == year], width="stretch", hide_index=True)
