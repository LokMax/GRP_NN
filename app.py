import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------- Page config ----------------------
st.set_page_config(
    page_title="Economic Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Экономический Dashboard")
st.caption("Интерактивные диаграммы по данным из листа 'Данные' Excel-файла")

# ---------------------- Helpers ----------------------
REQUIRED_COLS = ["real_GVA", "real_support", "real_subsidies"]

def _ensure_year_column(df: pd.DataFrame) -> pd.DataFrame:
    cols_lower = {c.lower(): c for c in df.columns}
    if "year" in cols_lower:
        df["year"] = pd.to_numeric(df[cols_lower["year"]], errors="coerce")
        return df
    if "год" in cols_lower:
        df["year"] = pd.to_numeric(df[cols_lower["год"]], errors="coerce")
        return df
    if "date" in cols_lower:
        df["year"] = pd.to_datetime(df[cols_lower["date"]], errors="coerce").dt.year
        return df
    for col in df.columns:
        if "year" in col.lower() or "год" in col.lower():
            df["year"] = pd.to_numeric(df[col], errors="coerce")
            return df
    raise ValueError("Не найдена колонка с годом")

@st.cache_data(show_spinner=False)
def load_excel(content_or_path, sheet_name: str = "Данные") -> pd.DataFrame:
    if isinstance(content_or_path, (bytes, bytearray)):
        return pd.read_excel(io.BytesIO(content_or_path), sheet_name=sheet_name, engine="openpyxl")
    if isinstance(content_or_path, io.BytesIO):
        return pd.read_excel(content_or_path, sheet_name=sheet_name, engine="openpyxl")
    return pd.read_excel(content_or_path, sheet_name=sheet_name, engine="openpyxl")

def ensure_required_columns(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"В данных отсутствуют обязательные колонки: {', '.join(missing)}")
        st.stop()

# ---------------------- Data input ----------------------
st.sidebar.header("Данные")
sheet_name = st.sidebar.text_input("Лист Excel", value="Данные")
uploaded = st.sidebar.file_uploader("Загрузите Excel-файл (.xlsx)", type=["xlsx"])
default_path = Path("Data.xlsx")

df = None
if uploaded is not None:
    df = load_excel(uploaded.getvalue(), sheet_name=sheet_name)
elif default_path.exists():
    df = load_excel(default_path, sheet_name=sheet_name)
    st.sidebar.info("Используется локальный файл Data.xlsx")
else:
    st.warning("Загрузите файл Excel или поместите Data.xlsx в корень репозитория")
    st.stop()

df = _ensure_year_column(df)
ensure_required_columns(df)
df["year"] = pd.to_numeric(df["year"], errors="coerce")
df = df.dropna(subset=["year"])
df["year"] = df["year"].astype(int)

# ---------------------- Sidebar filters ----------------------
df["year"] = pd.to_datetime(df["year"], errors="coerce").dt.year
df = df.dropna(subset=["year"])
df["year"] = df["year"].astype(int)

year_min, year_max = int(df["year"].min()), int(df["year"].max())

selected_years = st.sidebar.slider(
    "Выберите диапазон лет",
    min_value=year_min,
    max_value=year_max,
    value=(year_min, year_max),
    step=1
)

# Фильтр по отраслям (если колонка есть)
industry_col = None
for candidate in ["industry", "отрасль", "sector"]:
    if candidate in df.columns:
        industry_col = candidate
        break

selected_industries = None
if industry_col:
    industries = sorted(df[industry_col].dropna().unique())
    selected_industries = st.sidebar.multiselect(
        "Выберите отрасли", options=industries, default=industries
    )

# Проигрыватель годов
year_player = st.sidebar.select_slider(
    "Год (проигрыватель)",
    options=sorted(df["year"].unique()),
    value=year_min
)

# ---------------------- Apply filters ----------------------
dff = df[(df["year"] >= selected_years[0]) & (df["year"] <= selected_years[1])].copy()

if selected_industries is not None:
    dff = dff[dff[industry_col].isin(selected_industries)]

# Отдельный DataFrame для проигрывателя (один год)
df_year = dff[dff["year"] == year_player]

# ---------------------- Charts ----------------------
st.markdown("### Фильтры")
c1, c2, c3 = st.columns(3)
c1.metric("Мин. год", selected_years[0])
c2.metric("Макс. год", selected_years[1])
c3.metric("Число наблюдений", len(dff))

st.divider()

# Chart 1
fig1 = px.scatter(
    dff,
    x="real_support", y="real_GVA",
    color="year", hover_data=dff.columns
)
st.plotly_chart(fig1, use_container_width=True)

# Chart 2
fig2 = px.scatter(
    dff,
    x="real_subsidies", y="real_GVA",
    color="year", hover_data=dff.columns
)
st.plotly_chart(fig2, use_container_width=True)

# Chart 3 (bubble) — динамический по выбранному году
fig3 = px.scatter(
    df_year,
    x="real_support", y="real_subsidies",
    size="real_GVA", color="year",
    size_max=40, hover_data=df_year.columns
)
st.plotly_chart(fig3, use_container_width=True)

with st.expander("Показать первые строки данных"):
    st.dataframe(dff.head(50), use_container_width=True)
