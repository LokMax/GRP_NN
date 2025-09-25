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
    """Убедиться, что есть колонка 'year' (или 'год', или извлечь из 'date')."""
    cols_lower = {c.lower(): c for c in df.columns}

    if "year" in cols_lower:
        c = cols_lower["year"]
        df["year"] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
        return df

    if "год" in cols_lower:
        c = cols_lower["год"]
        df["year"] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
        return df

    if "date" in cols_lower:
        c = cols_lower["date"]
        tmp = pd.to_datetime(df[c], errors="coerce")
        df["year"] = tmp.dt.year.astype("Int64")
        return df

    for col in df.columns:
        cl = col.lower()
        if "year" in cl or "год" in cl:
            df["year"] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
            return df

    raise ValueError("Не найдена колонка с годом. Добавьте колонку 'year', 'год' или 'date'.")

@st.cache_data(show_spinner=False)
def load_excel(content_or_path, sheet_name: str = "Данные") -> pd.DataFrame:
    """Загрузить Excel из файла или байтов."""
    if isinstance(content_or_path, (bytes, bytearray)):
        bio = io.BytesIO(content_or_path)
        df = pd.read_excel(bio, sheet_name=sheet_name, engine="openpyxl")
    elif isinstance(content_or_path, io.BytesIO):
        df = pd.read_excel(content_or_path, sheet_name=sheet_name, engine="openpyxl")
    else:
        df = pd.read_excel(content_or_path, sheet_name=sheet_name, engine="openpyxl")
    return df

def ensure_required_columns(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"В данных отсутствуют обязательные колонки: {', '.join(missing)}")
        st.stop()

def coerce_numeric(df: pd.DataFrame, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# ---------------------- Data input ----------------------
st.sidebar.header("Данные")
sheet_name = st.sidebar.text_input("Лист Excel", value="Данные")
uploaded = st.sidebar.file_uploader("Загрузите Excel-файл (.xlsx)", type=["xlsx"])
default_path = Path("Data.xlsx")

df = None

if uploaded is not None:
    try:
        df = load_excel(uploaded.getvalue(), sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Не удалось прочитать загруженный файл: {e}")
elif default_path.exists():
    try:
        df = load_excel(default_path, sheet_name=sheet_name)
        st.sidebar.info("Используется локальный файл Data.xlsx из корня репозитория.")
    except Exception as e:
        st.error(f"Не удалось прочитать файл Data.xlsx: {e}")
else:
    st.warning("Загрузите файл Excel (слева) или поместите 'Data.xlsx' в корень репозитория.")
    st.stop()

# Обработка колонок
try:
    df = _ensure_year_column(df)
except ValueError as e:
    st.error(str(e))
    st.stop()

ensure_required_columns(df)
df = coerce_numeric(df, REQUIRED_COLS + ["year"])

# Опциональная дополнительная категоризация (например, industry)
extra_group_col = None
for candidate in ["industry", "region", "sector", "отрасль", "регион"]:
    if candidate in df.columns:
        extra_group_col = candidate
        break

# ---------------------- Sidebar filters ----------------------
df["year"] = pd.to_datetime(df["year"], errors="coerce").dt.year
df = df.dropna(subset=["year"])
df["year"] = df["year"].astype(int)

years_series = df["year"]
if years_series.empty:
    st.error("Не удалось определить диапазон лет.")
    st.stop()

year_min, year_max = int(years_series.min()), int(years_series.max())

selected_years = st.sidebar.slider(
    "Выберите диапазон лет",
    min_value=year_min,
    max_value=year_max,
    value=(year_min, year_max),
    step=1
)

color_by = st.sidebar.selectbox(
    "Цвет точек по",
    ["year"] + ([extra_group_col] if extra_group_col else [])
)

# Фильтрация по выбранным годам
dff = df[(df["year"] >= selected_years[0]) & (df["year"] <= selected_years[1])].copy()

# ---------------------- Info metrics ----------------------
st.markdown("### Фильтры")
col1, col2, col3 = st.columns([1,1,1])
with col1:
    st.metric("Мин. год", selected_years[0])
with col2:
    st.metric("Макс. год", selected_years[1])
with col3:
    st.metric("Число наблюдений", len(dff))

st.divider()

# ---------------------- Chart 1: real_GVA vs real_support ----------------------
fig1 = px.scatter(
    dff,
    x="real_support",
    y="real_GVA",
    color=color_by,
    hover_data=dff.columns
)
fig1.update_layout(
    title="Диаграмма рассеяния: real_GVA vs real_support",
    xaxis_title="real_support",
    yaxis_title="real_GVA",
    legend_title=color_by
)
st.plotly_chart(fig1, use_container_width=True)

# ---------------------- Chart 2: real_GVA vs real_subsidies ----------------------
fig2 = px.scatter(
    dff,
    x="real_subsidies",
    y="real_GVA",
    color=color_by,
    hover_data=dff.columns
)
fig2.update_layout(
    title="Диаграмма рассеяния: real_GVA vs real_subsidies",
    xaxis_title="real_subsidies",
    yaxis_title="real_GVA",
    legend_title=color_by
)
st.plotly_chart(fig2, use_container_width=True)

# ---------------------- Chart 3: Bubble ----------------------
fig3 = px.scatter(
    dff,
    x="real_support",
    y="real_subsidies",
    size="real_GVA",
    color=color_by,
    size_max=40,
    hover_data=dff.columns
)
fig3.update_layout(
    title="Пузырьковая диаграмма: real_support vs real_subsidies (размер = real_GVA)",
    xaxis_title="real_support",
    yaxis_title="real_subsidies",
    legend_title=color_by
)
st.plotly_chart(fig3, use_container_width=True)

# ---------------------- Data preview ----------------------
with st.expander("Показать первые строки данных"):
    st.dataframe(dff.head(50), use_container_width=True)

st.caption("Разместите файл 'Data.xlsx' в корень репозитория или загрузите его через сайдбар.")
