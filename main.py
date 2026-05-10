import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# =====================================================
# STREAMLIT CONFIG
# =====================================================
st.set_page_config(
    page_title="Врачебные проекты | Dashboard",
    page_icon="👨‍⚕️",
    layout="wide"
)

# =====================================================
# STYLE
# =====================================================
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    .main-title {
        font-size: 34px;
        font-weight: 800;
        color: #1f2937;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 15px;
        color: #6b7280;
        margin-top: 0px;
        margin-bottom: 20px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =====================================================
# FUNCTIONS
# =====================================================
@st.cache_data(show_spinner=False)
def read_excel_file(uploaded_file):
    """Read all sheets from uploaded Excel file."""
    return pd.read_excel(uploaded_file, sheet_name=None)


def to_number(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def format_number(value):
    try:
        return f"{value:,.0f}".replace(",", " ")
    except Exception:
        return value


def make_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered_Data")
    return output.getvalue()


def find_column(df, possible_names):
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for name in possible_names:
        key = name.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
    return None

# =====================================================
# HEADER
# =====================================================
st.markdown('<p class="main-title">👨‍⚕️ Врачебные проекты — Dashboard</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-title">Загрузите актуальный Excel-файл, чтобы увидеть таблицу, фильтры, KPI и визуальную аналитику по МП, врачам, проектам и баллам.</p>',
    unsafe_allow_html=True
)

# =====================================================
# SIDEBAR FILE UPLOAD
# =====================================================
st.sidebar.header("📁 Загрузка файла")

uploaded_file = st.sidebar.file_uploader(
    "Загрузите Excel-файл",
    type=["xlsx", "xls"]
)

if uploaded_file is None:
    st.info("Загрузите Excel-файл слева, чтобы начать анализ.")
    st.stop()

try:
    sheets = read_excel_file(uploaded_file)
except Exception as e:
    st.error(f"Не удалось прочитать Excel-файл: {e}")
    st.stop()

sheet_names = list(sheets.keys())

# =====================================================
# SHEET SELECTION
# =====================================================
st.sidebar.header("📄 Выбор листа")

preferred_sheet = "Проекты" if "Проекты" in sheet_names else sheet_names[0]
selected_sheet = st.sidebar.selectbox(
    "Лист для анализа",
    sheet_names,
    index=sheet_names.index(preferred_sheet)
)

df = sheets[selected_sheet].copy()

# Remove fully empty rows and columns
df = df.dropna(how="all")
df = df.dropna(axis=1, how="all")

if df.empty:
    st.warning("Выбранный лист пустой.")
    st.stop()

# Clean column names
df.columns = [str(c).strip() for c in df.columns]

st.success(f"Файл загружен. Выбран лист: {selected_sheet}. Строк: {len(df):,}. Колонок: {len(df.columns)}".replace(",", " "))

# =====================================================
# AUTO COLUMN DETECTION
# =====================================================
doctor_col_auto = find_column(df, ["Врач", "ФИО врача", "Доктор", "Doctor"])
mp_col_auto = find_column(df, ["МП", "Мед пред", "Медпред", "Медицинский представитель", "Medical representative"])
region_col_auto = find_column(df, ["Регион", "Область", "Region"])
project_col_auto = find_column(df, ["Проект", "Название проекта", "Project"])
brand_col_auto = find_column(df, ["Бренд", "Препарат", "SKU", "СКЮ", "Brand"])
points_col_auto = find_column(df, ["Баллы", "Балл", "Проектные баллы", "Points", "Score"])
plan_col_auto = find_column(df, ["План", "Plan"])
fact_col_auto = find_column(df, ["Факт", "Fact"])

columns = df.columns.tolist()

# =====================================================
# SIDEBAR COLUMN SETTINGS
# =====================================================
st.sidebar.header("⚙️ Настройка колонок")

mp_col = st.sidebar.selectbox(
    "Колонка МП",
    [None] + columns,
    index=([None] + columns).index(mp_col_auto) if mp_col_auto in columns else 0
)

doctor_col = st.sidebar.selectbox(
    "Колонка врача",
    [None] + columns,
    index=([None] + columns).index(doctor_col_auto) if doctor_col_auto in columns else 0
)

region_col = st.sidebar.selectbox(
    "Колонка региона",
    [None] + columns,
    index=([None] + columns).index(region_col_auto) if region_col_auto in columns else 0
)

project_col = st.sidebar.selectbox(
    "Колонка проекта",
    [None] + columns,
    index=([None] + columns).index(project_col_auto) if project_col_auto in columns else 0
)

brand_col = st.sidebar.selectbox(
    "Колонка бренда / препарата",
    [None] + columns,
    index=([None] + columns).index(brand_col_auto) if brand_col_auto in columns else 0
)

points_col = st.sidebar.selectbox(
    "Колонка баллов / основного показателя",
    columns,
    index=columns.index(points_col_auto) if points_col_auto in columns else 0
)

plan_col = st.sidebar.selectbox(
    "Колонка плана",
    [None] + columns,
    index=([None] + columns).index(plan_col_auto) if plan_col_auto in columns else 0
)

fact_col = st.sidebar.selectbox(
    "Колонка факта",
    [None] + columns,
    index=([None] + columns).index(fact_col_auto) if fact_col_auto in columns else 0
)

# =====================================================
# FILTERS
# =====================================================
filtered_df = df.copy()

st.sidebar.header("🔎 Фильтры")

filter_columns = [region_col, mp_col, doctor_col, project_col, brand_col]
filter_columns = [c for c in filter_columns if c is not None]

for col in filter_columns:
    values = sorted(filtered_df[col].dropna().astype(str).unique())
    if len(values) > 0:
        selected_values = st.sidebar.multiselect(
            f"Фильтр: {col}",
            values,
            default=values
        )
        filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_values)]

# Numeric conversion
filtered_df[points_col] = to_number(filtered_df[points_col])

if plan_col:
    filtered_df[plan_col] = to_number(filtered_df[plan_col])

if fact_col:
    filtered_df[fact_col] = to_number(filtered_df[fact_col])

# =====================================================
# KPI BLOCK
# =====================================================
st.markdown("### 🔢 Основные KPI")

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

total_points = filtered_df[points_col].sum()
rows_count = len(filtered_df)
unique_doctors = filtered_df[doctor_col].nunique() if doctor_col else 0
unique_mp = filtered_df[mp_col].nunique() if mp_col else 0
avg_points = filtered_df[points_col].mean() if rows_count > 0 else 0

with kpi1:
    st.metric("Всего баллов", format_number(total_points))
with kpi2:
    st.metric("Строк", format_number(rows_count))
with kpi3:
    st.metric("Врачей", format_number(unique_doctors))
with kpi4:
    st.metric("МП", format_number(unique_mp))
with kpi5:
    st.metric("Средний балл", format_number(avg_points))

if plan_col and fact_col:
    st.markdown("### 📌 План / факт")

    plan_total = filtered_df[plan_col].sum()
    fact_total = filtered_df[fact_col].sum()
    execution = fact_total / plan_total if plan_total else 0
    diff = fact_total - plan_total

    p1, p2, p3, p4 = st.columns(4)
    with p1:
        st.metric("План", format_number(plan_total))
    with p2:
        st.metric("Факт", format_number(fact_total))
    with p3:
        st.metric("Выполнение", f"{execution:.1%}")
    with p4:
        st.metric("Отклонение", format_number(diff))

# =====================================================
# TABS
# =====================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Таблица",
    "📊 Аналитика по МП",
    "👨‍⚕️ Аналитика по врачам",
    "🏷️ Проекты / бренды"
])

# =====================================================
# TAB 1 TABLE
# =====================================================
with tab1:
    st.markdown("### 📋 Детальная таблица")

    st.dataframe(
        filtered_df,
        use_container_width=True,
        height=600
    )

    excel_data = make_excel_download(filtered_df)
    st.download_button(
        label="⬇️ Скачать отфильтрованную таблицу в Excel",
        data=excel_data,
        file_name="filtered_doctor_projects.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# TAB 2 MP ANALYTICS
# =====================================================
with tab2:
    st.markdown("### 📊 Баллы по медицинским представителям")

    if mp_col:
        mp_summary = (
            filtered_df.groupby(mp_col, as_index=False)[points_col]
            .sum()
            .sort_values(points_col, ascending=False)
        )

        fig_mp = px.bar(
            mp_summary,
            x=mp_col,
            y=points_col,
            text_auto=True,
            title="Рейтинг МП по баллам"
        )
        fig_mp.update_layout(xaxis_title="МП", yaxis_title="Баллы")
        st.plotly_chart(fig_mp, use_container_width=True)

        st.dataframe(mp_summary, use_container_width=True)
    else:
        st.warning("Выберите колонку МП в настройках слева.")

# =====================================================
# TAB 3 DOCTOR ANALYTICS
# =====================================================
with tab3:
    st.markdown("### 👨‍⚕️ Баллы по врачам")

    if doctor_col:
        top_n_doctors = st.slider("Количество врачей в ТОП", 5, 100, 20)

        doctor_summary = (
            filtered_df.groupby(doctor_col, as_index=False)[points_col]
            .sum()
            .sort_values(points_col, ascending=False)
            .head(top_n_doctors)
        )

        fig_doctor = px.bar(
            doctor_summary,
            x=points_col,
            y=doctor_col,
            orientation="h",
            text_auto=True,
            title=f"ТОП-{top_n_doctors} врачей по баллам"
        )
        fig_doctor.update_layout(xaxis_title="Баллы", yaxis_title="Врач")
        st.plotly_chart(fig_doctor, use_container_width=True)

        st.dataframe(doctor_summary, use_container_width=True)
    else:
        st.warning("Выберите колонку врача в настройках слева.")

# =====================================================
# TAB 4 PROJECT / BRAND ANALYTICS
# =====================================================
with tab4:
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("### 📌 Проекты")
        if project_col:
            project_summary = (
                filtered_df.groupby(project_col, as_index=False)[points_col]
                .sum()
                .sort_values(points_col, ascending=False)
            )

            fig_project = px.pie(
                project_summary,
                names=project_col,
                values=points_col,
                title="Доля проектов по баллам"
            )
            st.plotly_chart(fig_project, use_container_width=True)
            st.dataframe(project_summary, use_container_width=True)
        else:
            st.warning("Выберите колонку проекта.")

    with col_b:
        st.markdown("### 🏷️ Бренды / препараты")
        if brand_col:
            brand_summary = (
                filtered_df.groupby(brand_col, as_index=False)[points_col]
                .sum()
                .sort_values(points_col, ascending=False)
            )

            fig_brand = px.bar(
                brand_summary,
                x=brand_col,
                y=points_col,
                text_auto=True,
                title="Баллы по брендам / препаратам"
            )
            fig_brand.update_layout(xaxis_title="Бренд / препарат", yaxis_title="Баллы")
            st.plotly_chart(fig_brand, use_container_width=True)
            st.dataframe(brand_summary, use_container_width=True)
        else:
            st.warning("Выберите колонку бренда / препарата.")

# =====================================================
# FOOTER
# =====================================================
st.markdown("---")
st.caption("Dashboard для анализа врачебных проектов. Источник данных: Excel-файл, загружаемый пользователем вручную.")