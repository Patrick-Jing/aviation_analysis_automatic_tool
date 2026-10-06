"""Aviation Market Analysis Automation Tool Streamlit MVP。"""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib import font_manager


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
OAG_FILE = DATA_DIR / "oag_data.xlsx"
CIRIUM_FILE = DATA_DIR / "cirium_trend_data.xlsx"
IATA_FILE = DATA_DIR / "iata_data.xlsx"

OAG_COLUMNS = [
    "Carrier Name",
    "Dep Airport Code",
    "Arr Airport Code",
    "Frequency",
]
FLEET_YEARS = [str(year) for year in range(2015, 2026)]
COUNTRY_FLEET_YEARS = [str(year) for year in range(1990, 2027)]
CIRIUM_COLUMNS = [
    "CAT",
    "Type",
    "Operator",
    "Manufacturer",
    "Metric",
    *FLEET_YEARS,
]
COUNTRY_FLEET_COLUMNS = [
    "Operator",
    "Operator Country/Subregion",
    "Metric",
    *COUNTRY_FLEET_YEARS,
]
IATA_COLUMNS = ["Seg Orig", "Seg Dest", "Fare", "Pax Count", "Distance"]
PASSENGER_GROUP_ORDER = [">400", "100-400", "50-100", "0-50"]
TARGET_FLEET_METRICS = {"total in service", "total in storage"}

OPTION_ROUTE = (
    "查询某航司从某机场出发的C909航程覆盖范围内主要机场三字码"
    "（航程圈图输入）"
)
OPTION_FLEET = "生成某航司2015至2025年机队规模变化表"
OPTION_DISTRIBUTION = "生成某航司基于日单向客流的航线分布表"
OPTION_COUNTRY_FLEET = "生成某国家各航司1990至2026年机队规模变化表"


def standardize_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    """清理列名与文本字段，返回标准化后的宽表。"""
    cleaned = dataframe.copy()
    cleaned.columns = [str(column).strip() for column in cleaned.columns]

    for column in cleaned.select_dtypes(include=["object", "string"]).columns:
        cleaned[column] = cleaned[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )

    return cleaned


def read_excel(source: object) -> pd.DataFrame:
    """读取单工作表 Excel，并统一年份列名等字段格式。"""
    dataframe = pd.read_excel(source, sheet_name=0)
    return standardize_dataframe(dataframe)


def find_missing_values(dataframe: pd.DataFrame) -> dict[str, int]:
    """返回存在缺失值的列及各列缺失数量。"""
    missing_counts = dataframe.isna().sum()
    return {
        str(column): int(count)
        for column, count in missing_counts.items()
        if count > 0
    }


def validate_required_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str],
) -> list[str]:
    """找出分析所需但未出现在数据中的字段。"""
    return [
        column for column in required_columns if column not in dataframe.columns
    ]


def validate_dataframe(
    dataframe: pd.DataFrame,
    table_name: str,
    required_columns: list[str],
) -> bool:
    """校验整张表的缺失值与必要字段，并展示明确错误。"""
    missing_values = find_missing_values(dataframe)
    if missing_values:
        details = "；".join(
            f"{column}（{count} 个）"
            for column, count in missing_values.items()
        )
        st.error(
            f"{table_name}存在缺失值：{details}。"
            "请处理后再上传。"
        )
        return False

    missing_columns = validate_required_columns(dataframe, required_columns)
    if missing_columns:
        st.error(
            f"{table_name}缺少分析所需列："
            f"{', '.join(missing_columns)}。请修正后再上传。"
        )
        return False

    return True


def save_uploaded_file(uploaded_file: object, target_path: Path) -> None:
    """将已通过校验的原始 Excel 持久化保存到 data 目录。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    target_path.write_bytes(uploaded_file.getvalue())


def process_upload(
    uploaded_file: object | None,
    table_name: str,
    target_path: Path,
    required_columns: list[str],
) -> bool:
    """上传后立即读取、校验；仅将合格文件写入后台。"""
    if uploaded_file is None:
        return True

    try:
        dataframe = read_excel(BytesIO(uploaded_file.getvalue()))
    except Exception as error:  # Excel 引擎会提供具体的文件错误原因
        st.error(f"无法读取{table_name}：{error}")
        return False

    if not validate_dataframe(dataframe, table_name, required_columns):
        return False

    save_uploaded_file(uploaded_file, target_path)
    st.success(f"{table_name}上传并保存成功。")
    return True


def load_persisted_file(
    path: Path,
    table_name: str,
    required_columns: list[str],
) -> pd.DataFrame | None:
    """从 data 目录加载持久化数据，并再次执行完整性校验。"""
    if not path.exists():
        return None

    try:
        dataframe = read_excel(path)
    except Exception as error:
        st.error(f"后台保存的{table_name}无法读取：{error}")
        return None

    if not validate_dataframe(dataframe, table_name, required_columns):
        return None
    return dataframe


def dataframe_to_excel(
    dataframe: pd.DataFrame,
    sheet_name: str,
    include_index: bool = False,
) -> bytes:
    """将分析结果转换为可下载的 Excel 二进制内容。"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(
            writer,
            index=include_index,
            sheet_name=sheet_name,
        )
    return output.getvalue()


def render_route_analysis(oag_data: pd.DataFrame) -> None:
    """筛选指定航司、出发机场及高于中位数的航线。"""
    st.info("需提供 Carrier Name 和 Dep Airport Code")

    with st.form("route_analysis_form"):
        carrier_name = st.text_input(
            "Carrier Name",
            placeholder="例如：China Eastern Airlines",
        )
        dep_airport_code = st.text_input(
            "Dep Airport Code",
            placeholder="例如：PVG",
            max_chars=3,
        )
        submitted = st.form_submit_button("开始分析", type="primary")

    if not submitted:
        return
    if not carrier_name.strip() or not dep_airport_code.strip():
        st.warning("请完整填写 Carrier Name 和 Dep Airport Code。")
        return

    working_data = oag_data[OAG_COLUMNS].copy()
    working_data["Frequency"] = pd.to_numeric(
        working_data["Frequency"], errors="coerce"
    )
    if working_data["Frequency"].isna().any():
        st.error(
            "OAG数据的 Frequency 列包含非数值内容，请处理后再上传。"
        )
        return

    selected = working_data.loc[
        working_data["Carrier Name"].astype(str).str.casefold()
        == carrier_name.strip().casefold()
    ]
    selected = selected.loc[
        selected["Dep Airport Code"].astype(str).str.upper()
        == dep_airport_code.strip().upper()
    ]

    if selected.empty:
        st.warning("未找到同时匹配该航司与出发机场的数据。")
        return

    frequency_median = selected["Frequency"].median()
    result = selected.loc[
        selected["Frequency"] > frequency_median,
        ["Arr Airport Code"],
    ].reset_index(drop=True)

    if result.empty:
        st.warning("匹配数据中没有 Frequency 高于中位数的航线。")
        return

    st.success(
        f"分析完成：Frequency 中位数为 {frequency_median:g}，"
        f"共提取 {len(result)} 条到达机场记录。"
    )
    st.dataframe(result, width="stretch", hide_index=True)
    st.download_button(
        "下载航程圈图输入 Excel",
        data=dataframe_to_excel(result, "Airport Codes"),
        file_name="c909_range_airport_codes.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def build_fleet_pivot(
    cirium_data: pd.DataFrame,
    operator_name: str,
) -> pd.DataFrame:
    """建立按 CAT 汇总的 2015—2025 年机队规模透视表。"""
    working_data = cirium_data[CIRIUM_COLUMNS].copy()
    selected = working_data.loc[
        (
            working_data["Operator"].astype(str).str.casefold()
            == operator_name.strip().casefold()
        )
        & working_data["Metric"]
        .astype(str)
        .str.casefold()
        .isin(TARGET_FLEET_METRICS)
    ].copy()

    if selected.empty:
        return pd.DataFrame()

    for year in FLEET_YEARS:
        selected[year] = pd.to_numeric(selected[year], errors="coerce")

    invalid_years = [
        year for year in FLEET_YEARS if selected[year].isna().any()
    ]
    if invalid_years:
        raise ValueError(
            "以下年份列包含非数值内容：" + ", ".join(invalid_years)
        )

    return pd.pivot_table(
        selected,
        index="CAT",
        values=FLEET_YEARS,
        aggfunc="sum",
        fill_value=0,
    ).reindex(columns=FLEET_YEARS)


def plot_fleet_pivot(pivot_table: pd.DataFrame) -> plt.Figure:
    """使用商飞蓝风格绘制 CAT 分类的机队规模堆积柱状图。"""
    # 优先使用已安装的中文字体，避免标题和坐标轴显示为方框。
    available_fonts = {font.name for font in font_manager.fontManager.ttflist}
    chinese_font_candidates = [
        "PingFang SC",
        "Arial Unicode MS",
        "Hiragino Sans GB",
        "Heiti SC",
        "Microsoft YaHei",
        "SimHei",
        "Noto Sans CJK SC",
    ]
    chinese_font = next(
        (
            font_name
            for font_name in chinese_font_candidates
            if font_name in available_fonts
        ),
        "DejaVu Sans",
    )
    plt.rcParams["font.sans-serif"] = [chinese_font, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    category_count = len(pivot_table.index)
    color_map = plt.get_cmap("Blues")
    colors = [
        color_map(0.48 + 0.44 * index / max(category_count - 1, 1))
        for index in range(category_count)
    ]
    # 主色 #004EA2 对应需求中的 R0 G78 B162（商飞蓝风格）。
    if category_count == 1:
        colors = ["#004EA2"]

    figure, axis = plt.subplots(figsize=(8, 4.5), dpi=120)
    pivot_table.T.plot(
        kind="bar",
        stacked=True,
        ax=axis,
        color=colors,
        width=0.78,
    )
    axis.set_title("2015—2025年机队规模变化", fontsize=14, pad=12)
    axis.set_xlabel("年份", fontsize=11)
    axis.set_ylabel("机队数量", fontsize=11)
    axis.tick_params(axis="x", rotation=0)
    axis.grid(axis="y", linestyle="--", alpha=0.25)
    axis.legend(title="CAT", bbox_to_anchor=(1.02, 1), loc="upper left")
    figure.tight_layout()
    return figure


def render_fleet_analysis(cirium_data: pd.DataFrame) -> None:
    """生成机队规模透视表、堆积柱状图及 Excel 下载。"""
    st.info("需提供 Operator")

    with st.form("fleet_analysis_form"):
        operator_name = st.text_input(
            "Operator",
            placeholder="请输入与 Operator 列一致的英文名",
        )
        submitted = st.form_submit_button("开始分析", type="primary")

    if not submitted:
        return
    if not operator_name.strip():
        st.warning("请输入 Operator。")
        return

    try:
        pivot_table = build_fleet_pivot(cirium_data, operator_name)
    except ValueError as error:
        st.error(str(error))
        return

    if pivot_table.empty:
        st.warning(
            "未找到该 Operator 且 Metric 为 Total in Service 或 "
            "Total in Storage 的数据。"
        )
        return

    st.success("机队规模变化表生成完成。")
    st.dataframe(pivot_table, width="stretch")

    figure = plot_fleet_pivot(pivot_table)
    # 中间列占页面宽度的 1/2，图表面积约缩小至原来的 1/4。
    _, chart_column, _ = st.columns([1, 2, 1])
    with chart_column:
        st.pyplot(figure, width="stretch")
    plt.close(figure)

    st.download_button(
        "下载机队规模变化表 Excel",
        data=dataframe_to_excel(
            pivot_table,
            "Fleet Size 2015-2025",
            include_index=True,
        ),
        file_name="fleet_size_2015_2025.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def parse_time_range(time_range: str) -> tuple[int, int]:
    """解析并校验 1990—2026 范围内的 YYYY-YYYY 时间区间。"""
    match = re.fullmatch(r"\s*(\d{4})\s*-\s*(\d{4})\s*", time_range)
    if match is None:
        raise ValueError("Time range 格式错误，请按 YYYY-YYYY 填写。")

    start_year, end_year = map(int, match.groups())
    if start_year < 1990 or end_year > 2026:
        raise ValueError("Time range 必须位于 1990-2026 范围内。")
    if start_year > end_year:
        raise ValueError("Time range 的开始年份不能晚于结束年份。")
    return start_year, end_year


def build_country_fleet_pivot(
    cirium_data: pd.DataFrame,
    country_name: str,
    start_year: int,
    end_year: int,
) -> pd.DataFrame:
    """按国家、航司和指定年份汇总在役及停场机队规模。"""
    selected_years = [
        str(year) for year in range(start_year, end_year + 1)
    ]
    missing_columns = validate_required_columns(
        cirium_data,
        COUNTRY_FLEET_COLUMNS,
    )
    if missing_columns:
        raise ValueError(
            "Cirium Trend数据表缺少分析所需列："
            + ", ".join(missing_columns)
        )

    working_data = cirium_data[COUNTRY_FLEET_COLUMNS].copy()
    selected = working_data.loc[
        (
            working_data["Operator Country/Subregion"]
            .astype(str)
            .str.casefold()
            == country_name.strip().casefold()
        )
        & working_data["Metric"]
        .astype(str)
        .str.casefold()
        .isin(TARGET_FLEET_METRICS)
    ].copy()

    if selected.empty:
        return pd.DataFrame()

    # 用户选定时间区间后，删除 1990—2026 中不在范围内的年份列。
    selected = selected[["Operator", *selected_years]].copy()
    for year in selected_years:
        selected[year] = pd.to_numeric(selected[year], errors="coerce")

    invalid_years = [
        year for year in selected_years if selected[year].isna().any()
    ]
    if invalid_years:
        raise ValueError(
            "以下年份列包含非数值内容：" + ", ".join(invalid_years)
        )

    pivot_table = pd.pivot_table(
        selected,
        index="Operator",
        values=selected_years,
        aggfunc="sum",
        fill_value=0,
    ).reindex(columns=selected_years)
    return pivot_table.sort_index(key=lambda index: index.str.casefold())


def plot_country_fleet_trend(
    pivot_table: pd.DataFrame,
    country_name: str,
    start_year: int,
    end_year: int,
) -> plt.Figure:
    """使用商飞蓝色系绘制各航司机队规模变化折线图。"""
    available_fonts = {font.name for font in font_manager.fontManager.ttflist}
    chinese_font_candidates = [
        "PingFang SC",
        "Arial Unicode MS",
        "Hiragino Sans GB",
        "Heiti SC",
        "Microsoft YaHei",
        "SimHei",
        "Noto Sans CJK SC",
    ]
    chinese_font = next(
        (
            font_name
            for font_name in chinese_font_candidates
            if font_name in available_fonts
        ),
        "DejaVu Sans",
    )
    plt.rcParams["font.sans-serif"] = [chinese_font, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False

    operator_count = len(pivot_table.index)
    color_map = plt.get_cmap("Blues")
    colors = [
        color_map(0.38 + 0.57 * index / max(operator_count - 1, 1))
        for index in range(operator_count)
    ]
    if operator_count == 1:
        colors = ["#004EA2"]

    years = [int(year) for year in pivot_table.columns]
    figure, axis = plt.subplots(figsize=(11, 5.8), dpi=120)
    for color, (operator, values) in zip(
        colors,
        pivot_table.iterrows(),
    ):
        axis.plot(
            years,
            values.to_numpy(dtype=float),
            label=str(operator),
            color=color,
            linewidth=2,
            marker="o",
            markersize=3,
        )

    tick_step = max(1, (len(years) + 11) // 12)
    axis.set_xticks(years[::tick_step])
    axis.set_title(
        f"{country_name}各航司{start_year}—{end_year}年机队规模变化",
        fontsize=14,
        pad=12,
    )
    axis.set_xlabel("年份", fontsize=11)
    axis.set_ylabel("机队数量", fontsize=11)
    axis.grid(axis="both", linestyle="--", alpha=0.25)
    axis.legend(
        title="Operator",
        bbox_to_anchor=(1.02, 1),
        loc="upper left",
        fontsize=8,
        ncol=2 if operator_count > 10 else 1,
    )
    figure.tight_layout()
    return figure


def render_country_fleet_analysis(cirium_data: pd.DataFrame) -> None:
    """渲染某国家各航司机队规模变化分析模块。"""
    st.info(
        "需提供 Operator Country/Subregion 和 Time range"
        "（请输入时间区间，例如2000-2025）"
    )

    with st.form("country_fleet_analysis_form"):
        input_column_1, input_column_2 = st.columns(2)
        with input_column_1:
            country_name = st.text_input(
                "Operator Country/Subregion",
                placeholder="例如：China",
            )
        with input_column_2:
            time_range = st.text_input(
                "Time range",
                placeholder="例如：2000-2025",
            )
        submitted = st.form_submit_button("开始分析", type="primary")

    if not submitted:
        return
    if not country_name.strip() or not time_range.strip():
        st.warning(
            "请完整填写 Operator Country/Subregion 和 Time range。"
        )
        return

    try:
        start_year, end_year = parse_time_range(time_range)
        pivot_table = build_country_fleet_pivot(
            cirium_data,
            country_name,
            start_year,
            end_year,
        )
    except ValueError as error:
        st.error(str(error))
        return

    if pivot_table.empty:
        st.warning(
            "未找到该 Operator Country/Subregion 且 Metric 为 "
            "Total in Service 或 Total in Storage 的数据。"
        )
        return

    st.success(
        f"分析完成：共汇总 {len(pivot_table)} 家航司，"
        f"时间范围为 {start_year}-{end_year}。"
    )
    st.dataframe(pivot_table, width="stretch")

    figure = plot_country_fleet_trend(
        pivot_table,
        country_name.strip(),
        start_year,
        end_year,
    )
    st.pyplot(figure, width="stretch")
    plt.close(figure)

    st.download_button(
        "下载国家航司机队规模变化表 Excel",
        data=dataframe_to_excel(
            pivot_table,
            "Country Fleet Trend",
            include_index=True,
        ),
        file_name=(
            f"country_operator_fleet_{start_year}_{end_year}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def classify_passenger_volume(volume: float) -> str:
    """按照已确认的边界划分日单向客流等级。"""
    if volume > 400:
        return ">400"
    if volume >= 100:
        return "100-400"
    if volume >= 50:
        return "50-100"
    return "0-50"


def build_route_distribution(
    iata_data: pd.DataFrame,
    oag_data: pd.DataFrame,
    carrier_name: str,
) -> tuple[pd.DataFrame, int, int, float]:
    """汇总 IATA 与 OAG 航线，生成日单向客流分布透视表。"""
    iata_working = iata_data[IATA_COLUMNS].copy()
    numeric_columns = ["Fare", "Pax Count", "Distance"]
    for column in numeric_columns:
        iata_working[column] = pd.to_numeric(
            iata_working[column], errors="coerce"
        )

    invalid_columns = [
        column
        for column in numeric_columns
        if iata_working[column].isna().any()
    ]
    if invalid_columns:
        raise ValueError(
            "IATA数据的以下列包含非数值内容："
            + ", ".join(invalid_columns)
        )
    if (iata_working["Pax Count"] < 0).any():
        raise ValueError(
            "IATA数据的 Pax Count 存在负客流，请处理后再上传。"
        )

    iata_working["Orig-Dest"] = (
        iata_working["Seg Orig"].astype(str).str.upper()
        + "-"
        + iata_working["Seg Dest"].astype(str).str.upper()
    )
    # 每条 IATA 航线仅保留一行：Fare 平均、客流求和、距离平均。
    iata_routes = (
        iata_working.groupby("Orig-Dest", as_index=False)
        .agg(
            {
                "Fare": "mean",
                "Pax Count": "sum",
                "Distance": "mean",
            }
        )
    )

    if (iata_routes["Distance"] <= 0).any():
        raise ValueError(
            "IATA汇总数据存在 Distance 小于或等于 0 的航线，"
            "无法计算 Fare per pax kilometer。"
        )

    iata_routes["Daily One-way Passenger Volume"] = (
        iata_routes["Pax Count"] / 365
    )
    iata_routes["Daily One-way Passenger Volume Group"] = (
        iata_routes["Daily One-way Passenger Volume"].map(
            classify_passenger_volume
        )
    )
    iata_routes["Daily One-way Passenger Volume Group"] = pd.Categorical(
        iata_routes["Daily One-way Passenger Volume Group"],
        categories=PASSENGER_GROUP_ORDER,
        ordered=True,
    )
    iata_routes["Fare per pax kilometer"] = (
        iata_routes["Fare"] / iata_routes["Distance"]
    )

    oag_working = oag_data[
        [
            "Carrier Name",
            "Dep Airport Code",
            "Arr Airport Code",
            "Frequency",
        ]
    ].copy()
    oag_working["Frequency"] = pd.to_numeric(
        oag_working["Frequency"], errors="coerce"
    )
    if oag_working["Frequency"].isna().any():
        raise ValueError(
            "OAG数据的 Frequency 列包含非数值内容，请处理后再上传。"
        )

    oag_working = oag_working.loc[
        oag_working["Carrier Name"].astype(str).str.casefold()
        == carrier_name.strip().casefold()
    ].copy()
    oag_working["Dep-Arr"] = (
        oag_working["Dep Airport Code"].astype(str).str.upper()
        + "-"
        + oag_working["Arr Airport Code"].astype(str).str.upper()
    )
    # 连接前先汇总同航司的重复航线，避免 Frequency 被重复放大。
    oag_routes = (
        oag_working.groupby("Dep-Arr", as_index=False)["Frequency"].sum()
    )

    merged = iata_routes.merge(
        oag_routes,
        how="left",
        left_on="Orig-Dest",
        right_on="Dep-Arr",
        validate="one_to_one",
        indicator=True,
    )
    unmatched_count = int(merged["Frequency"].isna().sum())
    total_count = len(merged)
    unmatched_ratio = unmatched_count / total_count if total_count else 0.0
    merged["Frequency"] = merged["Frequency"].fillna(0)

    pivot_table = pd.pivot_table(
        merged,
        index="Daily One-way Passenger Volume Group",
        values=["Fare per pax kilometer", "Distance", "Frequency"],
        aggfunc={
            "Fare per pax kilometer": "mean",
            "Distance": "mean",
            "Frequency": "sum",
        },
        observed=True,
        sort=False,
    )
    existing_groups = [
        group for group in PASSENGER_GROUP_ORDER if group in pivot_table.index
    ]
    pivot_table = pivot_table.reindex(index=existing_groups)
    pivot_table = pivot_table.reindex(
        columns=["Fare per pax kilometer", "Distance", "Frequency"]
    )
    return pivot_table, unmatched_count, total_count, unmatched_ratio


def render_route_distribution(
    iata_data: pd.DataFrame,
    oag_data: pd.DataFrame,
) -> None:
    """渲染航司日单向客流航线分布分析模块。"""
    st.info("需提供 Carrier Name")

    with st.form("route_distribution_form"):
        carrier_name = st.text_input(
            "Carrier Name",
            placeholder="请输入与 OAG Carrier Name 列一致的航司名称",
        )
        submitted = st.form_submit_button("开始分析", type="primary")

    if not submitted:
        return
    if not carrier_name.strip():
        st.warning("请输入 Carrier Name。")
        return

    try:
        result, unmatched_count, total_count, unmatched_ratio = (
            build_route_distribution(iata_data, oag_data, carrier_name)
        )
    except ValueError as error:
        st.error(str(error))
        return

    if unmatched_ratio >= 0.3:
        st.warning(
            "OAG和IATA航线匹配效果不够理想，建议手动核对。"
            f"未匹配航线数：{unmatched_count}；"
            f"总航线数：{total_count}；"
            f"未匹配比例：{unmatched_ratio:.1%}。"
        )
    else:
        st.caption(
            f"航线匹配结果：{unmatched_count}/{total_count} 条未匹配"
            f"（{unmatched_ratio:.1%}）。"
        )

    st.success("基于日单向客流的航线分布表生成完成。")
    st.dataframe(result, width="stretch")
    st.download_button(
        "下载航线分布表 Excel",
        data=dataframe_to_excel(
            result,
            "Route Distribution",
            include_index=True,
        ),
        file_name="daily_one_way_route_distribution.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


def main() -> None:
    """渲染应用页面并调度四个分析分支。"""
    st.set_page_config(
        page_title="航空市场分析自动化工具",
        layout="wide",
    )
    st.title("航空市场分析自动化工具")
    st.caption(
        "上传合格数据后将保存至项目 data 目录，"
        "应用重启后仍可使用。"
    )

    upload_column_1, upload_column_2, upload_column_3 = st.columns(3)
    with upload_column_1:
        oag_upload = st.file_uploader(
            "导入OAG数据",
            type=["xlsx"],
            key="oag_upload",
        )
    with upload_column_2:
        cirium_upload = st.file_uploader(
            "导入Cirium Trend数据",
            type=["xlsx"],
            key="cirium_upload",
        )
    with upload_column_3:
        iata_upload = st.file_uploader(
            "导入IATA数据",
            type=["xlsx"],
            key="iata_upload",
        )

    oag_valid = process_upload(
        oag_upload,
        "OAG数据表",
        OAG_FILE,
        OAG_COLUMNS,
    )
    cirium_valid = process_upload(
        cirium_upload,
        "Cirium Trend数据表",
        CIRIUM_FILE,
        CIRIUM_COLUMNS,
    )
    iata_valid = process_upload(
        iata_upload,
        "IATA数据表",
        IATA_FILE,
        IATA_COLUMNS,
    )
    if not oag_valid or not cirium_valid or not iata_valid:
        st.stop()

    oag_data = load_persisted_file(OAG_FILE, "OAG数据表", OAG_COLUMNS)
    cirium_data = load_persisted_file(
        CIRIUM_FILE,
        "Cirium Trend数据表",
        CIRIUM_COLUMNS,
    )
    iata_data = load_persisted_file(
        IATA_FILE,
        "IATA数据表",
        IATA_COLUMNS,
    )

    if any(
        dataframe is not None
        for dataframe in [oag_data, cirium_data, iata_data]
    ):
        with st.expander("查看清洗后的标准结构化宽表（前5行）"):
            if oag_data is not None:
                st.subheader("OAG数据")
                st.dataframe(oag_data.head(5), width="stretch")
            if cirium_data is not None:
                st.subheader("Cirium Trend数据")
                st.dataframe(cirium_data.head(5), width="stretch")
            if iata_data is not None:
                st.subheader("IATA数据")
                st.dataframe(iata_data.head(5), width="stretch")

    st.divider()
    option = st.selectbox(
        "请选择分析模块",
        options=[
            OPTION_ROUTE,
            OPTION_FLEET,
            OPTION_DISTRIBUTION,
            OPTION_COUNTRY_FLEET,
        ],
        index=None,
        placeholder="请选择一个分析模块",
    )

    if option == OPTION_ROUTE:
        if oag_data is None:
            st.warning("请先上传并保存 OAG 数据。")
        else:
            render_route_analysis(oag_data)
    elif option == OPTION_FLEET:
        if cirium_data is None:
            st.warning("请先上传并保存 Cirium Trend 数据。")
        else:
            render_fleet_analysis(cirium_data)
    elif option == OPTION_DISTRIBUTION:
        if iata_data is None or oag_data is None:
            st.warning("请先上传并保存 IATA 数据和 OAG 数据。")
        else:
            render_route_distribution(iata_data, oag_data)
    elif option == OPTION_COUNTRY_FLEET:
        if cirium_data is None:
            st.warning("请先上传并保存 Cirium Trend 数据。")
        else:
            render_country_fleet_analysis(cirium_data)


if __name__ == "__main__":
    main()
