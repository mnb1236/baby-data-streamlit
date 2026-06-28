import os
import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from datetime import date, datetime, timedelta
from pyecharts import options as opts
from pyecharts.charts import Line, Bar, Pie, Map, Boxplot, Timeline
from pyecharts.globals import ThemeType
from streamlit.components.v1 import html
import matplotlib.pyplot as plt
import seaborn as sns
from PIL import Image

# 容错导入statsmodels，避免部署崩溃
try:
    from statsmodels.tsa.arima.model import ARIMA
    from statsmodels.tsa.stattools import adfuller

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False
    ARIMA = None
    adfuller = None
import plotly.express as px
import plotly.io as pio

# ====================== 修复1：全局字体设置，彻底解决中文方块乱码 ======================
warnings.filterwarnings("ignore")
# 强制使用英文无衬线字体，云端不会出现方块乱码
plt.rcParams["font.family"] = ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
# pyecharts全局不依赖本地字体，文字自动渲染正常

st.set_page_config(page_title="E-commerce Sales Analysis Platform", page_icon="📊", layout="wide")
st.markdown("""
<style>
.main {background-color: #f7f9fc;}
.block-container {padding: 2rem;}
.stMetric {background: white; border-radius: 10px; padding:10px; box-shadow: 0 2px 8px #eee;}
div.stButton > button:first-child {background-color: #2563eb; color:white; border-radius:6px; font-size:15px;}
div.stDownloadButton > button {background-color: #059669; color:white; font-size:15px;}
[data-testid="stSidebar"] {background-color:#ffffff;}
iframe {width:100% !important;}
.stExpander {border-radius: 8px;}
.stDataFrame {border-radius: 8px; overflow: hidden;}
</style>
""", unsafe_allow_html=True)


# 会话状态初始化
def init_session_state():
    default_states = {
        "page": "home",
        "raw_df": None,
        "clean_df": None,
        "preprocess_log": [],
        "sel_price_range": ["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"],
        "start_date": None,
        "end_date": None,
        "sel_prov": ["所有省份"],
        "min_p": 12.0,
        "max_p": 99999.0,
        "filter_df_cache": None,
        "cache_timestamp": None,
        "arima_model": None,
        "forecast_result": None,
        "forecast_target": "销售额",
        "forecast_period": 30,
        "data_loaded": False
    }
    for key, val in default_states.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_session_state()
state = st.session_state


def on_filter_change():
    if "price_key" in st.session_state:
        state.sel_price_range = st.session_state["price_key"]
    if "start_key" in st.session_state:
        state.start_date = st.session_state["start_key"]
    if "end_key" in st.session_state:
        state.end_date = st.session_state["end_key"]
    if "prov_key" in st.session_state:
        state.sel_prov = st.session_state["prov_key"]
    if "slider_key" in st.session_state:
        state.min_p, state.max_p = st.session_state["slider_key"]
    state.filter_df_cache = None
    state.cache_timestamp = datetime.now()
    state.arima_model = None
    state.forecast_result = None


def chart_init(height=480, theme=ThemeType.MACARONS):
    return opts.InitOpts(width="100%", height=f"{height}px", bg_color="#ffffff", theme=theme, renderer="canvas")


def chart_config(title_name, min_y=None, max_y=None, min_x=None, max_x=None, zoom=True):
    cfg = {
        "title_opts": opts.TitleOpts(title=title_name, pos_left="center",
                                     title_textstyle_opts=opts.TextStyleOpts(font_size=16, font_weight="bold")),
        "xaxis_opts": opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=12, rotate=15),
                                    splitline_opts=opts.SplitLineOpts(is_show=False), min_=min_x, max_=max_x),
        "yaxis_opts": opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=12),
                                    splitline_opts=opts.SplitLineOpts(is_show=True,
                                                                      linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
                                    min_=min_y, max_=max_y),
        "legend_opts": opts.LegendOpts(pos_bottom="2%", textstyle_opts=opts.TextStyleOpts(font_size=12),
                                       orient="horizontal"),
        "tooltip_opts": opts.TooltipOpts(trigger="axis", axis_pointer_type="shadow",
                                         textstyle_opts=opts.TextStyleOpts(font_size=11)),
        "toolbox_opts": opts.ToolboxOpts(is_show=True,
                                         feature={"saveAsImage": {"title": "Save Image", "pixel_ratio": 2},
                                                  "restore": {"title": "Reset"},
                                                  "dataView": {"title": "Data View", "readOnly": False}})
    }
    if zoom:
        cfg["datazoom_opts"] = [opts.DataZoomOpts(range_start=0, range_end=100, orient="horizontal"),
                                opts.DataZoomOpts(range_start=0, range_end=100, orient="vertical")]
    return cfg


def export_excel(sheet_dict):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for name, data in sheet_dict.items():
                if not data.empty:
                    data.to_excel(writer, sheet_name=name[:31], index=False)
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.error(f"Excel Export Failed: {str(e)}")
        return None


def amount_range(val):
    val = float(val)
    if val < 0:
        return "Invalid"
    elif val <= 50:
        return "0-50"
    elif val <= 100:
        return "50-100"
    elif val <= 200:
        return "100-200"
    elif val <= 500:
        return "200-500"
    else:
        return ">500"


def iqr_outlier(series):
    if series.empty or series.nunique() <= 1:
        return pd.Series([False] * len(series)), 0, 0
    q1, q3 = series.quantile([0.25, 0.75])
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr
    outlier_mask = (series < lower_bound) | (series > upper_bound)
    return outlier_mask, lower_bound, upper_bound


@st.cache_data(show_spinner="Loading & Cleaning Data...", ttl=3600)
def load_data(uploaded_file):
    logs = []
    try:
        raw = pd.read_excel(uploaded_file, engine="openpyxl")
        logs.append(f"【Raw Data】Rows: {raw.shape[0]}, Columns: {raw.shape[1]}")
        df = raw.drop_duplicates()
        dup_count = raw.shape[0] - df.shape[0]
        logs.append(f"【Duplicates】Removed {dup_count} rows, remaining {df.shape[0]}")
        num_cols = df.select_dtypes(include=[np.number]).columns
        obj_cols = df.select_dtypes(include=["object"]).columns
        df[num_cols] = df[num_cols].fillna(0)
        df[obj_cols] = df[obj_cols].fillna("Unknown")
        logs.append(f"【Missing Value】Numeric filled with 0, Text filled with Unknown")
        df["district"] = df["district"].str.strip() if "district" in df.columns else "Unknown"
        df["省份标准化"] = df["district"]
        logs.append(f"【Province】Clean district name")

        date_col = "Date" if "Date" in df.columns else ("日期" if "日期" in df.columns else None)
        if date_col:
            df["日期"] = pd.to_datetime(df[date_col], errors="coerce")
            date_err = df["日期"].isna().sum()
            df = df.dropna(subset=["日期"])
            logs.append(f"【Date】Removed {date_err} invalid date rows, remaining {df.shape[0]}")
        else:
            st.warning("Date column not found")
            df["日期"] = pd.to_datetime("2024-01-01")
            logs.append(f"【Date】Use default date 2024-01-01")

        col_mapping = {
            "buy_mount": "购买数量",
            "Total": "买家实际支付金额",
            "user_id": "订单编号"
        }
        for old_col, new_col in col_mapping.items():
            if old_col in df.columns:
                df.rename(columns={old_col: new_col}, inplace=True)

        if "买家实际支付金额" in df.columns and not df["买家实际支付金额"].empty:
            mask, low, high = iqr_outlier(df["买家实际支付金额"])
            df = df[~mask]
            logs.append(f"【Outlier】Threshold [{low:.2f},{high:.2f}], removed {mask.sum()} abnormal orders")

        df["小时"] = df["日期"].dt.hour
        df["星期名称"] = df["日期"].dt.weekday.map(
            {0: 'Mon', 1: 'Tue', 2: 'Wed', 3: 'Thu', 4: 'Fri', 5: 'Sat', 6: 'Sun'})
        df["金额区间"] = df["买家实际支付金额"].apply(amount_range) if "买家实际支付金额" in df.columns else "0-50"
        df["退款金额"] = 0
        logs.append(f"【New Columns】Add hour, weekday, price range, refund")
        logs.append(f"【Clean Finish】Final valid rows: {df.shape[0]}")
        return raw, df, logs
    except Exception as e:
        logs.append(f"【Error】Load failed: {str(e)}")
        st.error(f"Data Error: {str(e)}")
        return None, None, logs


# 侧边栏文件上传
st.sidebar.header("📤 Upload Data")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel File",
    type=["xlsx"],
    help="Required Columns: Date, district, buy_mount, Total, user_id"
)

if uploaded_file is not None and not state.data_loaded:
    RAW, CLEAN, LOGS = load_data(uploaded_file)
    if RAW is not None and CLEAN is not None:
        state.raw_df, state.clean_df, state.preprocess_log = RAW, CLEAN, LOGS
        state.data_loaded = True
        st.sidebar.success("✅ Data loaded successfully!")
elif uploaded_file is None and not state.data_loaded:
    st.warning("Please upload Excel file in the sidebar first!")
    st.stop()

df = state.clean_df
if state.data_loaded:
    if state.start_date is None:
        state.start_date = df["日期"].min().date()
    if state.end_date is None:
        state.end_date = df["日期"].max().date()
    all_prov_list = sorted(df["省份标准化"].unique())
    if len(state.sel_prov) == 1 and state.sel_prov[0] == "所有省份":
        state.sel_prov = all_prov_list

    if "买家实际支付金额" in df.columns:
        max_val = float(df["买家实际支付金额"].max())
    else:
        max_val = 99999.0
    min_slider_fixed = 12.0
    if max_val <= min_slider_fixed:
        max_val = min_slider_fixed + 0.01

    if state.min_p < min_slider_fixed:
        state.min_p = min_slider_fixed
    if state.max_p > max_val:
        state.max_p = max_val


    def get_filtered_df():
        if state.filter_df_cache is not None:
            return state.filter_df_cache
        filter_conditions = [
            df["金额区间"].isin(state.sel_price_range),
            df["日期"].dt.date >= state.start_date,
            df["日期"].dt.date <= state.end_date,
            df["省份标准化"].isin(state.sel_prov),
        ]
        if "买家实际支付金额" in df.columns:
            filter_conditions.extend([
                df["买家实际支付金额"] >= state.min_p,
                df["买家实际支付金额"] <= state.max_p
            ])
        filter_df = df[np.all(filter_conditions, axis=0)].copy()
        state.filter_df_cache = filter_df
        return filter_df


    filter_df = get_filtered_df()

    with st.sidebar:
        st.header("📊 Navigation")
        nav_items = [
            ("🏠 Home Overview", "home"),
            ("📋 Preprocess Log", "preprocess"),
            ("📈 Statistical Analysis", "stat_analysis"),
            ("💰 Sales Detail", "sale_detail"),
            ("📦 Order Detail", "order_detail"),
            ("💵 Price Range Analysis", "price_detail"),
            ("🗺️ Provincial Map", "province_detail"),
            ("⏰ Hourly Analysis", "hour_detail"),
            ("🔮 ARIMA Forecast", "forecast")
        ]
        for txt, key in nav_items:
            btn_kwargs = {"use_container_width": True, "key": f"nav_{key}"}
            if state.page == key:
                btn_kwargs["type"] = "primary"
            if st.button(txt, **btn_kwargs):
                state.page = key
                st.rerun()
        st.divider()
        if not filter_df.empty:
            excel_bytes = export_excel({"Filtered Data": filter_df})
            if excel_bytes:
                st.download_button(
                    label="📥 Export Excel",
                    data=excel_bytes,
                    file_name=f"Filtered_Data_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.download_button(
                label="📥 Export Excel",
                data=b"",
                file_name="Filtered_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=True
            )
        st.divider()
        with st.expander("🔍 Global Filter", expanded=True):
            st.multiselect(
                "Price Range",
                options=["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"],
                default=state.sel_price_range,
                key="price_key",
                on_change=on_filter_change
            )
            col1, col2 = st.columns(2)
            with col1:
                st.date_input("Start Date", value=state.start_date, min_value=df["日期"].min().date(),
                              max_value=df["日期"].max().date(), key="start_key", on_change=on_filter_change)
            with col2:
                st.date_input("End Date", value=state.end_date, min_value=df["日期"].min().date(),
                              max_value=df["日期"].max().date(), key="end_key", on_change=on_filter_change)
            prov_options = ["All Provinces"] + sorted(df["省份标准化"].unique())
            st.multiselect(
                "Province",
                options=prov_options,
                default=["All Provinces"] if set(state.sel_prov) == set(all_prov_list) else state.sel_prov,
                key="prov_key",
                on_change=on_filter_change
            )
            selected_prov = st.session_state.get("prov_key", [])
            if "All Provinces" in selected_prov:
                state.sel_prov = all_prov_list
            else:
                state.sel_prov = selected_prov

            if "买家实际支付金额" in df.columns:
                st.slider(
                    "Payment Amount",
                    min_value=min_slider_fixed,
                    max_value=max_val,
                    value=(state.min_p, state.max_p),
                    key="slider_key",
                    on_change=on_filter_change
                )

    start_date = state.start_date
    end_date = state.end_date
    total_sales = filter_df["买家实际支付金额"].sum() if (
                "买家实际支付金额" in filter_df.columns and not filter_df.empty) else 0
    total_ord_cnt = len(filter_df)
    unique_ord = filter_df["订单编号"].nunique() if ("订单编号" in filter_df.columns and not filter_df.empty) else 0
    avg_price = round(total_sales / unique_ord, 3) if unique_ord > 0 else 0
    prov_count = filter_df["省份标准化"].nunique() if not filter_df.empty else 0
    page = state.page

    if page == "home":
        st.markdown("# 📊 E-commerce Sales Analysis Platform | Overview")
        kpi_row = st.columns(4)
        kpi_row[0].metric("Total Sales", f"¥{total_sales:,.2f}")
        kpi_row[1].metric("Total Orders", f"{total_ord_cnt:,}")
        kpi_row[2].metric("Avg Order Price", f"¥{avg_price:,.3f}")
        kpi_row[3].metric("Provinces Covered", f"{prov_count}")
        st.divider()

        daily_stats = filter_df.groupby("日期").agg(
            订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("日期", "count"),
            销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else ("日期", "count"),
            退款金额=("退款金额", "sum") if "退款金额" in filter_df.columns else ("日期", "count")
        ).reset_index()
        daily_stats_full = pd.merge(pd.DataFrame({"日期": pd.date_range(start_date, end_date)}), daily_stats,
                                    how="left").fillna(0)

        hour_stats = filter_df.groupby("小时").agg(
            订单量=("买家实际支付金额", "count") if "买家实际支付金额" in filter_df.columns else ("小时", "count"),
            平均订单金额=(
            "买家实际支付金额", lambda x: round(x.mean(), 3)) if "买家实际支付金额" in filter_df.columns else (
            "小时", lambda x: 0)
        ).reset_index()
        hour_stats_full = pd.merge(pd.DataFrame({"小时": range(24)}), hour_stats, how="left").fillna(0)

        prov_stats = filter_df.groupby("省份标准化").agg(
            订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("省份标准化", "count"),
            销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else ("省份标准化", "count")
        ).reset_index()
        prov_stats = prov_stats[prov_stats["订单量"] > 0].copy()
        top15_sales = prov_stats.sort_values(by="销售额", ascending=False).head(15).rename(
            columns={"省份标准化": "Province"})

        week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(
            name="订单量") if "订单编号" in filter_df.columns else filter_df.groupby("星期名称")[
            "日期"].count().reset_index(name="订单量")
        week_stats["排序"] = week_stats["星期名称"].map(
            {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5, "Sun": 6})
        week_stats = week_stats.sort_values("排序")

        price_order = ["0-50", "50-100", "100-200", "200-500", ">500"]
        price_stats = filter_df["金额区间"].value_counts().reset_index()
        price_stats.columns = ["Price Range", "Order Count"]
        if unique_ord > 0:
            price_stats["Ratio(%)"] = (price_stats["Order Count"] / unique_ord * 100).round(2)
        price_stats["sort_idx"] = price_stats["Price Range"].map(
            lambda x: price_order.index(x) if x in price_order else 99)
        price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

        r1 = st.columns(3)
        with r1[0]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("Daily Orders", daily_stats_full["订单量"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("Daily Order Trend", min_y=daily_stats_full["订单量"].min()))
            html(l.render_embed(), height=420)
        with r1[1]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("Daily Sales", daily_stats_full["销售额"].tolist())
                l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
                l.set_global_opts(**chart_config("Daily Sales Trend", min_y=daily_stats_full["销售额"].min()))
            html(l.render_embed(), height=420)
        with r1[2]:
            b = Bar(chart_init(460))
            if len(hour_stats_full):
                b.add_xaxis([str(i) for i in range(24)])
                b.add_yaxis("24H Orders", hour_stats_full["订单量"].tolist(), bar_width="60%",
                            label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
                b.set_global_opts(**chart_config("Hourly Order Bar", min_y=0, zoom=True))
            html(b.render_embed(), height=460)

        st.divider()
        r2 = st.columns(3)
        with r2[0]:
            m = Map(chart_init(420))
            if not prov_stats.empty:
                # ====================== 修复2：地图数据强制转为int，保证颜色梯度正常显示 ======================
                map_data = list(zip(prov_stats["省份标准化"], prov_stats["订单量"].astype(int)))
                m.add("Order Count", map_data, maptype="china", is_map_symbol_show=False)
                cfg = chart_config("", zoom=False)
                cfg["title_opts"] = opts.TitleOpts(title="National Provincial Order Map", pos_left="center")
                m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=int(prov_stats["订单量"].max())), **cfg)
            html(m.render_embed(), height=420)
        with r2[1]:
            b = Bar(chart_init(520))
            if len(top15_sales):
                b.add_xaxis(top15_sales["Province"].tolist())
                b.add_yaxis("Sales", top15_sales["销售额"].tolist(), bar_width="70%",
                            label_opts=opts.LabelOpts(is_show=True, position="right"))
                b.reversal_axis()
                b.set_global_opts(**chart_config("TOP15 Provincial Sales", min_x=0, zoom=False))
            html(b.render_embed(), height=520)
        with r2[2]:
            b = Bar(chart_init(420))
            if len(week_stats):
                b.add_xaxis(week_stats["星期名称"].tolist())
                b.add_yaxis("Order Count", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
                b.set_global_opts(
                    **chart_config("Weekly Order Distribution", min_y=week_stats["订单量"].min(), zoom=False))
            html(b.render_embed(), height=420)

        st.divider()
        r3 = st.columns(3)
        with r3[0]:
            p = Pie(chart_init(420))
            if len(price_stats) > 0:
                p.add("", [list(z) for z in zip(price_stats["Price Range"], price_stats["Ratio(%)"])],
                      radius=["30%", "70%"])
                p.set_global_opts(**chart_config("Price Range Pie Chart", zoom=False))
            html(p.render_embed(), height=420)
        with r3[1]:
            box = Boxplot(chart_init(440))
            if len(filter_df) > 0 and "买家实际支付金额" in filter_df.columns:
                box.add_xaxis(["Order Amount Distribution"])
                box.add_yaxis("Order Amount", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
                box.set_series_opts(markpoint_opts=opts.MarkPointOpts(
                    data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
                box.set_global_opts(**chart_config("Order Amount Boxplot", min_y=0, zoom=False))
            html(box.render_embed(), height=440)
        with r3[2]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("Refund Amount", daily_stats_full["退款金额"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("Daily Refund Trend", min_y=daily_stats_full["退款金额"].min()))
            html(l.render_embed(), height=420)

    elif page == "preprocess":
        st.header("📋 Data Preprocess Log")
        for idx, log in enumerate(state.preprocess_log, 1):
            st.info(f"{idx}. {log}")
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Raw Data Top 10 Rows")
            st.dataframe(state.raw_df.head(10), use_container_width=True, hide_index=True)
        with col2:
            st.subheader("Cleaned Data Top 10 Rows")
            st.dataframe(state.clean_df.head(10), use_container_width=True, hide_index=True)

    elif page == "stat_analysis":
        st.header("📈 Exploratory Statistical Analysis")
        if len(filter_df) > 0:
            stat_cols = []
            if "购买数量" in filter_df.columns:
                stat_cols.append("购买数量")
            if "买家实际支付金额" in filter_df.columns:
                stat_cols.append("买家实际支付金额")
            if "小时" in filter_df.columns:
                stat_cols.append("小时")

            if stat_cols:
                desc_df = filter_df[stat_cols].describe().round(3)
                st.dataframe(desc_df, use_container_width=True, hide_index=True)
                st.divider()

                # ====================== 修复3：热力图英文坐标轴，彻底杜绝方块乱码 ======================
                corr_matrix = filter_df[stat_cols].corr(method="pearson")
                # 把列名替换为英文，不会出现中文方块
                corr_matrix.columns = ["Purchase Qty", "Payment Amount", "Hour"]
                corr_matrix.index = ["Purchase Qty", "Payment Amount", "Hour"]

                fig, ax = plt.subplots(figsize=(7, 5), dpi=300)
                sns.heatmap(corr_matrix, annot=True, cmap="Blues", vmin=-0.1, vmax=1, ax=ax, fmt=".3f", linewidths=0.5)
                ax.set_title("Pearson Correlation Heatmap", fontsize=14, pad=14)
                plt.tight_layout()
                buf = io.BytesIO()
                plt.savefig(buf, format="png", bbox_inches="tight", dpi=300)
                buf.seek(0)
                st.image(buf, use_container_width=True)
                plt.close()
            else:
                st.warning("No numeric columns for correlation analysis!")

    elif page == "sale_detail":
        st.header("💰 Sales Deep Analysis")
        if len(filter_df) > 0:
            daily_stats = filter_df.groupby("日期").agg(
                订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("日期", "count"),
                销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else ("日期", "count"),
                退款金额=("退款金额", "sum") if "退款金额" in filter_df.columns else ("日期", "count")
            ).reset_index()
            daily_stats_full = pd.merge(pd.DataFrame({"日期": pd.date_range(start_date, end_date)}), daily_stats,
                                        how="left").fillna(0)

            prov_stats = filter_df.groupby("省份标准化").agg(
                订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("省份标准化", "count"),
                销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else (
                "省份标准化", "count")
            ).reset_index()
            prov_stats = prov_stats[prov_stats["订单量"] > 0]
            top15_sales = prov_stats.sort_values(by="销售额", ascending=False).head(15).rename(
                columns={"省份标准化": "Province"})

            col1, col2 = st.columns(2)
            with col1:
                l = Line(chart_init(480))
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("Daily Sales", daily_stats_full["销售额"].tolist())
                l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
                l.set_global_opts(**chart_config("Daily Sales Trend", min_y=daily_stats_full["销售额"].min()))
                html(l.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(520))
                b.add_xaxis(top15_sales["Province"].tolist())
                b.add_yaxis("Provincial Sales", top15_sales["销售额"].tolist(), bar_width="70%",
                            label_opts=opts.LabelOpts(is_show=True, position="right"))
                b.reversal_axis()
                b.set_global_opts(**chart_config("TOP15 Provincial Sales", min_x=0, zoom=False))
            html(b.render_embed(), height=520)
            st.dataframe(top15_sales.reset_index(drop=True), use_container_width=True, hide_index=True)

    elif page == "order_detail":
        st.header("📦 Order Volume Analysis")
        if len(filter_df) > 0:
            daily_stats = filter_df.groupby("日期").agg(
                订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("日期", "count"),
                销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else ("日期", "count"),
                退款金额=("退款金额", "sum") if "退款金额" in filter_df.columns else ("日期", "count")
            ).reset_index()
            daily_stats_full = pd.merge(pd.DataFrame({"日期": pd.date_range(start_date, end_date)}), daily_stats,
                                        how="left").fillna(0)

            week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(
                name="订单量") if "订单编号" in filter_df.columns else filter_df.groupby("星期名称")[
                "日期"].count().reset_index(name="订单量")
            week_stats["排序"] = week_stats["星期名称"].map(
                {"Mon": 0, "Tue": 1, "Wed": 2, "Thu": 3, "Fri": 4, "Sat": 5, "Sun": 6})
            week_stats = week_stats.sort_values("排序")

            col1, col2 = st.columns(2)
            with col1:
                l = Line(chart_init(480))
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("Daily Orders", daily_stats_full["订单量"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("Daily Order Trend", min_y=daily_stats_full["订单量"].min()))
                html(l.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(480))
                b.add_xaxis(week_stats["星期名称"].tolist())
                b.add_yaxis("Weekly Orders", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
                b.set_global_opts(
                    **chart_config("Weekly Order Distribution", min_y=week_stats["订单量"].min(), zoom=False))
                html(b.render_embed(), height=480)

    elif page == "price_detail":
        st.header("💵 Order Price Range Analysis")
        if len(filter_df) > 0:
            price_order = ["0-50", "50-100", "100-200", "200-500", ">500"]
            price_stats = filter_df["金额区间"].value_counts().reset_index()
            price_stats.columns = ["Price Range", "Order Count"]
            if unique_ord > 0:
                price_stats["Ratio(%)"] = (price_stats["Order Count"] / unique_ord * 100).round(2)
            price_stats["sort_idx"] = price_stats["Price Range"].map(
                lambda x: price_order.index(x) if x in price_order else 99)
            price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

            p = Pie(chart_init(500))
            p.add("", [list(z) for z in zip(price_stats["Price Range"], price_stats["Ratio(%)"])],
                  radius=["30%", "70%"])
            p.set_global_opts(**chart_config("Price Range Pie Chart", zoom=False))
            html(p.render_embed(), height=500)
            st.dataframe(price_stats, use_container_width=True, hide_index=True)

            st.divider()
            box = Boxplot(chart_init(500))
            if "买家实际支付金额" in filter_df.columns:
                box.add_xaxis(["Order Amount Distribution"])
                box.add_yaxis("Order Amount(¥)", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
                box.set_series_opts(markpoint_opts=opts.MarkPointOpts(
                    data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
                box.set_global_opts(**chart_config("Order Amount Boxplot", min_y=0, zoom=False))
            html(box.render_embed(), height=500)

    elif page == "province_detail":
        st.header("🗺️ Provincial Geographic Sales Analysis")
        if len(filter_df) > 0:
            filter_df["年月"] = filter_df["日期"].dt.to_period("M")
            month_group = filter_df.groupby(["年月", "省份标准化"])["订单编号"].count().reset_index(
                name="订单量") if "订单编号" in filter_df.columns else filter_df.groupby(["年月", "省份标准化"])[
                "日期"].count().reset_index(name="订单量")
            month_group = month_group[month_group["订单量"] > 0]

            tl = Timeline(chart_init(550))
            tl.add_schema(play_interval=1000, is_auto_play=False, is_loop_play=False, pos_bottom="5%",
                          label_opts=opts.LabelOpts(font_size=12))
            for ym in sorted(month_group["年月"].unique()):
                sub_df = month_group[month_group["年月"] == ym]
                # 再次强制转为int，保证地图色块正常渲染，不会全黄无梯度
                map_data = list(zip(sub_df["省份标准化"], sub_df["订单量"].astype(int)))
                m = Map(chart_init(550))
                m.add("Order Count", map_data, maptype="china")
                cfg = chart_config("", zoom=False)
                cfg["title_opts"] = opts.TitleOpts(title=f"{ym} Provincial Order Distribution", pos_left="center")
                # 动态最大值，保证每个月份颜色梯度正确
                max_value = int(sub_df["订单量"].max())
                m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=max_value), **cfg)
                tl.add(m, str(ym))
            html(tl.render_embed(), height=580)

            prov_stats = filter_df.groupby("省份标准化").agg(
                订单量=("订单编号", "count") if "订单编号" in filter_df.columns else ("省份标准化", "count"),
                销售额=("买家实际支付金额", "sum") if "买家实际支付金额" in filter_df.columns else (
                "省份标准化", "count")
            ).reset_index()
            prov_stats = prov_stats[prov_stats["订单量"] > 0]
            st.dataframe(prov_stats.sort_values("销售额", ascending=False).reset_index(drop=True),
                         use_container_width=True, hide_index=True)

    elif page == "hour_detail":
        st.header("⏰ 24-Hour Order Analysis")
        if len(filter_df) > 0:
            hour_stats = filter_df.groupby("小时").agg(
                订单量=("买家实际支付金额", "count") if "买家实际支付金额" in filter_df.columns else ("小时", "count"),
                平均订单金额=(
                "买家实际支付金额", lambda x: round(x.mean(), 3)) if "买家实际支付金额" in filter_df.columns else (
                "小时", lambda x: 0)
            ).reset_index()
            hour_stats_full = pd.merge(pd.DataFrame({"小时": range(24)}), hour_stats, how="left").fillna(0)

            col1, col2 = st.columns(2)
            with col1:
                b = Bar(chart_init(480))
                b.add_xaxis([str(i) for i in range(24)])
                b.add_yaxis("Hourly Orders", hour_stats_full["订单量"].tolist(), bar_width="60%",
                            label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
                b.set_global_opts(**chart_config("Hourly Order Count", min_y=0, zoom=True))
                html(b.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(480))
                b.add_xaxis([str(i) for i in range(24)])
                b.add_yaxis("Avg Order Amount", hour_stats_full["平均订单金额"].tolist(),
                            label_opts=opts.LabelOpts(font_size=10))
                b.set_global_opts(
                    **chart_config("Hourly Avg Amount", min_y=hour_stats_full["平均订单金额"].min(), zoom=True))
                html(b.render_embed(), height=480)

    elif page == "forecast":
        st.header("🔮 ARIMA Time Series Forecasting")
        if not HAS_STATSMODELS:
            st.error("⚠️ statsmodels library not installed, forecasting module unavailable!")
            st.stop()

        st.info("Forecast future sales & order volume with ARIMA model")
        if len(filter_df) == 0:
            st.warning("⚠️ No filtered data, please adjust filter conditions!")
        else:
            agg_dict = {}
            if "买家实际支付金额" in filter_df.columns:
                agg_dict["Sales"] = ("买家实际支付金额", "sum")
            if "订单编号" in filter_df.columns:
                agg_dict["Orders"] = ("订单编号", "count")
            else:
                agg_dict["Orders"] = ("日期", "count")

            if agg_dict:
                daily_ts = filter_df.groupby(filter_df["日期"].dt.date).agg(**agg_dict).reset_index()
                daily_ts.columns = ["Date"] + list(agg_dict.keys())
                daily_ts["Date"] = pd.to_datetime(daily_ts["Date"])
                daily_ts = daily_ts.sort_values("Date").reset_index(drop=True)

                st.subheader("1. Forecast Parameter Setting")
                col_cfg1, col_cfg2 = st.columns(2)
                with col_cfg1:
                    target_options = list(agg_dict.keys())
                    target_col = st.radio("Select Target", options=target_options, index=0)
                    state.forecast_target = target_col
                with col_cfg2:
                    pred_days = st.selectbox("Forecast Days", options=[7, 15, 30, 60], index=2)
                    state.forecast_period = pred_days
                st.divider()
                run_btn = st.button("🚀 Train ARIMA & Predict", type="primary", use_container_width=True)

                if run_btn:
                    if len(daily_ts) < 20:
                        st.error(f"Only {len(daily_ts)} days history, need at least 20 days!")
                    else:
                        with st.spinner("Training ARIMA model..."):
                            ts_series = daily_ts.set_index("Date")[target_col]
                            adf_result = adfuller(ts_series)
                            st.info(f"ADF P-value = {adf_result[1]:.4f}, P<0.05 means stationary")
                            model = ARIMA(ts_series, order=(1, 1, 1))
                            result = model.fit()
                            forecast_res = result.get_forecast(steps=pred_days)
                            pred_mean = forecast_res.predicted_mean
                            pred_ci = forecast_res.conf_int()

                            last_date = daily_ts["Date"].max()
                            future_dates = [last_date + timedelta(days=i + 1) for i in range(pred_days)]

                            pred_df = pd.DataFrame({
                                "Date": future_dates,
                                "Forecast": pred_mean.values,
                                "Lower Bound": pred_ci.iloc[:, 0].values,
                                "Upper Bound": pred_ci.iloc[:, 1].values
                            })

                            state.arima_model = result
                            state.forecast_result = {
                                "history_df": daily_ts,
                                "pred_df": pred_df
                            }
                            st.success(f"✅ ARIMA(1,1,1) finished, predicted next {pred_days} days {target_col}")

                if state.forecast_result is not None:
                    res = state.forecast_result
                    history_df = res["history_df"]
                    pred_df = res["pred_df"]

                    st.divider()
                    st.subheader("2. History + Forecast Chart")
                    fig = px.line()
                    fig.add_scatter(x=history_df["Date"], y=history_df[target_col], name="Actual History")
                    fig.add_scatter(x=pred_df["Date"], y=pred_df["Forecast"], name="Future Forecast")
                    fig.add_scatter(x=pred_df["Date"], y=pred_df["Lower Bound"], fill="tonexty", mode="lines",
                                    line={"dash": "dot"}, name="95% CI Lower")
                    fig.add_scatter(x=pred_df["Date"], y=pred_df["Upper Bound"], fill="tonexty", mode="lines",
                                    line={"dash": "dot"}, name="95% CI Upper")
                    fig.update_layout(height=600, title=f"{target_col} Time Series Forecast")
                    st.plotly_chart(fig, use_container_width=True)
                    st.download_button(
                        label="📥 Export Chart HTML",
                        data=pio.to_html(fig),
                        file_name=f"ARIMA_{target_col}_Forecast.html",
                        mime="text/html"
                    )

                    st.divider()
                    st.subheader("3. Future Forecast Detail")
                    pred_export = pred_df.copy()
                    pred_export["Date"] = pred_export["Date"].dt.date
                    st.dataframe(pred_export, use_container_width=True, hide_index=True)

                    excel_data = {
                        "History Data": history_df,
                        "Forecast Detail": pred_export
                    }
                    excel_bytes = export_excel(excel_data)
                    st.download_button(
                        label="📥 Download Forecast Excel",
                        data=excel_bytes,
                        file_name=f"ARIMA_{target_col}_{pred_days}days.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.divider()
                    st.subheader("4. Raw Time Series Data")
                    st.dataframe(history_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Click the button above to generate forecast result")
            else:
                st.warning("No enough time series data for ARIMA!")