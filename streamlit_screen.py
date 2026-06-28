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

# 基础配置
warnings.filterwarnings("ignore")
# 修复云端中文乱码
plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
st.set_page_config(page_title="电商销售数据可视化分析平台", page_icon="📊", layout="wide")
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
                                         feature={"saveAsImage": {"title": "保存为图片", "pixel_ratio": 2},
                                                  "restore": {"title": "重置"},
                                                  "dataView": {"title": "数据视图", "readOnly": False}})
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
        st.error(f"Excel导出失败：{str(e)}")
        return None


def amount_range(val):
    val = float(val)
    if val < 0:
        return "无效金额"
    elif val <= 50:
        return "0-50元"
    elif val <= 100:
        return "50-100元"
    elif val <= 200:
        return "100-200元"
    elif val <= 500:
        return "200-500元"
    else:
        return "500元以上"


def iqr_outlier(series):
    if series.empty or series.nunique() <= 1:
        return pd.Series([False] * len(series)), 0, 0
    q1, q3 = series.quantile([0.25, 0.75])
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr
    outlier_mask = (series < lower_bound) | (series > upper_bound)
    return outlier_mask, lower_bound, upper_bound


@st.cache_data(show_spinner="正在加载并清洗数据...", ttl=3600)
def load_data(uploaded_file):
    logs = []
    try:
        raw = pd.read_excel(uploaded_file, engine="openpyxl")
        logs.append(f"【原始数据】总行数：{raw.shape[0]}，总列数：{raw.shape[1]}")
        df = raw.drop_duplicates()
        dup_count = raw.shape[0] - df.shape[0]
        logs.append(f"【重复值】删除{dup_count}条，剩余{df.shape[0]}行")
        num_cols = df.select_dtypes(include=[np.number]).columns
        obj_cols = df.select_dtypes(include=["object"]).columns
        df[num_cols] = df[num_cols].fillna(0)
        df[obj_cols] = df[obj_cols].fillna("未知")
        logs.append(f"【缺失值】数值列填充0，文本列填充'未知'")
        df["district"] = df["district"].str.strip() if "district" in df.columns else "未知"
        df["省份标准化"] = df["district"]
        logs.append(f"【省份标准化】直接沿用原始行政区名称")

        date_col = "Date" if "Date" in df.columns else ("日期" if "日期" in df.columns else None)
        if date_col:
            df["日期"] = pd.to_datetime(df[date_col], errors="coerce")
            date_err = df["日期"].isna().sum()
            df = df.dropna(subset=["日期"])
            logs.append(f"【日期转换】删除无效日期{date_err}条，剩余{df.shape[0]}行")
        else:
            st.warning("未检测到日期列（Date/日期）")
            df["日期"] = pd.to_datetime("2024-01-01")
            logs.append(f"【日期转换】未找到日期列，默认填充2024-01-01")

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
            logs.append(f"【异常值】阈值[{low:.2f},{high:.2f}]，剔除{mask.sum()}条异常订单")

        df["小时"] = df["日期"].dt.hour
        df["星期名称"] = df["日期"].dt.weekday.map(
            {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'})
        df["金额区间"] = df["买家实际支付金额"].apply(amount_range) if "买家实际支付金额" in df.columns else "0-50元"
        df["退款金额"] = 0
        logs.append(f"【衍生字段】新增小时、星期名称、金额区间、退款金额字段")
        logs.append(f"【清洗完成】最终有效数据：{df.shape[0]}行")
        return raw, df, logs
    except Exception as e:
        logs.append(f"【错误】数据加载失败：{str(e)}")
        st.error(f"数据处理错误：{str(e)}")
        return None, None, logs


# 侧边栏文件上传
st.sidebar.header("📤 数据上传")
uploaded_file = st.sidebar.file_uploader(
    "上传Excel数据文件",
    type=["xlsx"],
    help="字段：Date、district、buy_mount、Total、user_id"
)

if uploaded_file is not None and not state.data_loaded:
    RAW, CLEAN, LOGS = load_data(uploaded_file)
    if RAW is not None and CLEAN is not None:
        state.raw_df, state.clean_df, state.preprocess_log = RAW, CLEAN, LOGS
        state.data_loaded = True
        st.sidebar.success("✅ 数据上传并清洗完成！")
elif uploaded_file is None and not state.data_loaded:
    st.warning("请先在左侧侧边栏上传Excel数据文件！")
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
        st.header("📊 功能导航")
        nav_items = [
            ("🏠 首页总览", "home"),
            ("📋 数据预处理日志", "preprocess"),
            ("📈 基础统计分析", "stat_analysis"),
            ("💰 销售额详情", "sale_detail"),
            ("📦 订单量详情", "order_detail"),
            ("💵 客单价区间分析", "price_detail"),
            ("🗺️ 省份地理分析", "province_detail"),
            ("⏰ 分时时段分析", "hour_detail"),
            ("🔮 时序预测分析(ARIMA)", "forecast")
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
            excel_bytes = export_excel({"筛选数据": filter_df})
            if excel_bytes:
                st.download_button(
                    label="📥 导出当前筛选数据.xlsx",
                    data=excel_bytes,
                    file_name=f"筛选后销售数据_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.download_button(
                label="📥 导出当前筛选数据.xlsx",
                data=b"",
                file_name="筛选后销售数据.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=True
            )
        st.divider()
        with st.expander("🔍 全局筛选", expanded=True):
            st.multiselect(
                "金额区间",
                options=["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"],
                default=state.sel_price_range,
                key="price_key",
                on_change=on_filter_change
            )
            col1, col2 = st.columns(2)
            with col1:
                st.date_input("起始日期", value=state.start_date, min_value=df["日期"].min().date(),
                              max_value=df["日期"].max().date(), key="start_key", on_change=on_filter_change)
            with col2:
                st.date_input("结束日期", value=state.end_date, min_value=df["日期"].min().date(),
                              max_value=df["日期"].max().date(), key="end_key", on_change=on_filter_change)
            prov_options = ["所有省份"] + sorted(df["省份标准化"].unique())
            st.multiselect(
                "省份",
                options=prov_options,
                default=["所有省份"] if set(state.sel_prov) == set(all_prov_list) else state.sel_prov,
                key="prov_key",
                on_change=on_filter_change
            )
            selected_prov = st.session_state.get("prov_key", [])
            if "所有省份" in selected_prov:
                state.sel_prov = all_prov_list
            else:
                state.sel_prov = selected_prov

            if "买家实际支付金额" in df.columns:
                st.slider(
                    "支付金额范围",
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
        st.markdown("# 📊 电商销售数据可视化分析平台 | 综合总览")
        kpi_row = st.columns(4)
        kpi_row[0].metric("筛选总销售额", f"¥{total_sales:,.2f}")
        kpi_row[1].metric("筛选订单总数", f"{total_ord_cnt:,}")
        kpi_row[2].metric("平均客单价", f"¥{avg_price:,.3f}")
        kpi_row[3].metric("覆盖省份数量", f"{prov_count}")
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
            columns={"省份标准化": "省份"})

        week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(
            name="订单量") if "订单编号" in filter_df.columns else filter_df.groupby("星期名称")[
            "日期"].count().reset_index(name="订单量")
        week_stats["排序"] = week_stats["星期名称"].map(
            {"周一": 0, "周二": 1, "周三": 2, "周四": 3, "周五": 4, "周六": 5, "周日": 6})
        week_stats = week_stats.sort_values("排序")

        price_order = ["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"]
        price_stats = filter_df["金额区间"].value_counts().reset_index()
        price_stats.columns = ["金额区间", "订单数"]
        if unique_ord > 0:
            price_stats["占比"] = (price_stats["订单数"] / unique_ord * 100).round(2)
        price_stats["sort_idx"] = price_stats["金额区间"].map(
            lambda x: price_order.index(x) if x in price_order else 99)
        price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

        r1 = st.columns(3)
        with r1[0]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("每日订单量", daily_stats_full["订单量"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("每日订单趋势", min_y=daily_stats_full["订单量"].min()))
            html(l.render_embed(), height=420)
        with r1[1]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("每日销售额", daily_stats_full["销售额"].tolist())
                l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
                l.set_global_opts(**chart_config("日销售额面积趋势图", min_y=daily_stats_full["销售额"].min()))
            html(l.render_embed(), height=420)
        with r1[2]:
            b = Bar(chart_init(460))
            if len(hour_stats_full):
                b.add_xaxis([str(i) for i in range(24)])
                b.add_yaxis("24小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%",
                            label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
                b.set_global_opts(**chart_config("分时订单柱状图", min_y=0, zoom=True))
            html(b.render_embed(), height=460)

        st.divider()
        r2 = st.columns(3)
        with r2[0]:
            m = Map(chart_init(420))
            if not prov_stats.empty:
                map_data = list(zip(prov_stats["省份标准化"], prov_stats["订单量"].astype(int)))
                m.add("订单量分布", map_data, maptype="china", is_map_symbol_show=False)
                cfg = chart_config("", zoom=False)
                cfg["title_opts"] = opts.TitleOpts(title="全国省份订单地图", pos_left="center")
                m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=int(prov_stats["订单量"].max())), **cfg)
            html(m.render_embed(), height=420)
        with r2[1]:
            b = Bar(chart_init(520))
            if len(top15_sales):
                b.add_xaxis(top15_sales["省份"].tolist())
                b.add_yaxis("销售额", top15_sales["销售额"].tolist(), bar_width="70%",
                            label_opts=opts.LabelOpts(is_show=True, position="right"))
                b.reversal_axis()
                b.set_global_opts(**chart_config("TOP15省份销售额横向柱状图", min_x=0, zoom=False))
            html(b.render_embed(), height=520)
        with r2[2]:
            b = Bar(chart_init(420))
            if len(week_stats):
                b.add_xaxis(week_stats["星期名称"].tolist())
                b.add_yaxis("订单量", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
                b.set_global_opts(**chart_config("星期订单分布", min_y=week_stats["订单量"].min(), zoom=False))
            html(b.render_embed(), height=420)

        st.divider()
        r3 = st.columns(3)
        with r3[0]:
            p = Pie(chart_init(420))
            if len(price_stats) > 0:
                p.add("", [list(z) for z in zip(price_stats["金额区间"], price_stats["占比"])], radius=["30%", "70%"])
                p.set_global_opts(**chart_config("客单价区间占比饼图", zoom=False))
            html(p.render_embed(), height=420)
        with r3[1]:
            box = Boxplot(chart_init(440))
            if len(filter_df) > 0 and "买家实际支付金额" in filter_df.columns:
                box.add_xaxis(["全量订单客单价"])
                box.add_yaxis("客单价分布", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
                box.set_series_opts(markpoint_opts=opts.MarkPointOpts(
                    data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
                box.set_global_opts(**chart_config("客单价分布箱线图", min_y=0, zoom=False))
            html(box.render_embed(), height=440)
        with r3[2]:
            l = Line(chart_init(420))
            if len(daily_stats_full):
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("退款金额", daily_stats_full["退款金额"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("每日退款趋势", min_y=daily_stats_full["退款金额"].min()))
            html(l.render_embed(), height=420)

    elif page == "preprocess":
        st.header("📋 数据预处理完整日志")
        for idx, log in enumerate(state.preprocess_log, 1):
            st.info(f"{idx}. {log}")
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("原始数据前10行")
            st.dataframe(state.raw_df.head(10), use_container_width=True, hide_index=True)
        with col2:
            st.subheader("清洗后数据前10行")
            st.dataframe(state.clean_df.head(10), use_container_width=True, hide_index=True)

    elif page == "stat_analysis":
        st.header("📈 基础探索性统计分析")
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

                corr_matrix = filter_df[stat_cols].corr(method="pearson")
                fig, ax = plt.subplots(figsize=(7, 5), dpi=300)
                sns.heatmap(corr_matrix, annot=True, cmap="Blues", vmin=-0.1, vmax=1, ax=ax, fmt=".3f", linewidths=0.5)
                ax.set_title("皮尔逊相关系数热力图", fontsize=10, pad=14)
                ax.set_xticklabels(stat_cols, rotation=0)
                plt.tight_layout()
                buf = io.BytesIO()
                plt.savefig(buf, format="png", bbox_inches="tight", dpi=300)
                buf.seek(0)
                st.image(buf, use_container_width=True)
                plt.close()
            else:
                st.warning("无可用的数值列进行统计分析！")

    elif page == "sale_detail":
        st.header("💰 销售额深度详情")
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
                columns={"省份标准化": "省份"})

            col1, col2 = st.columns(2)
            with col1:
                l = Line(chart_init(480))
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("日销售额", daily_stats_full["销售额"].tolist())
                l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
                l.set_global_opts(**chart_config("日销售额面积趋势图", min_y=daily_stats_full["销售额"].min()))
                html(l.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(520))
                b.add_xaxis(top15_sales["省份"].tolist())
                b.add_yaxis("省份销售额", top15_sales["销售额"].tolist(), bar_width="70%",
                            label_opts=opts.LabelOpts(is_show=True, position="right"))
                b.reversal_axis()
                b.set_global_opts(**chart_config("TOP15省份销售额", min_x=0, zoom=False))
                html(b.render_embed(), height=520)
            st.dataframe(top15_sales.reset_index(drop=True), use_container_width=True, hide_index=True)

    elif page == "order_detail":
        st.header("📦 订单量深度详情")
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
                {"周一": 0, "周二": 1, "周三": 2, "周四": 3, "周五": 4, "周六": 5, "周日": 6})
            week_stats = week_stats.sort_values("排序")

            col1, col2 = st.columns(2)
            with col1:
                l = Line(chart_init(480))
                l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
                l.add_yaxis("日订单数", daily_stats_full["订单量"].tolist(), is_smooth=True)
                l.set_global_opts(**chart_config("每日订单趋势", min_y=daily_stats_full["订单量"].min()))
                html(l.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(480))
                b.add_xaxis(week_stats["星期名称"].tolist())
                b.add_yaxis("周订单量", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
                b.set_global_opts(**chart_config("星期订单分布", min_y=week_stats["订单量"].min(), zoom=False))
                html(b.render_embed(), height=480)

    elif page == "price_detail":
        st.header("💵 客单价&金额区间分析")
        if len(filter_df) > 0:
            price_order = ["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"]
            price_stats = filter_df["金额区间"].value_counts().reset_index()
            price_stats.columns = ["金额区间", "订单数"]
            if unique_ord > 0:
                price_stats["占比"] = (price_stats["订单数"] / unique_ord * 100).round(2)
            price_stats["sort_idx"] = price_stats["金额区间"].map(
                lambda x: price_order.index(x) if x in price_order else 99)
            price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

            p = Pie(chart_init(500))
            p.add("", [list(z) for z in zip(price_stats["金额区间"], price_stats["占比"])], radius=["30%", "70%"])
            p.set_global_opts(**chart_config("价格区间占比饼图", zoom=False))
            html(p.render_embed(), height=500)
            st.dataframe(price_stats, use_container_width=True, hide_index=True)

            st.divider()
            box = Boxplot(chart_init(500))
            if "买家实际支付金额" in filter_df.columns:
                box.add_xaxis(["客单价分布"])
                box.add_yaxis("客单价(元)", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
                box.set_series_opts(markpoint_opts=opts.MarkPointOpts(
                    data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
                box.set_global_opts(**chart_config("客单价分布箱线图", min_y=0, zoom=False))
            html(box.render_embed(), height=500)

    elif page == "province_detail":
        st.header("🗺️ 全国省份地理销售详情")
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
                map_data = list(zip(sub_df["省份标准化"], sub_df["订单量"]))
                m = Map(chart_init(550))
                m.add("订单量", map_data, maptype="china")
                cfg = chart_config("", zoom=False)
                cfg["title_opts"] = opts.TitleOpts(title=f"{ym} 各省份订单分布", pos_left="center")
                m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=sub_df["订单量"].max()), **cfg)
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
        st.header("⏰ 24小时分时深度分析")
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
                b.add_yaxis("每小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%",
                            label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
                b.set_global_opts(**chart_config("分时订单量", min_y=0, zoom=True))
                html(b.render_embed(), height=480)
            with col2:
                b = Bar(chart_init(480))
                b.add_xaxis([str(i) for i in range(24)])
                b.add_yaxis("每小时平均客单价", hour_stats_full["平均订单金额"].tolist(),
                            label_opts=opts.LabelOpts(font_size=10))
                b.set_global_opts(
                    **chart_config("分时平均客单价", min_y=hour_stats_full["平均订单金额"].min(), zoom=True))
                html(b.render_embed(), height=480)

    elif page == "forecast":
        st.header("🔮 ARIMA 时间序列预测分析（加分模块）")
        # 关键容错：没有库就直接提示，程序不会崩溃
        if not HAS_STATSMODELS:
            st.error("⚠️ statsmodels依赖库未安装，暂时无法使用预测功能！请更新requirements.txt后重新部署。")
            st.stop()

        st.info("功能说明：基于每日销售数据，使用ARIMA时序模型做趋势预测")
        if len(filter_df) == 0:
            st.warning("⚠️ 当前筛选无数据，请调整左侧全局筛选条件！")
        else:
            agg_dict = {}
            if "买家实际支付金额" in filter_df.columns:
                agg_dict["销售额"] = ("买家实际支付金额", "sum")
            if "订单编号" in filter_df.columns:
                agg_dict["订单量"] = ("订单编号", "count")
            else:
                agg_dict["订单量"] = ("日期", "count")

            if agg_dict:
                daily_ts = filter_df.groupby(filter_df["日期"].dt.date).agg(**agg_dict).reset_index()
                daily_ts.columns = ["日期"] + list(agg_dict.keys())
                daily_ts["日期"] = pd.to_datetime(daily_ts["日期"])
                daily_ts = daily_ts.sort_values("日期").reset_index(drop=True)

                st.subheader("一、预测参数配置")
                col_cfg1, col_cfg2 = st.columns(2)
                with col_cfg1:
                    target_options = list(agg_dict.keys())
                    target_col = st.radio("选择预测指标", options=target_options,
                                          index=0 if "销售额" in target_options else 0)
                    state.forecast_target = target_col
                with col_cfg2:
                    pred_days = st.selectbox("预测未来天数", options=[7, 15, 30, 60], index=2)
                    state.forecast_period = pred_days
                st.divider()
                run_btn = st.button("🚀 开始训练ARIMA模型并预测", type="primary", use_container_width=True)

                if run_btn:
                    if len(daily_ts) < 20:
                        st.error(f"历史数据仅{len(daily_ts)}天，至少需要20天数据才能建模，请扩大日期筛选范围！")
                    else:
                        with st.spinner("ARIMA模型训练中..."):
                            ts_series = daily_ts.set_index("日期")[target_col]
                            adf_result = adfuller(ts_series)
                            st.info(f"ADF平稳性检验 p值 = {adf_result[1]:.4f}，p<0.05则序列平稳")
                            model = ARIMA(ts_series, order=(1, 1, 1))
                            result = model.fit()
                            forecast_res = result.get_forecast(steps=pred_days)
                            pred_mean = forecast_res.predicted_mean
                            pred_ci = forecast_res.conf_int()

                            last_date = daily_ts["日期"].max()
                            future_dates = [last_date + timedelta(days=i + 1) for i in range(pred_days)]

                            pred_df = pd.DataFrame({
                                "日期": future_dates,
                                "预测值": pred_mean.values,
                                "下限": pred_ci.iloc[:, 0].values,
                                "上限": pred_ci.iloc[:, 1].values
                            })

                            state.arima_model = result
                            state.forecast_result = {
                                "history_df": daily_ts,
                                "pred_df": pred_df
                            }
                            st.success(f"✅ ARIMA(1,1,1)模型训练完成，已预测未来{pred_days}天{target_col}")

                if state.forecast_result is not None:
                    res = state.forecast_result
                    history_df = res["history_df"]
                    pred_df = res["pred_df"]

                    st.divider()
                    st.subheader("二、历史真实值 + ARIMA预测值对比")
                    fig = px.line()
                    fig.add_scatter(x=history_df["日期"], y=history_df[target_col], name="历史真实值")
                    fig.add_scatter(x=pred_df["日期"], y=pred_df["预测值"], name="未来预测值")
                    fig.add_scatter(x=pred_df["日期"], y=pred_df["下限"], fill="tonexty", mode="lines",
                                    line={"dash": "dot"}, name="置信下限")
                    fig.add_scatter(x=pred_df["日期"], y=pred_df["上限"], fill="tonexty", mode="lines",
                                    line={"dash": "dot"}, name="置信上限")
                    fig.update_layout(height=600, title=f"{target_col}时序预测结果")
                    st.plotly_chart(fig, use_container_width=True)
                    st.download_button(
                        label="📥 导出图表HTML",
                        data=pio.to_html(fig),
                        file_name=f"ARIMA_{target_col}_预测图.html",
                        mime="text/html"
                    )

                    st.divider()
                    st.subheader("三、未来预测明细数据")
                    pred_export = pred_df.copy()
                    pred_export["日期"] = pred_export["日期"].dt.date
                    st.dataframe(pred_export, use_container_width=True, hide_index=True)

                    excel_data = {
                        "历史时序数据": history_df,
                        "未来预测明细": pred_export
                    }
                    excel_bytes = export_excel(excel_data)
                    st.download_button(
                        label="📥 下载预测结果Excel",
                        data=excel_bytes,
                        file_name=f"ARIMA_{target_col}_{pred_days}天预测.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.divider()
                    st.subheader("四、原始历史时序数据")
                    st.dataframe(history_df, use_container_width=True, hide_index=True)
                else:
                    st.info("请点击上方【开始训练ARIMA模型并预测】按钮生成预测结果")
            else:
                st.warning("无足够数据列进行时序预测！")