import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from datetime import date, datetime
from pyecharts import options as opts
from pyecharts.charts import Line, Bar, Pie, Map, Boxplot, Timeline
from pyecharts.globals import ThemeType
from streamlit.components.v1 import html
import matplotlib.pyplot as plt
import seaborn as sns
from PIL import Image
import os

# 基础配置
warnings.filterwarnings("ignore")
plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "SimHei", "WenQuanYi Zen Hei"]  # 适配云端字体
plt.rcParams["axes.unicode_minus"] = False

st.set_page_config(page_title="母婴电商销售数据可视化分析平台", page_icon="📊", layout="wide")

# 自定义样式优化
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
        "min_p": 0.0,
        "max_p": 99999.0,
        "filter_df_cache": None,
        "cache_timestamp": None
    }
    for key, val in default_states.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_session_state()
state = st.session_state


# 全局筛选变更处理
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


# 图表初始化
def chart_init(height=480, theme=ThemeType.MACARONS):
    return opts.InitOpts(
        width="100%",
        height=f"{height}px",
        bg_color="#ffffff",
        theme=theme,
        renderer="canvas"
    )


# 图表全局配置
def chart_config(title_name, min_y=None, max_y=None, min_x=None, max_x=None, zoom=True):
    cfg = {
        "title_opts": opts.TitleOpts(
            title=title_name,
            pos_left="center",
            title_textstyle_opts=opts.TextStyleOpts(font_size=16, font_weight="bold")
        ),
        "xaxis_opts": opts.AxisOpts(
            axislabel_opts=opts.LabelOpts(font_size=12, rotate=15),
            splitline_opts=opts.SplitLineOpts(is_show=False),
            min_=min_x,
            max_=max_x
        ),
        "yaxis_opts": opts.AxisOpts(
            axislabel_opts=opts.LabelOpts(font_size=12),
            splitline_opts=opts.SplitLineOpts(is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=0.3)),
            min_=min_y,
            max_=max_y
        ),
        "legend_opts": opts.LegendOpts(
            pos_bottom="2%",
            textstyle_opts=opts.TextStyleOpts(font_size=12),
            orient="horizontal"
        ),
        "tooltip_opts": opts.TooltipOpts(
            trigger="axis",
            axis_pointer_type="shadow",
            textstyle_opts=opts.TextStyleOpts(font_size=11)
        ),
        "toolbox_opts": opts.ToolboxOpts(
            is_show=True,
            feature={
                "saveAsImage": {"title": "保存为图片", "pixel_ratio": 2},
                "restore": {"title": "重置"},
                "dataView": {"title": "数据视图", "readOnly": False}
            }
        )
    }
    if zoom:
        cfg["datazoom_opts"] = [
            opts.DataZoomOpts(range_start=0, range_end=100, orient="horizontal"),
            opts.DataZoomOpts(range_start=0, range_end=100, orient="vertical")
        ]
    return cfg


# Excel导出
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


# 金额区间划分
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


# IQR异常值检测
def iqr_outlier(series):
    if series.empty or series.nunique() <= 1:
        return pd.Series([False] * len(series)), 0, 0
    q1, q3 = series.quantile([0.25, 0.75])
    iqr = q3 - q1
    lower_bound = q1 - 1.5 * iqr
    upper_bound = q3 + 1.5 * iqr
    outlier_mask = (series < lower_bound) | (series > upper_bound)
    return outlier_mask, lower_bound, upper_bound


# 适配云端的数据加载逻辑（支持上传文件 + 示例数据）
@st.cache_data(show_spinner="正在加载并清洗数据...", ttl=3600)
def load_data(uploaded_file=None):
    logs = []
    try:
        # 优先使用上传的文件，无则使用示例数据
        if uploaded_file is not None:
            raw = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            # 生成示例数据（避免本地路径依赖）
            st.info("未上传数据文件，使用示例数据演示功能")
            dates = pd.date_range(start="2023-01-01", end="2023-12-31", freq="H")[:1000]
            provinces = ["北京市", "上海市", "广东省", "江苏省", "浙江省", "山东省", "四川省", "湖北省", "河南省",
                         "河北省"]
            raw = pd.DataFrame({
                "Date": dates,
                "district": np.random.choice(provinces, size=len(dates)),
                "buy_mount": np.random.randint(1, 10, size=len(dates)),
                "Total": np.random.uniform(10, 1000, size=len(dates)),
                "user_id": [f"ORD{str(i).zfill(6)}" for i in range(len(dates))]
            })

        logs.append(f"【原始数据】总行数：{raw.shape[0]}，总列数：{raw.shape[1]}")

        df = raw.drop_duplicates()
        dup_count = raw.shape[0] - df.shape[0]
        logs.append(f"【重复值】删除{dup_count}条，剩余{df.shape[0]}行")

        num_cols = df.select_dtypes(include=[np.number]).columns
        obj_cols = df.select_dtypes(include=["object"]).columns
        df[num_cols] = df[num_cols].fillna(0)
        df[obj_cols] = df[obj_cols].fillna("未知")
        logs.append(f"【缺失值】数值列填充0，文本列填充'未知'")

        # 只去除空格，完全保留原始省份名称，不做任何文字追加
        df["district"] = df["district"].str.strip()
        df["省份标准化"] = df["district"]
        logs.append(f"【省份标准化】直接沿用原始行政区名称，无文字追加")

        df["日期"] = pd.to_datetime(df["Date"], errors="coerce")
        date_err = df["日期"].isna().sum()
        df = df.dropna(subset=["日期"])
        logs.append(f"【日期转换】删除无效日期{date_err}条，剩余{df.shape[0]}行")

        df.rename(
            columns={"buy_mount": "购买数量", "Total": "买家实际支付金额", "user_id": "订单编号"},
            inplace=True
        )

        if "买家实际支付金额" in df.columns and not df["买家实际支付金额"].empty:
            mask, low, high = iqr_outlier(df["买家实际支付金额"])
            df = df[~mask]
            logs.append(f"【异常值】阈值[{low:.2f},{high:.2f}]，剔除{mask.sum()}条异常订单")

        df["小时"] = df["日期"].dt.hour
        df["星期名称"] = df["日期"].dt.weekday.map(
            {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'})
        df["金额区间"] = df["买家实际支付金额"].apply(amount_range)
        df["退款金额"] = 0
        logs.append(f"【衍生字段】新增小时、星期名称、金额区间、退款金额字段")

        logs.append(f"【清洗完成】最终有效数据：{df.shape[0]}行")
        return raw, df, logs

    except Exception as e:
        logs.append(f"【错误】数据加载失败：{str(e)}")
        raise


# ======== 云端适配核心修改：文件上传功能 ========
st.sidebar.header("📤 数据文件上传")
uploaded_file = st.sidebar.file_uploader(
    "上传Excel数据文件（.xlsx）",
    type=["xlsx"],
    help="上传包含母婴电商数据的Excel文件，字段需包含：Date、district、buy_mount、Total、user_id"
)

# 加载数据（适配云端，无本地路径）
try:
    RAW, CLEAN, LOGS = load_data(uploaded_file)
    state.raw_df, state.clean_df, state.preprocess_log = RAW, CLEAN, LOGS
except Exception as e:
    st.error(f"❌ 数据加载失败：{str(e)}")
    st.info("✅ 可直接使用示例数据继续演示功能")
    # 强制使用示例数据
    RAW, CLEAN, LOGS = load_data(None)
    state.raw_df, state.clean_df, state.preprocess_log = RAW, CLEAN, LOGS

# 初始化筛选值
df = CLEAN
if state.start_date is None:
    state.start_date = df["日期"].min().date()
if state.end_date is None:
    state.end_date = df["日期"].max().date()

all_prov_list = sorted(df["省份标准化"].unique())
if len(state.sel_prov) == 1 and state.sel_prov[0] == "所有省份":
    state.sel_prov = all_prov_list

if state.min_p == 0.0 and state.max_p == 99999.0:
    state.min_p = float(df["买家实际支付金额"].min())
    state.max_p = float(df["买家实际支付金额"].max())


# 筛选数据
def get_filtered_df():
    if state.filter_df_cache is not None:
        return state.filter_df_cache

    filter_conditions = [
        df["金额区间"].isin(state.sel_price_range),
        df["日期"].dt.date >= state.start_date,
        df["日期"].dt.date <= state.end_date,
        df["省份标准化"].isin(state.sel_prov),
        df["买家实际支付金额"] >= state.min_p,
        df["买家实际支付金额"] <= state.max_p
    ]

    filter_df = df[np.all(filter_conditions, axis=0)].copy()
    state.filter_df_cache = filter_df
    return filter_df


filter_df = get_filtered_df()

# 侧边栏
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
        ("⏰ 分时时段分析", "hour_detail")
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
            disabled=True,
            help="暂无筛选数据可导出"
        )

    st.divider()

    with st.expander("🔍 全局筛选", expanded=True):
        st.multiselect(
            "金额区间",
            options=["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"],
            default=state.sel_price_range,
            key="price_key",
            on_change=on_filter_change,
            help="选择需要分析的订单金额区间"
        )

        col1, col2 = st.columns(2)
        with col1:
            st.date_input(
                "起始日期",
                value=state.start_date,
                min_value=df["日期"].min().date(),
                max_value=df["日期"].max().date(),
                key="start_key",
                on_change=on_filter_change,
                help="筛选起始日期（包含）"
            )
        with col2:
            st.date_input(
                "结束日期",
                value=state.end_date,
                min_value=df["日期"].min().date(),
                max_value=df["日期"].max().date(),
                key="end_key",
                on_change=on_filter_change,
                help="筛选结束日期（包含）"
            )

        prov_options = ["所有省份"] + sorted(df["省份标准化"].unique())
        st.multiselect(
            "省份",
            options=prov_options,
            default=["所有省份"] if set(state.sel_prov) == set(all_prov_list) else state.sel_prov,
            key="prov_key",
            on_change=on_filter_change,
            help="选择需要分析的省份，勾选【所有省份】则自动选中全部省份"
        )
        selected_prov = st.session_state.get("prov_key", [])
        if "所有省份" in selected_prov:
            state.sel_prov = all_prov_list
        else:
            state.sel_prov = selected_prov

        st.slider(
            "支付金额范围",
            min_value=float(df["买家实际支付金额"].min()),
            max_value=float(df["买家实际支付金额"].max()),
            value=(state.min_p, state.max_p),
            key="slider_key",
            on_change=on_filter_change,
            help="自定义支付金额的筛选范围"
        )

# 全局指标
start_date = state.start_date
end_date = state.end_date

total_sales = filter_df["买家实际支付金额"].sum() if not filter_df.empty else 0
total_ord_cnt = len(filter_df)
unique_ord = filter_df["订单编号"].nunique() if not filter_df.empty else 0
avg_price = round(total_sales / unique_ord, 3) if unique_ord > 0 else 0
prov_count = filter_df["省份标准化"].nunique() if not filter_df.empty else 0
page = state.page

# 首页总览
if page == "home":
    st.header("▦ 电商销售数据可视化分析平台 | 综合总览")
    kpi_row = st.columns(4)
    kpi_row[0].metric("筛选总销售额", f"¥{total_sales:,.2f}")
    kpi_row[1].metric("筛选订单总数", f"{total_ord_cnt:,}")
    kpi_row[2].metric("平均客单价", f"¥{avg_price:.3f}")
    kpi_row[3].metric("覆盖省份数量", f"{prov_count}")
    st.divider()

    daily_stats = filter_df.groupby("日期").agg(
        订单量=("订单编号", "count"),
        销售额=("买家实际支付金额", "sum"),
        退款金额=("退款金额", "sum")
    ).reset_index()
    daily_stats_full = pd.merge(
        pd.DataFrame({"日期": pd.date_range(start_date, end_date)}),
        daily_stats,
        how="left"
    ).fillna(0)

    hour_stats = filter_df.groupby("小时").agg(
        订单量=("买家实际支付金额", "count"),
        平均订单金额=("买家实际支付金额", lambda x: round(x.mean(), 3))
    ).reset_index()
    hour_stats_full = pd.merge(
        pd.DataFrame({"小时": range(24)}),
        hour_stats,
        how="left"
    ).fillna(0)

    prov_stats = filter_df.groupby("省份标准化").agg(
        订单量=("订单编号", "count"),
        销售额=("买家实际支付金额", "sum")
    ).reset_index()
    prov_stats = prov_stats[prov_stats["订单量"] > 0].copy()
    top15_sales = prov_stats.sort_values(by="销售额", ascending=False).head(15).rename(columns={"省份标准化": "省份"})

    week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(name="订单量")
    week_stats["排序"] = week_stats["星期名称"].map(
        {"周一": 0, "周二": 1, "周三": 2, "周四": 3, "周五": 4, "周六": 5, "周日": 6})
    week_stats = week_stats.sort_values("排序")

    price_order = ["0-50元", "50-100元", "100-200元", "200-500元", "500元以上"]
    price_stats = filter_df["金额区间"].value_counts().reset_index()
    price_stats.columns = ["金额区间", "订单数"]
    if unique_ord > 0:
        price_stats["占比"] = (price_stats["订单数"] / unique_ord * 100).round(2)
    price_stats["sort_idx"] = price_stats["金额区间"].map(lambda x: price_order.index(x) if x in price_order else 99)
    price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

    r1 = st.columns(3)
    with r1[0]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("每日订单量", daily_stats_full["订单量"].tolist(), is_smooth=True)
            l.set_global_opts(**chart_config("每日订单趋势", min_y=daily_stats_full["订单量"].min()))
        html(l.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=l.render_embed(), file_name="每日订单趋势.html",
                           mime="text/html", use_container_width=True)

    with r1[1]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("每日销售额", daily_stats_full["销售额"].tolist())
            l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
            l.set_global_opts(**chart_config("日销售额面积趋势图", min_y=daily_stats_full["销售额"].min()))
        html(l.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=l.render_embed(), file_name="日销售额趋势图.html",
                           mime="text/html", use_container_width=True)

    with r1[2]:
        b = Bar(chart_init(460))
        if len(hour_stats_full):
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("24小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%",
                        label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
            b.set_global_opts(**chart_config("分时订单柱状图", min_y=0, zoom=True))
        html(b.render_embed(), height=460)
        st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="分时订单柱状图.html",
                           mime="text/html", use_container_width=True)

    st.divider()
    r2 = st.columns(3)
    with r2[0]:
        m = Map(chart_init(420))
        if not prov_stats.empty:
            map_data = list(zip(prov_stats["省份标准化"], prov_stats["订单量"].astype(int)))
            m.add("订单量分布", map_data, maptype="china", is_map_symbol_show=False)
            cfg = chart_config("", zoom=False)
            cfg["title_opts"] = opts.TitleOpts(title="全国省份订单地图", pos_left="center")
            m.set_global_opts(
                visualmap_opts=opts.VisualMapOpts(max_=int(prov_stats["订单量"].max())),
                **cfg
            )
        html(m.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=m.render_embed(), file_name="省份订单地图.html",
                           mime="text/html", use_container_width=True)

    with r2[1]:
        b = Bar(chart_init(520))
        if len(top15_sales):
            b.add_xaxis(top15_sales["省份"].tolist())
            b.add_yaxis("销售额", top15_sales["销售额"].tolist(), bar_width="70%",
                        label_opts=opts.LabelOpts(is_show=True, position="right"))
            b.reversal_axis()
            b.set_global_opts(**chart_config("TOP15省份销售额横向柱状图", min_x=0, zoom=False))
        html(b.render_embed(), height=520)
        st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="TOP15省份销售额.html",
                           mime="text/html", use_container_width=True)

    with r2[2]:
        b = Bar(chart_init(420))
        if len(week_stats):
            b.add_xaxis(week_stats["星期名称"].tolist())
            b.add_yaxis("订单量", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
            b.set_global_opts(**chart_config("星期订单分布", min_y=week_stats["订单量"].min(), zoom=False))
        html(b.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="星期订单分布.html",
                           mime="text/html", use_container_width=True)

    st.divider()
    r3 = st.columns(3)
    with r3[0]:
        p = Pie(chart_init(420))
        if len(price_stats) > 0:
            p.add("", [list(z) for z in zip(price_stats["金额区间"], price_stats["占比"])], radius=["30%", "70%"])
            p.set_global_opts(**chart_config("客单价区间占比饼图", zoom=False))
        html(p.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=p.render_embed(), file_name="客单价饼图.html", mime="text/html",
                           use_container_width=True)

    with r3[1]:
        box = Boxplot(chart_init(440))
        if len(filter_df) > 0:
            box.add_xaxis(["全量订单客单价"])
            box.add_yaxis("客单价分布", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
            box.set_series_opts(markpoint_opts=opts.MarkPointOpts(
                data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
            box.set_global_opts(**chart_config("客单价分布箱线图", min_y=0, zoom=False))
        html(box.render_embed(), height=440)
        st.download_button("📥 下载为HTML（可编辑）", data=box.render_embed(), file_name="客单价箱线图.html",
                           mime="text/html", use_container_width=True)

    with r3[2]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("退款金额", daily_stats_full["退款金额"].tolist(), is_smooth=True)
            l.set_global_opts(**chart_config("每日退款趋势", min_y=daily_stats_full["退款金额"].min()))
        html(l.render_embed(), height=420)
        st.download_button("📥 下载为HTML（可编辑）", data=l.render_embed(), file_name="每日退款趋势.html",
                           mime="text/html", use_container_width=True)

# 数据预处理日志页
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

# 基础统计分析页
elif page == "stat_analysis":
    st.header("📈 基础探索性统计分析")
    if len(filter_df) > 0:
        desc_df = filter_df[["购买数量", "买家实际支付金额", "小时"]].describe()
        desc_df = desc_df.round(3)
        st.dataframe(desc_df, use_container_width=True, hide_index=True)
        st.divider()

        corr_cols = ["购买数量", "买家实际支付金额", "小时"]
        corr_matrix = filter_df[corr_cols].corr(method="pearson")

        fig, ax = plt.subplots(figsize=(7, 5), dpi=300)
        sns.heatmap(
            corr_matrix,
            annot=True,
            cmap="Blues",
            vmin=-0.1,
            vmax=1,
            ax=ax,
            fmt=".3f",
            linewidths=0.5
        )
        ax.set_title("皮尔逊相关系数热力图", fontsize=16, pad=20)
        ax.set_xticklabels(["购买数量", "买家实际支付金额", "小时"], rotation=0)
        ax.set_yticklabels(["购买数量", "买家实际支付金额", "小时"], rotation=0)
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight", dpi=300)
        buf.seek(0)
        st.image(buf, use_container_width=True)

        st.download_button(
            label="📥 下载热力图图片",
            data=buf,
            file_name=f"相关性热力图_{date.today()}.png",
            mime="image/png",
            use_container_width=True
        )
        plt.close()
    else:
        st.warning("⚠️ 暂无筛选数据，无法进行统计分析")

# 销售额详情页
elif page == "sale_detail":
    st.header("💰 销售额深度详情")
    if len(filter_df) > 0:
        daily_stats = filter_df.groupby("日期").agg(
            订单量=("订单编号", "count"),
            销售额=("买家实际支付金额", "sum"),
            退款金额=("退款金额", "sum")
        ).reset_index()
        daily_stats_full = pd.merge(
            pd.DataFrame({"日期": pd.date_range(start_date, end_date)}),
            daily_stats,
            how="left"
        ).fillna(0)

        prov_stats = filter_df.groupby("省份标准化").agg(
            订单量=("订单编号", "count"),
            销售额=("买家实际支付金额", "sum")
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
            st.download_button("📥 下载为HTML（可编辑）", data=l.render_embed(), file_name="日销售额趋势图.html",
                               mime="text/html", use_container_width=True)

        with col2:
            b = Bar(chart_init(520))
            b.add_xaxis(top15_sales["省份"].tolist())
            b.add_yaxis("省份销售额", top15_sales["销售额"].tolist(), bar_width="70%",
                        label_opts=opts.LabelOpts(is_show=True, position="right"))
            b.reversal_axis()
            b.set_global_opts(**chart_config("TOP15省份销售额", min_x=0, zoom=False))
            html(b.render_embed(), height=520)
            st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="TOP15省份销售额.html",
                               mime="text/html", use_container_width=True)

        top15_sales_ranked = top15_sales.reset_index(drop=True)
        top15_sales_ranked.insert(0, "排名", range(1, len(top15_sales_ranked) + 1))
        st.dataframe(top15_sales_ranked, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ 暂无筛选数据，无法展示销售额详情")

# 订单量详情页
elif page == "order_detail":
    st.header("📦 订单量深度详情")
    if len(filter_df) > 0:
        daily_stats = filter_df.groupby("日期").agg(
            订单量=("订单编号", "count"),
            销售额=("买家实际支付金额", "sum"),
            退款金额=("退款金额", "sum")
        ).reset_index()
        daily_stats_full = pd.merge(
            pd.DataFrame({"日期": pd.date_range(start_date, end_date)}),
            daily_stats,
            how="left"
        ).fillna(0)

        week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(name="订单量")
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
            st.download_button("📥 下载为HTML（可编辑）", data=l.render_embed(), file_name="每日订单趋势.html",
                               mime="text/html", use_container_width=True)

        with col2:
            b = Bar(chart_init(480))
            b.add_xaxis(week_stats["星期名称"].tolist())
            b.add_yaxis("周订单量", week_stats["订单量"].tolist(), label_opts=opts.LabelOpts(is_show=True))
            b.set_global_opts(**chart_config("星期订单分布", min_y=week_stats["订单量"].min(), zoom=False))
            html(b.render_embed(), height=480)
            st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="星期订单分布.html",
                               mime="text/html", use_container_width=True)
    else:
        st.warning("⚠️ 暂无筛选数据，无法展示订单量详情")

# 客单价区间分析页
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
        st.download_button("📥 下载为HTML（可编辑）", data=p.render_embed(), file_name="价格区间占比饼图.html",
                           mime="text/html", use_container_width=True)

        st.dataframe(price_stats, use_container_width=True, hide_index=True)
        st.divider()

        box = Boxplot(chart_init(500))
        box.add_xaxis(["客单价分布"])
        box.add_yaxis("客单价(元)", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
        box.set_series_opts(
            markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
        box.set_global_opts(**chart_config("客单价分布箱线图", min_y=0, zoom=False))
        html(box.render_embed(), height=500)
        st.download_button("📥 下载为HTML（可编辑）", data=box.render_embed(), file_name="客单价箱线图.html",
                           mime="text/html", use_container_width=True)
    else:
        st.warning("⚠️ 暂无筛选数据，无法展示客单价分析")

# 省份地理分析页
elif page == "province_detail":
    st.header("🗺️ 全国省份地理销售详情")
    if len(filter_df) > 0:
        filter_df["年月"] = filter_df["日期"].dt.to_period("M")
        month_group = filter_df.groupby(["年月", "省份标准化"])["订单编号"].count().reset_index()
        month_group = month_group[month_group["订单编号"] > 0]

        tl = Timeline(chart_init(550))
        tl.add_schema(
            play_interval=1000,
            is_auto_play=False,
            is_loop_play=False,
            pos_bottom="5%",
            label_opts=opts.LabelOpts(font_size=12)
        )

        for ym in sorted(month_group["年月"].unique()):
            sub_df = month_group[month_group["年月"] == ym]
            map_data = list(zip(sub_df["省份标准化"], sub_df["订单编号"]))
            m = Map(chart_init(550))
            m.add("订单量", map_data, maptype="china")
            cfg = chart_config("", zoom=False)
            cfg["title_opts"] = opts.TitleOpts(title=f"{ym} 各省份订单分布", pos_left="center")
            m.set_global_opts(
                visualmap_opts=opts.VisualMapOpts(max_=sub_df["订单编号"].max()),
                **cfg
            )
            tl.add(m, str(ym))

        html(tl.render_embed(), height=580)
        st.download_button("📥 下载为HTML（可编辑）", data=tl.render_embed(), file_name="时间轴轮播地图.html",
                           mime="text/html", use_container_width=True)

        prov_stats = filter_df.groupby("省份标准化").agg(
            订单量=("订单编号", "count"),
            销售额=("买家实际支付金额", "sum")
        ).reset_index()
        prov_stats = prov_stats[prov_stats["订单量"] > 0]
        prov_stats_sorted = prov_stats.sort_values("销售额", ascending=False).reset_index(drop=True)
        prov_stats_sorted.insert(0, "排名", range(1, len(prov_stats_sorted) + 1))
        st.dataframe(prov_stats_sorted, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ 暂无筛选数据，无法展示省份地理分析")

# 分时时段分析页
elif page == "hour_detail":
    st.header("⏰ 24小时分时深度分析")
    if len(filter_df) > 0:
        hour_stats = filter_df.groupby("小时").agg(
            订单量=("买家实际支付金额", "count"),
            平均订单金额=("买家实际支付金额", lambda x: round(x.mean(), 3))
        ).reset_index()
        hour_stats_full = pd.merge(
            pd.DataFrame({"小时": range(24)}),
            hour_stats,
            how="left"
        ).fillna(0)

        col1, col2 = st.columns(2)
        with col1:
            b = Bar(chart_init(480))
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("每小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%",
                        label_opts=opts.LabelOpts(is_show=True, font_size=9, rotate=30))
            b.set_global_opts(**chart_config("分时订单量", min_y=0, zoom=True))
            html(b.render_embed(), height=480)
            st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="分时订单量.html",
                               mime="text/html", use_container_width=True)

        with col2:
            b = Bar(chart_init(480))
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("每小时平均客单价", hour_stats_full["平均订单金额"].tolist(),
                        label_opts=opts.LabelOpts(font_size=10))
            b.set_global_opts(**chart_config("分时平均客单价", min_y=hour_stats_full["平均订单金额"].min(), zoom=True))
            html(b.render_embed(), height=480)
            st.download_button("📥 下载为HTML（可编辑）", data=b.render_embed(), file_name="分时平均客单价.html",
                               mime="text/html", use_container_width=True)
    else:
        st.warning("⚠️ 暂无筛选数据，无法展示分时时段分析")
