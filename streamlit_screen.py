import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from datetime import date
from pyecharts import options as opts
from pyecharts.charts import Line, Bar, Pie, Map, Boxplot, Timeline
from pyecharts.globals import ThemeType
from streamlit.components.v1 import html
import matplotlib.pyplot as plt
import seaborn as sns

# 屏蔽警告
warnings.filterwarnings("ignore")

# 解决matplotlib中文乱码（云端Linux兼容）
plt.rcParams["font.sans-serif"] = ["WenQuanYi Zen Hei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

# 页面基础配置
st.set_page_config(
    page_title="母婴电商销售数据可视化分析平台",
    page_icon="📊",
    layout="wide"
)

# 仅保留最小样式，彻底规避DOM节点冲突
st.markdown("""
<style>
iframe {width:100% !important;}
</style>
""", unsafe_allow_html=True)

# 会话状态初始化
state = st.session_state
state.setdefault("page", "home")
state.setdefault("raw_df", None)
state.setdefault("clean_df", None)
state.setdefault("preprocess_log", [])
# 默认选中全部金额区间
state.setdefault("sel_price_range", ["0-50元","50-100元","100-200元","200-500元","500元以上"])
state.setdefault("start_date", None)
state.setdefault("end_date", None)
# 省份默认空 = 全部省份
state.setdefault("sel_prov", [])
state.setdefault("min_p", None)
state.setdefault("max_p", None)


# 方案2：增加key存在判断，修复KeyError
def on_filter_change():
    if "price_key" not in state:
        return
    if "start_key" not in state:
        return
    if "end_key" not in state:
        return
    if "prov_key" not in state:
        return
    if "slider_key" not in state:
        return
    state.sel_price_range = state["price_key"]
    state.start_date = state["start_key"]
    state.end_date = state["end_key"]
    state.sel_prov = state["prov_key"]
    state.min_p, state.max_p = state["slider_key"]


def chart_init(height=480):
    return opts.InitOpts(
        width="100%",
        height=f"{height}px",
        bg_color="#ffffff",
        theme=ThemeType.MACARONS
    )


def chart_config(title_name, min_y=0, min_x=None, zoom=True):
    cfg = {
        "title_opts": opts.TitleOpts(title=title_name, pos_left="center", title_textstyle_opts=opts.TextStyleOpts(font_size=16)),
        "xaxis_opts": opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=12, rotate=15), splitline_opts=opts.SplitLineOpts(is_show=False)),
        "yaxis_opts": opts.AxisOpts(min_=min_y, axislabel_opts=opts.LabelOpts(font_size=12), splitline_opts=opts.SplitLineOpts(is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=0.3))),
        "legend_opts": opts.LegendOpts(pos_bottom="2%", textstyle_opts=opts.TextStyleOpts(font_size=12)),
        "tooltip_opts": opts.TooltipOpts(trigger="axis", axis_pointer_type="shadow"),
        "toolbox_opts": opts.ToolboxOpts(is_show=True, feature={"saveAsImage": {"title": "保存为图片", "pixel_ratio": 2}})
    }
    if min_x is not None:
        cfg["xaxis_opts"].min_ = min_x
    if zoom:
        cfg["datazoom_opts"] = opts.DataZoomOpts(range_start=0, range_end=100)
    return cfg


def export_excel(sheet_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheet_dict.items():
            data.to_excel(writer, sheet_name=name, index=False)
    return output.getvalue()


def amount_range(val):
    bins = [50, 100, 200, 500]
    labels = ("0-50元","50-100元","100-200元","200-500元","500元以上")
    idx = np.searchsorted(bins, val)
    return labels[idx]


def iqr_outlier(series):
    q1, q3 = series.quantile([0.25, 0.75])
    iqr = q3 - q1
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    return (series < lower) | (series > upper), lower, upper


@st.cache_data(show_spinner="正在加载并清洗数据...")
def load_data(file_bytes):
    logs = []
    raw = pd.read_excel(io.BytesIO(file_bytes))
    logs.append(f"【原始数据】总行数：{raw.shape[0]}，总列数：{raw.shape[1]}")
    df = raw.drop_duplicates()
    logs.append(f"【重复值】删除{raw.shape[0]-df.shape[0]}条，剩余{df.shape[0]}行")
    num_cols = df.select_dtypes(include=[np.number]).columns
    df[num_cols] = df[num_cols].fillna(0)
    obj_cols = df.select_dtypes(include=["object"]).columns
    df[obj_cols] = df[obj_cols].fillna("未知")

    df["district"] = df["district"].str.replace("省$", "", regex=True)
    province_map = {
        '北京':'北京市','天津':'天津市','上海':'上海市','重庆':'重庆市',
        '内蒙古':'内蒙古自治区','广西壮族':'广西壮族自治区','西藏':'西藏自治区',
        '宁夏':'宁夏回族自治区','新疆维吾尔自治区':'新疆维吾尔自治区',
        '香港':'香港特别行政区','澳门':'澳门特别行政区','台湾':'台湾省'
    }
    df["省份标准化"] = df["district"].replace(province_map)
    df["日期"] = pd.to_datetime(df["Date"], errors="coerce")
    date_err = df["日期"].isna().sum()
    df = df.dropna(subset=["日期"])
    logs.append(f"【日期转换】删除无效日期{date_err}条")
    df.rename(columns={"buy_mount":"购买数量","Total":"买家实际支付金额","user_id":"订单编号"}, inplace=True)
    mask, low, high = iqr_outlier(df["买家实际支付金额"])
    df = df[~mask]
    logs.append(f"【异常值】阈值[{low:.2f},{high:.2f}]，剔除{mask.sum()}条异常订单")
    df["小时"] = df["日期"].dt.hour
    df["星期名称"] = df["日期"].dt.weekday.map({0:'周一',1:'周二',2:'周三',3:'周四',4:'周五',5:'周六',6:'周日'})
    df["金额区间"] = df["买家实际支付金额"].apply(amount_range)
    df["退款金额"] = 0
    logs.append(f"【清洗完成】最终有效数据：{df.shape[0]}行")
    return raw, df, logs


# 云端文件上传入口
st.title("📊 母婴电商销售数据可视化分析平台")
uploaded_file = st.file_uploader("请上传Excel数据文件（clean_baby_data.xlsx）", type=["xlsx"])

df = None
if uploaded_file is not None:
    RAW, CLEAN, LOGS = load_data(uploaded_file.read())
    state.raw_df, state.clean_df, state.preprocess_log = RAW, CLEAN, LOGS
    df = CLEAN

    # 每次上传新文件，自动把金额滑块重置为数据集真实最小、最大值
    real_min = float(df["买家实际支付金额"].min())
    real_max = float(df["买家实际支付金额"].max())
    state.min_p = real_min
    state.max_p = real_max

    # 日期自动重置为数据集完整区间
    state.start_date = df["日期"].min().date()
    state.end_date = df["日期"].max().date()

else:
    st.info("请先上传数据文件，否则无法继续分析！")
    st.stop()


# 侧边栏
with st.sidebar:
    st.header("📊 功能导航")
    nav = [
        ("🏠 首页总览", "home"),
        ("📋 数据预处理日志", "preprocess"),
        ("📈 基础统计分析", "stat_analysis"),
        ("💰 销售额详情", "sale_detail"),
        ("📦 订单量详情", "order_detail"),
        ("💵 客单价区间分析", "price_detail"),
        ("🗺️ 省份地理分析", "province_detail"),
        ("⏰ 分时时段分析", "hour_detail")
    ]
    for txt, key in nav:
        if st.button(txt, use_container_width=True):
            state.page = key
            st.rerun()

    st.divider()
    st.session_state["export_bytes"] = export_excel({"筛选后数据": df})
    st.download_button(
        label="📥 导出当前筛选数据.xlsx",
        data=st.session_state["export_bytes"],
        file_name="筛选销售数据.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()
    with st.expander("🔍 全局筛选", expanded=True):
        price_order = ["0-50元","50-100元","100-200元","200-500元","500元以上"]
        st.multiselect("金额区间", price_order, default=state.sel_price_range, key="price_key", on_change=on_filter_change)
        c1, c2 = st.columns(2)
        c1.date_input("起始日期", value=state.start_date, min_value=df["日期"].min().date(), max_value=df["日期"].max().date(), key="start_key", on_change=on_filter_change)
        c2.date_input("结束日期", value=state.end_date, min_value=df["日期"].min().date(), max_value=df["日期"].max().date(), key="end_key", on_change=on_filter_change)
        st.multiselect("省份", sorted(df["省份标准化"].unique()), default=state.sel_prov, key="prov_key", on_change=on_filter_change)

        # 滑块自动绑定当前数据集真实极值
        slider_min = float(df["买家实际支付金额"].min())
        slider_max = float(df["买家实际支付金额"].max())
        st.slider(
            "支付金额范围",
            min_value=slider_min,
            max_value=slider_max,
            value=(state.min_p, state.max_p),
            key="slider_key",
            on_change=on_filter_change
        )


# 筛选逻辑（省份为空=查询全部省份，和其他筛选条件同步生效）
if len(state.sel_prov) == 0:
    filter_df = df[
        (df["金额区间"].isin(state.sel_price_range)) &
        (df["日期"].dt.date >= state.start_date) &
        (df["日期"].dt.date <= state.end_date) &
        (df["买家实际支付金额"] >= state.min_p) &
        (df["买家实际支付金额"] <= state.max_p)
    ].copy()
else:
    filter_df = df[
        (df["金额区间"].isin(state.sel_price_range)) &
        (df["日期"].dt.date >= state.start_date) &
        (df["日期"].dt.date <= state.end_date) &
        (df["省份标准化"].isin(state.sel_prov)) &
        (df["买家实际支付金额"] >= state.min_p) &
        (df["买家实际支付金额"] <= state.max_p)
    ].copy()

st.session_state["export_bytes"] = export_excel({"筛选后数据": filter_df})

# 指标
total_sales = filter_df["买家实际支付金额"].sum()
total_ord_cnt = len(filter_df)
unique_ord = filter_df["订单编号"].nunique()
avg_price = round(total_sales / unique_ord, 3) if unique_ord else 0
prov_count = filter_df["省份标准化"].nunique()
page = state.page


if page == "home":
    st.header("▦ 电商销售数据可视化分析平台 | 综合总览")
    kpi_row = st.columns(4)
    kpi_row[0].metric("筛选总销售额", f"¥{total_sales:,.2f}")
    kpi_row[1].metric("筛选订单总数", f"{total_ord_cnt:,}")
    kpi_row[2].metric("平均客单价", f"¥{avg_price:.3f}")
    kpi_row[3].metric("覆盖省份数量", f"{prov_count}")
    st.divider()

    daily_stats = filter_df.groupby("日期").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum"),退款金额=("退款金额","sum")).reset_index()
    daily_stats_full = pd.merge(pd.DataFrame({"日期":pd.date_range(state.start_date, state.end_date)}), daily_stats, how="left").fillna(0)
    hour_stats = filter_df.groupby("小时").agg(订单量=("买家实际支付金额","count"),平均订单金额=("买家实际支付金额",lambda x:round(x.mean(),3))).reset_index()
    hour_stats_full = pd.merge(pd.DataFrame({"小时":range(24)}), hour_stats, how="left").fillna(0)
    prov_stats = filter_df.groupby("省份标准化").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum")).reset_index()
    top15_sales = prov_stats.sort_values(by="销售额",ascending=False).head(15).rename(columns={"省份标准化":"省份"})
    week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(name="订单量")
    week_stats["排序"] = week_stats["星期名称"].map({"周一":0,"周二":1,"周三":2,"周四":3,"周五":4,"周六":5,"周日":6})
    week_stats = week_stats.sort_values("排序")
    price_order = ["0-50元","50-100元","100-200元","200-500元","500元以上"]
    price_stats = filter_df["金额区间"].value_counts().reset_index()
    price_stats.columns = ["金额区间","订单数"]
    if unique_ord>0:
        price_stats["占比"] = (price_stats["订单数"] / unique_ord * 100).round(2)
    price_stats["sort_idx"] = price_stats["金额区间"].map(lambda x:price_order.index(x))
    price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)

    r1 = st.columns(3)
    with r1[0]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("每日订单量", daily_stats_full["订单量"].tolist(), is_smooth=True)
            l.set_global_opts(**chart_config("每日订单趋势"))
        html(l.render_embed(), height=420)
    with r1[1]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("每日销售额", daily_stats_full["销售额"].tolist())
            l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
            l.set_global_opts(**chart_config("日销售额面积趋势图"))
        html(l.render_embed(), height=420)
    with r1[2]:
        b = Bar(chart_init(460))
        if len(hour_stats_full):
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("24小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%")
            b.set_global_opts(**chart_config("分时订单柱状图"))
        html(b.render_embed(), height=460)

    st.divider()
    r2 = st.columns(3)
    with r2[0]:
        m = Map(chart_init(420))
        if len(filter_df) > 0:
            map_data = list(zip(prov_stats["省份标准化"], prov_stats["订单量"].astype(int)))
            m.add("订单量分布", map_data, maptype="china", is_map_symbol_show=False)
            m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=int(prov_stats["订单量"].max())),**chart_config("全国省份订单地图"))
        html(m.render_embed(), height=420)
    with r2[1]:
        b = Bar(chart_init(520))
        if len(top15_sales):
            b.add_xaxis(top15_sales["省份"].tolist())
            b.add_yaxis("销售额", top15_sales["销售额"].tolist(), bar_width="70%")
            b.reversal_axis()
            b.set_global_opts(**chart_config("TOP15省份销售额横向柱状图", min_x=0, zoom=False))
        html(b.render_embed(), height=520)
    with r2[2]:
        b = Bar(chart_init(420))
        if len(week_stats):
            b.add_xaxis(week_stats["星期名称"].tolist())
            b.add_yaxis("订单量", week_stats["订单量"].tolist())
            b.set_global_opts(**chart_config("星期订单分布", zoom=False))
        html(b.render_embed(), height=420)

    st.divider()
    r3 = st.columns(3)
    with r3[0]:
        p = Pie(chart_init(420))
        if len(price_stats) > 0:
            p.add("", [list(z) for z in zip(price_stats["金额区间"], price_stats["占比"])], radius=["30%","70%"])
            p.set_global_opts(**chart_config("客单价区间占比饼图", zoom=False))
        html(p.render_embed(), height=420)
    with r3[1]:
        box = Boxplot(chart_init(440))
        if len(filter_df) > 0:
            box.add_xaxis(["全量订单客单价"])
            box.add_yaxis("客单价分布", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
            box.set_series_opts(markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
            box.set_global_opts(**chart_config("客单价分布箱线图", zoom=False))
        html(box.render_embed(), height=440)
    with r3[2]:
        l = Line(chart_init(420))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("退款金额", daily_stats_full["退款金额"].tolist(), is_smooth=True)
            l.set_global_opts(**chart_config("每日退款趋势"))
        html(l.render_embed(), height=420)

elif page == "preprocess":
    st.header("📋 数据预处理完整日志")
    for log in state.preprocess_log:
        st.info(log)
    st.divider()
    c1, c2 = st.columns(2)
    c1.subheader("原始数据前10行")
    c1.dataframe(state.raw_df.head(10), use_container_width=True, hide_index=True)
    c2.subheader("清洗后数据前10行")
    c2.dataframe(state.clean_df.head(10), use_container_width=True, hide_index=True)

elif page == "stat_analysis":
    st.header("📈 基础探索性统计分析")
    if len(filter_df) > 0:
        st.dataframe(filter_df[["购买数量","买家实际支付金额","小时"]].describe(), use_container_width=True, hide_index=True)
        st.divider()
        corr_cols = ["购买数量", "买家实际支付金额", "小时"]
        corr_matrix = filter_df[corr_cols].corr(method="pearson")
        fig, ax = plt.subplots(figsize=(7, 5), dpi=300)
        sns.heatmap(corr_matrix, annot=True, cmap="Blues", vmin=-0.1, vmax=1, ax=ax)
        ax.set_title("皮尔逊相关系数热力图", fontsize=16)
        ax.set_xticklabels(corr_cols)
        ax.set_yticklabels(corr_cols)
        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        st.image(buf)
        plt.close()

elif page == "sale_detail":
    st.header("💰 销售额深度详情")
    daily_stats = filter_df.groupby("日期").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum"),退款金额=("退款金额","sum")).reset_index()
    daily_stats_full = pd.merge(pd.DataFrame({"日期":pd.date_range(state.start_date, state.end_date)}), daily_stats, how="left").fillna(0)
    prov_stats = filter_df.groupby("省份标准化").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum")).reset_index()
    top15_sales = prov_stats.sort_values(by="销售额",ascending=False).head(15).rename(columns={"省份标准化":"省份"})
    c1, c2 = st.columns(2)
    with c1:
        l = Line(chart_init(480))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("日销售额", daily_stats_full["销售额"].tolist())
            l.set_series_opts(areastyle_opts=opts.AreaStyleOpts(opacity=0.4))
            l.set_global_opts(**chart_config("日销售额面积趋势图"))
        html(l.render_embed(), height=480)
    with c2:
        b = Bar(chart_init(520))
        if len(top15_sales):
            b.add_xaxis(top15_sales["省份"].tolist())
            b.add_yaxis("省份销售额", top15_sales["销售额"].tolist(), bar_width="70%")
            b.reversal_axis()
            b.set_global_opts(**chart_config("TOP15省份销售额", min_x=0, zoom=False))
        html(b.render_embed(), height=520)
    if len(top15_sales) > 0:
        top15_sales_ranked = top15_sales.reset_index(drop=True)
        top15_sales_ranked.insert(0, "排名", range(1, len(top15_sales_ranked)+1))
        st.dataframe(top15_sales_ranked, use_container_width=True, hide_index=True)

elif page == "order_detail":
    st.header("📦 订单量深度详情")
    daily_stats = filter_df.groupby("日期").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum"),退款金额=("退款金额","sum")).reset_index()
    daily_stats_full = pd.merge(pd.DataFrame({"日期":pd.date_range(state.start_date, state.end_date)}), daily_stats, how="left").fillna(0)
    week_stats = filter_df.groupby("星期名称")["订单编号"].count().reset_index(name="订单量")
    week_stats["排序"] = week_stats["星期名称"].map({"周一":0,"周二":1,"周三":2,"周四":3,"周五":4,"周六":5,"周日":6})
    week_stats = week_stats.sort_values("排序")
    c1, c2 = st.columns(2)
    with c1:
        l = Line(chart_init(480))
        if len(daily_stats_full):
            l.add_xaxis([d.strftime("%m-%d") for d in daily_stats_full["日期"]])
            l.add_yaxis("日订单数", daily_stats_full["订单量"].tolist(), is_smooth=True)
            l.set_global_opts(**chart_config("每日订单趋势"))
        html(l.render_embed(), height=480)
    with c2:
        b = Bar(chart_init(480))
        if len(week_stats):
            b.add_xaxis(week_stats["星期名称"].tolist())
            b.add_yaxis("周订单量", week_stats["订单量"].tolist())
            b.set_global_opts(**chart_config("星期订单分布", zoom=False))
        html(b.render_embed(), height=480)

elif page == "price_detail":
    st.header("💵 客单价&金额区间分析")
    price_order = ["0-50元","50-100元","100-200元","200-500元","500元以上"]
    price_stats = filter_df["金额区间"].value_counts().reset_index()
    price_stats.columns = ["金额区间","订单数"]
    if unique_ord > 0:
        price_stats["占比"] = (price_stats["订单数"] / unique_ord * 100).round(2)
    price_stats["sort_idx"] = price_stats["金额区间"].map(lambda x:price_order.index(x))
    price_stats = price_stats.sort_values("sort_idx").reset_index(drop=True)
    p = Pie(chart_init(500))
    if len(price_stats) > 0:
        p.add("", [list(z) for z in zip(price_stats["金额区间"], price_stats["占比"])], radius=["30%","70%"])
        p.set_global_opts(**chart_config("价格区间占比饼图", zoom=False))
    html(p.render_embed(), height=500)
    if len(price_stats) > 0:
        st.dataframe(price_stats, use_container_width=True, hide_index=True)
    st.divider()
    box = Boxplot(chart_init(500))
    if len(filter_df) > 0:
        box.add_xaxis(["客单价分布"])
        box.add_yaxis("客单价(元)", box.prepare_data([filter_df["买家实际支付金额"].dropna().tolist()]))
        box.set_series_opts(markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="max"), opts.MarkPointItem(type_="min")]))
        box.set_global_opts(**chart_config("客单价分布箱线图", zoom=False))
    html(box.render_embed(), height=500)

elif page == "province_detail":
    st.header("🗺️ 全国省份地理销售详情")
    if len(filter_df) == 0:
        st.warning("当前筛选条件下无数据，请重新选择筛选条件！")
    else:
        filter_df["年月"] = filter_df["日期"].dt.to_period("M")
        month_group = filter_df.groupby(["年月", "省份标准化"])["订单编号"].count().reset_index()
        tl = Timeline(chart_init(550))
        for ym in sorted(month_group["年月"].unique()):
            sub_df = month_group[month_group["年月"] == ym]
            map_data = list(zip(sub_df["省份标准化"], sub_df["订单编号"]))
            m = Map(opts.InitOpts(width="100%", height="500px"))
            m.add("订单量", map_data, maptype="china", is_map_symbol_show=False)
            m.set_global_opts(visualmap_opts=opts.VisualMapOpts(max_=int(sub_df["订单编号"].max())),title_opts=opts.TitleOpts(title=f"{ym} 各省份订单分布"))
            tl.add(m, str(ym))
        html(tl.render_embed(), height=580)
        prov_stats = filter_df.groupby("省份标准化").agg(订单量=("订单编号","count"),销售额=("买家实际支付金额","sum")).reset_index()
        prov_stats_sorted = prov_stats.sort_values("销售额", ascending=False).reset_index(drop=True)
        prov_stats_sorted.insert(0, "排名", range(1, len(prov_stats_sorted) + 1))
        st.dataframe(prov_stats_sorted, use_container_width=True, hide_index=True)

elif page == "hour_detail":
    st.header("⏰ 24小时分时深度分析")
    hour_stats = filter_df.groupby("小时").agg(订单量=("买家实际支付金额","count"),平均订单金额=("买家实际支付金额",lambda x:round(x.mean(),3))).reset_index()
    hour_stats_full = pd.merge(pd.DataFrame({"小时":range(24)}), hour_stats, how="left").fillna(0)
    c1, c2 = st.columns(2)
    with c1:
        b = Bar(chart_init(480))
        if len(hour_stats_full):
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("每小时订单", hour_stats_full["订单量"].tolist(), bar_width="60%")
            b.set_global_opts(**chart_config("分时订单量"))
        html(b.render_embed(), height=480)
    with c2:
        b = Bar(chart_init(480))
        if len(hour_stats_full):
            b.add_xaxis([str(i) for i in range(24)])
            b.add_yaxis("每小时平均客单价", hour_stats_full["平均订单金额"].tolist())
            b.set_global_opts(**chart_config("分时平均客单价"))
        html(b.render_embed(), height=480)
