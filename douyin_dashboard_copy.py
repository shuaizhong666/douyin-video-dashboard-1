"""
抖音视频数据可视化看板（本地路径版）
数据文件：C:\RPAWorkspace\抖音视频数据汇总.xlsx
运行命令：streamlit run 本文件.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import warnings
from pathlib import Path
from datetime import date

warnings.filterwarnings('ignore')

# ======================== 页面配置 ========================
st.set_page_config(
    page_title="抖音视频数据分析看板",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================== 数据文件路径 ========================
DATA_FILE = Path(r"C:\RPAWorkspace\抖音视频数据汇总.xlsx")  # raw string 避免转义

# ======================== 数据加载与清洗 ========================
@st.cache_data
def load_data(file_path):
    """加载Excel，返回原始DataFrame（中文列名）"""
    if not file_path.exists():
        st.error(f"❌ 文件不存在: {file_path}\n请确保文件已放置在指定目录。")
        return pd.DataFrame()
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        st.success(f"✅ 成功加载数据，共 {df.shape[0]} 行，{df.shape[1]} 列")
        return df
    except Exception as e:
        st.error(f"读取Excel失败: {e}")
        return pd.DataFrame()

def preprocess_for_analysis(df):
    """
    为分析功能预处理数据：
    - 转换数值列
    - 转换日期列
    - 保留原始中文列名，仅新增派生列（如 'publish_date' 为日期类型）
    """
    if df.empty:
        return df
    df_work = df.copy()

    # 数值列转换
    num_cols_cn = ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '获赞总数']
    for col in num_cols_cn:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce')

    # 日期列转换
    if '创建时间' in df_work.columns:
        df_work['publish_date'] = pd.to_datetime(df_work['创建时间'], errors='coerce')

    return df_work

def compute_kpis(df):
    """计算关键指标（df已包含数值列）"""
    if df.empty:
        return {}
    total_videos = len(df)
    total_likes = df['点赞数'].sum() if '点赞数' in df else None
    total_comments = df['评论数'].sum() if '评论数' in df else None
    total_shares = df['分享数'].sum() if '分享数' in df else None
    total_collects = df['收藏数'].sum() if '收藏数' in df else None
    avg_likes = df['点赞数'].mean() if '点赞数' in df else None
    if '作者昵称' in df and '粉丝数' in df:
        total_fans = df.groupby('作者昵称')['粉丝数'].max().sum()
    else:
        total_fans = None
    return {
        'total_videos': total_videos,
        'total_likes': total_likes,
        'total_comments': total_comments,
        'total_shares': total_shares,
        'total_collects': total_collects,
        'avg_likes_per_video': avg_likes,
        'total_fans': total_fans
    }

# ======================== 图表绘制函数 ========================
def plot_trend(df, date_col, value_col, title, color='#FF5722'):
    if df.empty or date_col not in df or value_col not in df:
        return None
    temp = df[[date_col, value_col]].dropna()
    if temp.empty:
        return None
    temp['date'] = temp[date_col].dt.date
    daily = temp.groupby('date')[value_col].sum().reset_index()
    fig = px.line(daily, x='date', y=value_col, title=title, markers=True)
    fig.update_traces(line_color=color)
    return fig

def plot_top10(df, value_col, title_col=None, title="Top10"):
    if df.empty or value_col not in df:
        return None
    if title_col and title_col in df:
        plot_df = df[[title_col, value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = title_col
    elif '视频描述' in df:
        plot_df = df[['视频描述', value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = '视频描述'
    elif '视频ID' in df:
        plot_df = df[['视频ID', value_col]].dropna().sort_values(value_col, ascending=False).head(10)
        x_axis = '视频ID'
    else:
        plot_df = df[[value_col]].dropna().reset_index()
        plot_df['index'] = plot_df['index'] + 1
        plot_df.rename(columns={'index': '视频序号'}, inplace=True)
        x_axis = '视频序号'
    fig = px.bar(plot_df, x=x_axis, y=value_col, title=title, text_auto=True)
    fig.update_layout(xaxis_title="视频", yaxis_title=value_col, xaxis_tickangle=-45)
    return fig

def plot_scatter(df, x_col, y_col, size_col=None, color_col=None, title="相关关系"):
    if df.empty or x_col not in df or y_col not in df:
        return None
    cols = [x_col, y_col]
    if size_col and size_col in df.columns:
        cols.append(size_col)
    if color_col and color_col in df.columns:
        cols.append(color_col)
    plot_df = df[cols].dropna()
    if plot_df.empty:
        return None
    fig = px.scatter(plot_df, x=x_col, y=y_col,
                     size=size_col if size_col in plot_df.columns else None,
                     color=color_col if color_col in plot_df.columns else None,
                     title=title, opacity=0.6)
    return fig

def plot_publish_hour_heatmap(df, date_col):
    if df.empty or date_col not in df:
        return None
    temp = df[date_col].dropna()
    if temp.empty:
        return None
    hours = temp.dt.hour
    weekdays = temp.dt.dayofweek
    heat_data = pd.crosstab(weekdays, hours)
    heat_data.index = ['周一','周二','周三','周四','周五','周六','周日']
    fig = px.imshow(heat_data, text_auto=True, aspect="auto",
                    title="发布时段热力图 (星期-小时)",
                    labels=dict(x="小时", y="星期", color="发布数量"))
    return fig

def get_author_aggregation(df):
    """基于筛选后的df，返回作者维度的聚合数据"""
    if df.empty or '作者昵称' not in df.columns:
        return pd.DataFrame()
    agg_dict = {'视频ID': 'count'}
    if '点赞数' in df:
        agg_dict['点赞数'] = 'sum'
    if '评论数' in df:
        agg_dict['评论数'] = 'sum'
    if '收藏数' in df:
        agg_dict['收藏数'] = 'sum'
    author_stats = df.groupby('作者昵称').agg(agg_dict).reset_index()
    author_stats.rename(columns={'视频ID': '发布数量'}, inplace=True)
    if '点赞数' in df:
        author_stats.rename(columns={'点赞数': '总点赞数'}, inplace=True)
    if '评论数' in df:
        author_stats.rename(columns={'评论数': '总评论数'}, inplace=True)
    if '收藏数' in df:
        author_stats.rename(columns={'收藏数': '总收藏数'}, inplace=True)
    if '粉丝数' in df.columns:
        fans = df.groupby('作者昵称')['粉丝数'].max().reset_index()
        fans.rename(columns={'粉丝数': '粉丝数'}, inplace=True)
        author_stats = author_stats.merge(fans, on='作者昵称', how='left')
    else:
        author_stats['粉丝数'] = None
    return author_stats

# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown(f"**数据文件**：`{DATA_FILE}`")

    # 加载原始数据（保留中文列名）
    raw_df = load_data(DATA_FILE)
    if raw_df.empty:
        st.stop()

    # 预处理分析用数据（添加派生列 publish_date，数值转换）
    df_ana = preprocess_for_analysis(raw_df)

    # ========== 侧边栏全局筛选 ==========
    st.sidebar.header("🔍 全局数据筛选")
    st.sidebar.markdown(f"**原始视频数**: {len(df_ana)}")

    # 日期范围筛选（处理单日期/双日期）
    filter_mask = pd.Series([True] * len(df_ana))
    if 'publish_date' in df_ana.columns:
        valid_dates = df_ana['publish_date'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_data_date = valid_dates.max().date()
            date_range = st.sidebar.date_input(
                "选择日期范围",
                value=(min_date, max_data_date),
                min_value=min_date,
                max_value=None
            )
            # 如果只选择了单一日，则转为开始=结束
            if isinstance(date_range, (date, pd.Timestamp)):
                start_date = end_date = date_range
            else:
                start_date, end_date = date_range
            mask = (df_ana['publish_date'].dt.date >= start_date) & (df_ana['publish_date'].dt.date <= end_date)
            filter_mask &= mask
            st.sidebar.info(f"筛选后视频数: {filter_mask.sum()}")

    # 数值指标滑块筛选
    for col_cn in ['点赞数', '评论数', '分享数', '收藏数']:
        if col_cn in df_ana.columns and not df_ana[col_cn].isna().all():
            min_val = float(df_ana[col_cn].min())
            max_val = float(df_ana[col_cn].max())
            if min_val < max_val:
                selected = st.sidebar.slider(f"{col_cn} 范围", min_val, max_val, (min_val, max_val))
                mask = (df_ana[col_cn] >= selected[0]) & (df_ana[col_cn] <= selected[1])
                filter_mask &= mask

    # 应用筛选
    df_filtered = df_ana[filter_mask].copy()
    raw_filtered = raw_df.loc[filter_mask] if filter_mask.dtype == bool else raw_df.iloc[filter_mask]

    # ========== KPI 指标卡片 ==========
    kpi = compute_kpis(df_filtered)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📹 视频总数", f"{kpi.get('total_videos', 0):,}")
    with col2:
        st.metric("❤️ 总点赞数", f"{kpi.get('total_likes', 0):,.0f}" if kpi.get('total_likes') else "无数据")
    with col3:
        st.metric("💬 总评论数", f"{kpi.get('total_comments', 0):,.0f}" if kpi.get('total_comments') else "无数据")
    with col4:
        st.metric("⭐ 平均点赞/视频", f"{kpi.get('avg_likes_per_video', 0):.1f}" if kpi.get('avg_likes_per_video') else "--")
    col5, col6, col7 = st.columns(3)
    with col5:
        st.metric("🔄 总分享数", f"{kpi.get('total_shares', 0):,.0f}" if kpi.get('total_shares') else "无数据")
    with col6:
        st.metric("🔖 总收藏数", f"{kpi.get('total_collects', 0):,.0f}" if kpi.get('total_collects') else "无数据")
    with col7:
        st.metric("👥 覆盖粉丝数", f"{kpi.get('total_fans', 0):,.0f}" if kpi.get('total_fans') else "无数据")

    st.markdown("---")

    # ========== 趋势图 ==========
    if 'publish_date' in df_filtered and '点赞数' in df_filtered:
        st.subheader("📈 每日点赞数趋势")
        fig = plot_trend(df_filtered, 'publish_date', '点赞数', "每日点赞总数变化", '#E63946')
        if fig:
            st.plotly_chart(fig, width='stretch')

    if 'publish_date' in df_filtered:
        st.subheader("📆 每日发布视频数量趋势")
        daily_videos = df_filtered.groupby(df_filtered['publish_date'].dt.date).size().reset_index(name='发布数量')
        fig2 = px.line(daily_videos, x='publish_date', y='发布数量', title="每日发布视频数量变化", markers=True)
        fig2.update_traces(line_color='#2E86C1')
        st.plotly_chart(fig2, width='stretch')

    # 双列布局
    left, right = st.columns(2)
    with left:
        st.subheader("🏆 点赞数 Top10 视频")
        fig = plot_top10(df_filtered, '点赞数', title_col='视频描述', title="点赞量最高的10个视频")
        if fig:
            st.plotly_chart(fig, width='stretch')
    with right:
        st.subheader("📊 点赞 vs 评论 关系")
        fig = plot_scatter(df_filtered, '点赞数', '评论数',
                           size_col='分享数' if '分享数' in df_filtered else None,
                           title="点赞数与评论数散点图")
        if fig:
            st.plotly_chart(fig, width='stretch')

    # ========== 作者双榜 ==========
    st.subheader("👥 作者维度综合排行榜")
    author_df = get_author_aggregation(df_filtered)
    if not author_df.empty:
        tab1, tab2 = st.tabs(["📦 发布数量 Top10", "❤️ 总点赞数 Top10"])
        with tab1:
            top_publish = author_df.sort_values('发布数量', ascending=False).head(10)
            st.dataframe(top_publish, use_container_width=True)
            fig_pub = px.bar(top_publish, x='作者昵称', y='发布数量', title="发布数量 Top10 作者",
                             text_auto=True, color='发布数量')
            st.plotly_chart(fig_pub, width='stretch')
        with tab2:
            if '总点赞数' in author_df.columns:
                top_likes = author_df.sort_values('总点赞数', ascending=False).head(10)
                st.dataframe(top_likes, use_container_width=True)
                fig_likes = px.bar(top_likes, x='作者昵称', y='总点赞数', title="总点赞数 Top10 作者",
                                   text_auto=True, color='总点赞数')
                st.plotly_chart(fig_likes, width='stretch')
            else:
                st.info("数据中不含点赞数，无法展示点赞榜。")
    else:
        st.info("未找到作者昵称列，无法进行作者排行榜分析。")

    # ==================== 优化后的单日发布监控（含未发布筛选） ====================
    st.subheader("🔍 单日发布监控（支持筛选未发布作者）")
    st.markdown("选择日期，查看每位作者当天的视频发布数量，并可筛选显示未发布/已发布作者。")

    # 基于原始数据（不受全局筛选影响）以便查看所有作者的发布状态
    # 使用原始 raw_df（未经筛选）进行统计，保证能看到所有作者
    if '创建时间' in raw_df.columns and '作者昵称' in raw_df.columns:
        # 转换日期
        raw_df_date = raw_df.copy()
        raw_df_date['publish_date'] = pd.to_datetime(raw_df_date['创建时间'], errors='coerce')
        valid_pub = raw_df_date['publish_date'].dropna()

        if not valid_pub.empty:
            min_date_all = valid_pub.min().date()
            max_date_all = valid_pub.max().date()
            default_date = max_date_all
            selected_date = st.date_input(
                "📅 选择检查日期",
                value=default_date,
                min_value=min_date_all,
                max_value=None
            )

            # 统计当天各作者发布数量
            mask_today = (raw_df_date['publish_date'].dt.date == selected_date)
            daily_stats = raw_df_date[mask_today].groupby('作者昵称').size().reset_index(name='当天发布数')

            # 获取所有作者（去重）
            all_authors = raw_df_date['作者昵称'].dropna().unique()
            all_authors = sorted(all_authors)
            author_status = pd.DataFrame({'作者昵称': all_authors})
            author_status = author_status.merge(daily_stats, on='作者昵称', how='left')
            author_status['当天发布数'] = author_status['当天发布数'].fillna(0).astype(int)
            author_status['发布状态'] = author_status['当天发布数'].apply(lambda x: '✅ 已发布' if x > 0 else '❌ 未发布')

            # 筛选选项
            filter_option = st.radio(
                "筛选作者：",
                ["全部作者", "仅未发布作者 (当天发布数为0)", "仅已发布作者 (当天发布数>0)"],
                horizontal=True
            )
            if filter_option == "仅未发布作者 (当天发布数为0)":
                author_status = author_status[author_status['当天发布数'] == 0]
            elif filter_option == "仅已发布作者 (当天发布数>0)":
                author_status = author_status[author_status['当天发布数'] > 0]

            # 显示结果
            st.write(f"### {selected_date} 作者发布情况 (共 {len(author_status)} 位)")
            st.dataframe(author_status, use_container_width=True)

            # 摘要统计（基于原始完整的作者列表）
            total_authors = len(all_authors)
            published_authors = len(daily_stats)
            total_videos_today = daily_stats['当天发布数'].sum()
            st.info(f"📊 **摘要**：共 {total_authors} 位作者，其中 {published_authors} 位当天发布了视频，合计发布 {total_videos_today} 个视频。")
        else:
            st.info("原始数据中无有效发布日期，无法进行单日监控。")
    else:
        st.info("原始数据缺少'创建时间'或'作者昵称'列，无法进行单日监控。")

    # ========== 发布时间分析 ==========
    st.subheader("⏰ 发布时间分析")
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        if 'publish_date' in df_filtered:
            df_week = df_filtered['publish_date'].dropna().dt.dayofweek
            if not df_week.empty:
                week_counts = df_week.value_counts().reindex(range(7), fill_value=0)
                week_labels = ['周一','周二','周三','周四','周五','周六','周日']
                week_df = pd.DataFrame({'星期': week_labels, '视频数': week_counts.values})
                fig = px.bar(week_df, x='星期', y='视频数', title="按星期发布数量（筛选后数据）")
                st.plotly_chart(fig, width='stretch')
    with col_h2:
        fig = plot_publish_hour_heatmap(df_filtered, 'publish_date')
        if fig:
            st.plotly_chart(fig, width='stretch')

    # ========== 原始数据预览（使用筛选后的原始中文数据） ==========
    st.subheader("📄 原始数据预览（筛选后）")
    display_cols_cn = ['作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数', '创建时间']
    existing_cn = [c for c in display_cols_cn if c in raw_filtered.columns]
    if existing_cn:
        preview_df = raw_filtered[existing_cn]
    else:
        preview_df = raw_filtered
    show_all = st.checkbox("显示全部数据（默认仅显示前100行）")
    if show_all:
        st.dataframe(preview_df, use_container_width=True)
    else:
        st.dataframe(preview_df.head(100), use_container_width=True)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info("看板支持单日发布监控（可筛选未发布作者）、作者双排行榜，日期选择无上限。")

if __name__ == "__main__":
    main()