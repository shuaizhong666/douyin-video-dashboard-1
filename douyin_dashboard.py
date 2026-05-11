"""
抖音视频数据可视化看板（增强版 - 默认昨日数据）
数据源：GitHub 仓库中的 Excel 文件（通过 Raw URL 实时读取）
运行命令：streamlit run douyin_dashboard.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO
import warnings
from datetime import date, datetime, timedelta
import tempfile
import os

warnings.filterwarnings('ignore')

# ======================== 页面配置 ========================
st.set_page_config(
    page_title="抖音视频数据分析看板",
    page_icon="🎵",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================== 数据源配置 ========================
GITHUB_RAW_URL = "https://raw.githubusercontent.com/shuaizhong666/douyin-video-dashboard-1/main/抖音视频数据汇总.xlsx"

# ======================== 加载函数 ========================
@st.cache_data(ttl=180, show_spinner="正在从 GitHub 加载数据...")
def load_data_from_github(url, max_retries=3):
    for attempt in range(max_retries):
        try:
            random_param = f"?_={int(datetime.now().timestamp())}" if attempt > 0 else ""
            final_url = url + random_param
            headers = {"User-Agent": "Mozilla/5.0"}
            response = requests.get(final_url, timeout=30, headers=headers)
            response.raise_for_status()
            content = response.content
            try:
                df = pd.read_excel(BytesIO(content), engine='openpyxl')
                if df.empty:
                    st.warning("Excel 文件读取成功，但数据为空。")
                else:
                    st.success(f"✅ 数据加载成功，共 {len(df)} 行记录")
                return df
            except Exception as e1:
                st.warning(f"pandas 直接读取失败: {e1}")
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                        tmp.write(content)
                        tmp_path = tmp.name
                    from openpyxl import load_workbook
                    wb = load_workbook(tmp_path, read_only=True)
                    sheet = wb.active
                    data = sheet.values
                    cols = next(data)
                    df = pd.DataFrame(data, columns=cols)
                    wb.close()
                    os.unlink(tmp_path)
                    st.success("✅ 通过 openpyxl 直接加载成功")
                    return df
                except Exception as e2:
                    st.error(f"临时文件读取失败: {e2}")
                    raise e2
        except Exception as e:
            st.error(f"加载失败 (尝试 {attempt+1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return pd.DataFrame()
    return pd.DataFrame()

def preprocess_for_analysis(df):
    if df.empty:
        return df
    df_work = df.copy()
    num_cols_cn = ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '获赞总数']
    for col in num_cols_cn:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce')
    if '创建时间' in df_work.columns:
        df_work['publish_date'] = pd.to_datetime(df_work['创建时间'], errors='coerce')
    return df_work

def get_author_aggregation(df):
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
        author_stats = author_stats.merge(fans, on='作者昵称', how='left')
    else:
        author_stats['粉丝数'] = None
    return author_stats

def find_video_link_column(df):
    possible_names = ['视频链接', '链接', '分享链接', 'video_link', 'url', 'link']
    for col in df.columns:
        if col in possible_names:
            return col
    return None

# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown(f"**数据源**：GitHub 仓库（实时同步）")

    raw_df = load_data_from_github(GITHUB_RAW_URL)
    if raw_df.empty:
        st.stop()

    df_ana = preprocess_for_analysis(raw_df)
    link_col = find_video_link_column(raw_df)

    # 侧边栏全局筛选
    st.sidebar.header("🔍 全局数据筛选")
    st.sidebar.markdown(f"**原始视频数**: {len(df_ana)}")
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
            if isinstance(date_range, (date, pd.Timestamp)):
                start_date = end_date = date_range
            else:
                start_date, end_date = date_range
            mask = (df_ana['publish_date'].dt.date >= start_date) & (df_ana['publish_date'].dt.date <= end_date)
            filter_mask &= mask
            st.sidebar.info(f"筛选后视频数: {filter_mask.sum()}")

    for col_cn in ['点赞数', '评论数', '分享数', '收藏数']:
        if col_cn in df_ana.columns and not df_ana[col_cn].isna().all():
            min_val = float(df_ana[col_cn].min())
            max_val = float(df_ana[col_cn].max())
            if min_val < max_val:
                selected = st.sidebar.slider(f"{col_cn} 范围", min_val, max_val, (min_val, max_val))
                mask = (df_ana[col_cn] >= selected[0]) & (df_ana[col_cn] <= selected[1])
                filter_mask &= mask

    df_filtered = df_ana[filter_mask].copy()
    raw_filtered = raw_df.loc[filter_mask] if filter_mask.dtype == bool else raw_df.iloc[filter_mask]

    # 作者双榜
    st.subheader("👥 作者维度综合排行榜")
    st.caption("以下统计基于左侧全局筛选后的数据，发布数量 = 视频ID计数")
    author_df = get_author_aggregation(df_filtered)
    if not author_df.empty:
        tab1, tab2 = st.tabs(["📦 发布数量 Top10", "❤️ 总点赞数 Top10"])
        with tab1:
            top_publish = author_df.sort_values('发布数量', ascending=False).head(10)
            st.dataframe(top_publish, use_container_width=True)
            fig_pub = px.bar(top_publish, x='作者昵称', y='发布数量', title="发布数量 Top10 作者",
                             text_auto=True, color='发布数量')
            st.plotly_chart(fig_pub, use_container_width=True)
        with tab2:
            if '总点赞数' in author_df.columns:
                top_likes = author_df.sort_values('总点赞数', ascending=False).head(10)
                st.dataframe(top_likes, use_container_width=True)
                fig_likes = px.bar(top_likes, x='作者昵称', y='总点赞数', title="总点赞数 Top10 作者",
                                   text_auto=True, color='总点赞数')
                st.plotly_chart(fig_likes, use_container_width=True)
            else:
                st.info("数据中不含点赞数，无法展示点赞榜。")
    else:
        st.info("未找到作者昵称列，无法进行作者排行榜分析。")

    # ========== 单日/范围发布监控 ==========
    st.subheader("🔍 发布监控（支持单日或日期范围）")
    st.caption("基于原始全量数据统计，不受左侧筛选影响。")

    if '创建时间' in raw_df.columns and '作者昵称' in raw_df.columns:
        raw_df_date = raw_df.copy()
        raw_df_date['publish_date'] = pd.to_datetime(raw_df_date['创建时间'], errors='coerce')
        valid_pub = raw_df_date['publish_date'].dropna()
        if not valid_pub.empty:
            min_date_all = valid_pub.min().date()
            max_date_all = valid_pub.max().date()
            today = date.today()
            # 计算昨天日期（如果昨天大于最大日期，则取最大日期）
            yesterday = today - timedelta(days=1)
            if yesterday > max_date_all:
                default_single_date = max_date_all
            elif yesterday < min_date_all:
                default_single_date = min_date_all
            else:
                default_single_date = yesterday

            # 选择模式
            mode = st.radio(
                "📅 日期选择模式",
                ["单天模式", "范围模式"],
                horizontal=True,
                key="date_mode"
            )

            if mode == "单天模式":
                selected_date = st.date_input(
                    "选择检查日期",
                    value=default_single_date,
                    min_value=min_date_all,
                    max_value=None,
                    key="single_date"
                )
                # 过滤数据
                mask_selected = (raw_df_date['publish_date'].dt.date == selected_date)
                date_desc = selected_date.strftime("%Y-%m-%d")
                count_label = "当天发布数"
                title_suffix = f"（{date_desc}）"
            else:  # 范围模式
                date_range = st.date_input(
                    "选择日期范围",
                    value=(min_date_all, max_date_all),
                    min_value=min_date_all,
                    max_value=None,
                    key="range_dates"
                )
                if isinstance(date_range, (date, pd.Timestamp)):
                    start_date = end_date = date_range
                else:
                    start_date, end_date = date_range
                if start_date > end_date:
                    st.error("开始日期不能晚于结束日期")
                    st.stop()
                mask_selected = (raw_df_date['publish_date'].dt.date >= start_date) & \
                                (raw_df_date['publish_date'].dt.date <= end_date)
                date_desc = f"{start_date} 至 {end_date}"
                count_label = "范围内发布数"
                title_suffix = f"（{date_desc}）"

            # 获取选定范围内的数据
            selected_videos = raw_df_date[mask_selected].copy()
            # 按作者统计
            daily_stats = selected_videos.groupby('作者昵称').size().reset_index(name=count_label)

            all_authors = sorted(raw_df_date['作者昵称'].dropna().unique())
            author_status = pd.DataFrame({'作者昵称': all_authors})
            author_status = author_status.merge(daily_stats, on='作者昵称', how='left')
            author_status[count_label] = author_status[count_label].fillna(0).astype(int)
            author_status['发布状态'] = author_status[count_label].apply(lambda x: '✅ 有发布' if x > 0 else '❌ 无发布')

            filter_option = st.radio(
                "筛选作者：",
                ["全部作者", f"仅无发布作者 ({count_label}=0)", f"仅有发布作者 ({count_label}>0)"],
                horizontal=True
            )
            if filter_option == f"仅无发布作者 ({count_label}=0)":
                author_status = author_status[author_status[count_label] == 0]
            elif filter_option == f"仅有发布作者 ({count_label}>0)":
                author_status = author_status[author_status[count_label] > 0]

            st.write(f"### 作者发布统计 {title_suffix}")
            st.dataframe(author_status, use_container_width=True)

            total_authors = len(all_authors)
            published_authors = len(daily_stats)
            total_videos = daily_stats[count_label].sum() if not daily_stats.empty else 0
            st.info(f"📊 **摘要**：共 {total_authors} 位作者，其中 {published_authors} 位有发布，合计发布 {total_videos} 个视频。")

            # 查看视频详情
            if not selected_videos.empty:
                with st.expander("📹 点击查看视频详情（验证发布数量）"):
                    author_with_videos = sorted(selected_videos['作者昵称'].unique())
                    selected_author = st.selectbox("选择作者查看其发布的视频", options=author_with_videos, key="author_detail")
                    author_videos = selected_videos[selected_videos['作者昵称'] == selected_author].copy()
                    # 按时间排序
                    author_videos = author_videos.sort_values('publish_date', ascending=False)
                    st.write(f"**{selected_author}** 在 {date_desc} 发布了 {len(author_videos)} 个视频：")

                    # 显示列
                    display_cols = []
                    if '视频ID' in author_videos.columns:
                        display_cols.append('视频ID')
                    if '视频描述' in author_videos.columns:
                        display_cols.append('视频描述')
                    if '创建时间' in author_videos.columns:
                        display_cols.append('创建时间')
                    if link_col and link_col in author_videos.columns:
                        display_cols.append(link_col)
                    if not display_cols:
                        display_cols = author_videos.columns.tolist()

                    column_config = {}
                    if link_col and link_col in author_videos.columns:
                        column_config[link_col] = st.column_config.LinkColumn(
                            "视频链接",
                            width="small",
                            help="点击跳转"
                        )
                    st.dataframe(author_videos[display_cols], column_config=column_config, use_container_width=True)
            else:
                st.info("所选日期范围内没有作者发布视频。")
        else:
            st.info("原始数据中无有效发布日期，无法进行发布监控。")
    else:
        st.info("原始数据缺少'创建时间'或'作者昵称'列，无法进行发布监控。")

    # 原始数据预览
    st.subheader("📄 原始数据预览（应用全局筛选后）")
    base_cols = ['作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数', '创建时间']
    if link_col:
        base_cols.append(link_col)
    existing_cols = [c for c in base_cols if c in raw_filtered.columns]
    preview_df = raw_filtered[existing_cols] if existing_cols else raw_filtered.copy()
    column_config = {}
    if link_col and link_col in preview_df.columns:
        column_config[link_col] = st.column_config.LinkColumn("视频链接", width="small", help="点击跳转")
    show_all = st.checkbox("显示全部数据（默认仅显示前100行）")
    if show_all:
        st.dataframe(preview_df, column_config=column_config, use_container_width=True)
    else:
        st.dataframe(preview_df.head(100), column_config=column_config, use_container_width=True)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info(
        "**发布数量定义**：视频ID或记录条数。\n\n"
        "**发布监控**：支持单天或日期范围，基于原始全量数据。单天模式默认显示昨日数据。\n\n"
        "**视频链接**：若Excel中包含‘视频链接’等列，自动显示可点击链接。\n\n"
        "**数据缓存**：每次刷新页面最多延迟1小时（GitHub Raw CDN）。"
    )

if __name__ == "__main__":
    main()