"""
抖音视频数据可视化看板（增强版 - 支持姓名/工号/抖音号，昵称为空自动归零）
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
import numpy as np

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
    """
    预处理：转换数值类型、解析日期、处理昵称空值、添加姓名/工号/抖音号兜底、创建分组键
    对于无效昵称（为空或纯空白），强制所有指标为0，视频ID置NaN，不计入发布数量。
    抖音号字段用于辅助区分无效昵称的作者（避免错误合并）。
    """
    if df.empty:
        return df
    df_work = df.copy()

    # 1. 确保姓名、工号列存在（若原数据没有则补空白）
    if '姓名' not in df_work.columns:
        df_work['姓名'] = ''
        st.info("⚠️ 原始数据中缺少「姓名」列，已在看板中补空白，请后续补充映射关系。")
    if '工号' not in df_work.columns:
        df_work['工号'] = ''
        st.info("⚠️ 原始数据中缺少「工号」列，已在看板中补空白，请后续补充映射关系。")

    # 1.1 确保抖音号列存在（若无则补空白，用于辅助区分无效昵称的作者）
    if '抖音号' not in df_work.columns:
        df_work['抖音号'] = ''
        st.info("⚠️ 原始数据中缺少「抖音号」列，已补空白，仅供展示与辅助区分。")

    # 2. 数值列转换，缺失填充0
    num_cols_cn = ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '获赞总数']
    for col in num_cols_cn:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0)

    # 3. 处理作者昵称：空值或纯空白视为无效
    if '作者昵称' in df_work.columns:
        # 将可能的NaN、None转为字符串并去除首尾空白
        df_work['作者昵称'] = df_work['作者昵称'].astype(str).str.strip()
        # 有效条件：既不是空字符串也不是'nan'
        df_work['has_valid_nickname'] = (df_work['作者昵称'] != '') & (df_work['作者昵称'] != 'nan')
    else:
        df_work['作者昵称'] = ''
        df_work['has_valid_nickname'] = False
        st.warning("原始数据中缺少「作者昵称」列，将无法正确识别抖音账号。")

    # 4. 对于无有效昵称的行，强制所有指标为0，并且不计入发布数量（视频ID置为NaN）
    invalid_mask = ~df_work['has_valid_nickname']
    if invalid_mask.any():
        for col in num_cols_cn:
            if col in df_work.columns:
                df_work.loc[invalid_mask, col] = 0
        if '视频ID' in df_work.columns:
            df_work.loc[invalid_mask, '视频ID'] = np.nan
        # 无效昵称的统一显示为特殊标记
        df_work.loc[invalid_mask, '作者昵称'] = '(无抖音号)'

    # 5. 创建分组键（关键修复：无效昵称使用「姓名+工号+抖音号」组合，避免错误合并）
    # 先准备姓名、工号、抖音号字段（缺失值填空字符串）
    name_series = df_work['姓名'].fillna('').astype(str)
    id_series = df_work['工号'].fillna('').astype(str)
    douyin_series = df_work['抖音号'].fillna('').astype(str)

    # 有效昵称：分组键 = 作者昵称（唯一标识一个抖音账号）
    valid_mask = df_work['has_valid_nickname']
    df_work.loc[valid_mask, 'author_group_key'] = df_work.loc[valid_mask, '作者昵称']

    # 无效昵称：分组键 = 姓名_工号_抖音号（保证不同员工/不同抖音号不被错误合并）
    if invalid_mask.any():
        combined_key = name_series + '_' + id_series + '_' + douyin_series
        # 极端情况：如果三者全都是空，则使用行索引作为唯一后缀，避免合并所有未知员工为一个
        all_empty_mask = (name_series == '') & (id_series == '') & (douyin_series == '') & invalid_mask
        combined_key = combined_key.where(~all_empty_mask, combined_key + '_' + df_work.index.astype(str))
        df_work.loc[invalid_mask, 'author_group_key'] = combined_key.loc[invalid_mask]

    # 6. 解析发布日期
    if '创建时间' in df_work.columns:
        df_work['publish_date'] = pd.to_datetime(df_work['创建时间'], errors='coerce')
    else:
        df_work['publish_date'] = pd.NaT

    return df_work

def get_author_aggregation(df):
    """
    基于预处理后的 df 进行聚合统计，返回包含 姓名、工号、抖音号、作者昵称、发布数量、总点赞数、总评论数、总收藏数、粉丝数 的 DataFrame
    """
    if df.empty or 'author_group_key' not in df.columns:
        return pd.DataFrame()

    # 定义聚合字典
    agg_dict = {
        '作者昵称': 'first',
        '姓名': 'first',
        '工号': 'first',
        '抖音号': 'first',
        '视频ID': 'count',
    }
    if '点赞数' in df.columns:
        agg_dict['点赞数'] = 'sum'
    if '评论数' in df.columns:
        agg_dict['评论数'] = 'sum'
    if '收藏数' in df.columns:
        agg_dict['收藏数'] = 'sum'
    if '粉丝数' in df.columns:
        agg_dict['粉丝数'] = 'max'

    author_stats = df.groupby('author_group_key').agg(agg_dict).reset_index(drop=False)
    # 重命名为友好中文名
    rename_map = {
        'author_group_key': '分组键',
        '作者昵称': '作者昵称',
        '姓名': '姓名',
        '工号': '工号',
        '抖音号': '抖音号',
        '视频ID': '发布数量',
        '点赞数': '总点赞数',
        '评论数': '总评论数',
        '收藏数': '总收藏数',
        '粉丝数': '粉丝数'
    }
    author_stats.rename(columns=rename_map, inplace=True)

    # 补零
    for col in ['发布数量', '总点赞数', '总评论数', '总收藏数', '粉丝数']:
        if col in author_stats.columns:
            author_stats[col] = author_stats[col].fillna(0).astype(int)
        else:
            author_stats[col] = 0

    return author_stats

def find_video_link_column(df):
    possible_names = ['视频链接', '链接', '分享链接', 'video_link', 'url', 'link']
    for col in df.columns:
        if col in possible_names:
            return col
    return None

def reorder_columns_with_douyin(df):
    """将抖音号列移到工号后面，其他列顺序不变"""
    if df.empty or '抖音号' not in df.columns:
        return df
    cols = df.columns.tolist()
    if '工号' in cols:
        idx = cols.index('工号')
        if idx+1 >= len(cols) or cols[idx+1] != '抖音号':
            cols.remove('抖音号')
            cols.insert(idx+1, '抖音号')
    return df[cols]

# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown(f"**数据源**：GitHub 仓库（实时同步） | 支持姓名/工号/抖音号，无抖音号员工自动归零")

    raw_df = load_data_from_github(GITHUB_RAW_URL)
    if raw_df.empty:
        st.stop()

    # 预处理
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

    # ======================== 作者双榜 → 三榜 ========================
    st.subheader("👥 作者维度综合排行榜")
    st.caption("以下统计基于左侧全局筛选后的数据 | 发布数量 = 有效视频ID计数 | 无抖音号员工自动显示昵称为(无抖音号)且各项指标为0")
    author_df = get_author_aggregation(df_filtered)
    if not author_df.empty:
        # 确定展示列顺序（先姓名、工号、抖音号、作者昵称、然后指标）
        base_cols = ['姓名', '工号']
        if '抖音号' in author_df.columns:
            base_cols.append('抖音号')
        base_cols.append('作者昵称')
        base_cols.append('发布数量')
        for col in ['总点赞数', '总评论数', '总收藏数', '粉丝数']:
            if col in author_df.columns:
                base_cols.append(col)
        display_cols_order = [c for c in base_cols if c in author_df.columns]

        # 创建三个标签页
        tab1, tab2, tab3 = st.tabs(["📦 发布数量 Top10", "❤️ 总点赞数 Top10", "👥 粉丝数 Top10"])

        with tab1:
            top_publish = author_df.sort_values('发布数量', ascending=False).head(10)
            top_publish_display = reorder_columns_with_douyin(top_publish[display_cols_order])
            st.dataframe(top_publish_display, width='stretch')
            fig_pub = px.bar(top_publish, x='作者昵称', y='发布数量',
                             title="发布数量 Top10",
                             text_auto=True, color='发布数量',
                             hover_data=['姓名', '工号'])
            st.plotly_chart(fig_pub, use_container_width=True)

        with tab2:
            if '总点赞数' in author_df.columns:
                top_likes = author_df.sort_values('总点赞数', ascending=False).head(10)
                top_likes_display = reorder_columns_with_douyin(top_likes[display_cols_order])
                st.dataframe(top_likes_display, width='stretch')
                fig_likes = px.bar(top_likes, x='作者昵称', y='总点赞数',
                                   title="总点赞数 Top10",
                                   text_auto=True, color='总点赞数',
                                   hover_data=['姓名', '工号'])
                st.plotly_chart(fig_likes, use_container_width=True)
            else:
                st.info("数据中不含点赞数，无法展示点赞榜。")

        with tab3:
            if '粉丝数' in author_df.columns:
                top_fans = author_df.sort_values('粉丝数', ascending=False).head(10)
                top_fans_display = reorder_columns_with_douyin(top_fans[display_cols_order])
                st.dataframe(top_fans_display, width='stretch')
                fig_fans = px.bar(top_fans, x='作者昵称', y='粉丝数',
                                  title="粉丝数 Top10",
                                  text_auto=True, color='粉丝数',
                                  hover_data=['姓名', '工号'])
                st.plotly_chart(fig_fans, use_container_width=True)
            else:
                st.info("数据中不含粉丝数，无法展示粉丝榜。")
    else:
        st.info("未找到作者数据，无法进行作者排行榜分析。")

    # ======================== 发布监控 ========================
    st.subheader("🔍 发布监控（支持单日或日期范围）")
    st.caption("基于原始全量数据统计，不受左侧筛选影响。无抖音号员工（昵称为空）发布数量恒为0。")

    df_monitor = preprocess_for_analysis(raw_df.copy())
    if 'publish_date' in df_monitor.columns and 'author_group_key' in df_monitor.columns:
        valid_pub = df_monitor['publish_date'].dropna()
        if not valid_pub.empty:
            min_date_all = valid_pub.min().date()
            max_date_all = valid_pub.max().date()
            today = date.today()
            yesterday = today - timedelta(days=1)
            if yesterday > max_date_all:
                default_single_date = max_date_all
            elif yesterday < min_date_all:
                default_single_date = min_date_all
            else:
                default_single_date = yesterday

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
                mask_selected = (df_monitor['publish_date'].dt.date == selected_date)
                date_desc = selected_date.strftime("%Y-%m-%d")
                count_label = "当天发布数"
                title_suffix = f"（{date_desc}）"
            else:
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
                mask_selected = (df_monitor['publish_date'].dt.date >= start_date) & \
                                (df_monitor['publish_date'].dt.date <= end_date)
                date_desc = f"{start_date} 至 {end_date}"
                count_label = "范围内发布数"
                title_suffix = f"（{date_desc}）"

            selected_videos = df_monitor[mask_selected].copy()

            # 获取所有作者基本信息（分组键、姓名、工号、抖音号、昵称）
            info_cols = ['author_group_key', '姓名', '工号', '作者昵称']
            if '抖音号' in df_monitor.columns:
                info_cols.append('抖音号')
            all_authors_info = df_monitor[info_cols].drop_duplicates('author_group_key')

            if not selected_videos.empty:
                daily_stats = selected_videos.groupby('author_group_key').agg(
                    发布数量=('视频ID', 'count')
                ).reset_index()
                author_status = all_authors_info.merge(daily_stats, on='author_group_key', how='left')
            else:
                author_status = all_authors_info.copy()
                author_status['发布数量'] = 0

            author_status[count_label] = author_status['发布数量'].fillna(0).astype(int)
            if '发布数量' in author_status.columns:
                author_status.drop(columns=['发布数量'], inplace=True)
            author_status['发布状态'] = author_status[count_label].apply(lambda x: '✅ 有发布' if x > 0 else '❌ 无发布')

            for col in ['姓名', '工号', '作者昵称']:
                if col not in author_status.columns:
                    author_status[col] = ''
            if '抖音号' not in author_status.columns:
                author_status['抖音号'] = ''

            # 备注：当作者昵称为'(无抖音号)'时提示核查
            author_status['备注'] = author_status['作者昵称'].apply(
                lambda x: '⚠️ 该员工抖音号存在问题，请核查是否正确' if x == '(无抖音号)' else ''
            )

            # 调整列顺序：姓名、工号、抖音号、作者昵称、发布数、状态、备注
            display_cols = ['姓名', '工号']
            if '抖音号' in author_status.columns:
                display_cols.append('抖音号')
            display_cols.extend(['作者昵称', count_label, '发布状态', '备注'])
            author_status = author_status[[c for c in display_cols if c in author_status.columns]]

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
            st.dataframe(author_status, width='stretch')

            total_authors = len(all_authors_info)
            published_authors = (author_status[count_label] > 0).sum() if not author_status.empty else 0
            total_videos = author_status[count_label].sum() if not author_status.empty else 0
            st.info(f"📊 **摘要**：共 {total_authors} 位作者，其中 {published_authors} 位有发布，合计发布 {total_videos} 个视频。")

            if not selected_videos.empty:
                with st.expander("📹 点击查看视频详情（验证发布数量）"):
                    author_list = sorted(selected_videos['author_group_key'].unique())
                    if author_list:
                        author_display = {}
                        for key in author_list:
                            row = selected_videos[selected_videos['author_group_key'] == key].iloc[0]
                            author_display[key] = f"{row['姓名']}({row['工号']}) - {row['作者昵称']}"
                        selected_display = st.selectbox("选择作者查看其发布的视频", options=author_list, format_func=lambda x: author_display[x], key="author_detail")
                        author_videos = selected_videos[selected_videos['author_group_key'] == selected_display].copy()
                        author_videos = author_videos.sort_values('publish_date', ascending=False)
                        st.write(f"**{author_display[selected_display]}** 在 {date_desc} 发布了 {len(author_videos)} 个视频：")

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
                        # 显示姓名、工号、抖音号在前
                        for prefix in ['姓名', '工号', '抖音号']:
                            if prefix in author_videos.columns and prefix not in display_cols:
                                display_cols.insert(0, prefix)
                        display_cols = list(dict.fromkeys(display_cols))

                        column_config = {}
                        if link_col and link_col in author_videos.columns:
                            column_config[link_col] = st.column_config.LinkColumn("视频链接", width="small", help="点击跳转")
                        st.dataframe(author_videos[display_cols], column_config=column_config, width='stretch')
            else:
                st.info("所选日期范围内没有作者发布视频。")
        else:
            st.info("原始数据中无有效发布日期，无法进行发布监控。")
    else:
        st.info("原始数据缺少'创建时间'或'作者昵称'列，无法进行发布监控。")

    # ======================== 原始数据预览 ========================
    st.subheader("📄 原始数据预览（应用全局筛选后）")
    base_cols = ['姓名', '工号', '抖音号', '作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数', '创建时间']
    if link_col:
        base_cols.append(link_col)
    existing_cols = [c for c in base_cols if c in raw_filtered.columns]
    preview_df = raw_filtered[existing_cols] if existing_cols else raw_filtered.copy()
    preview_df = reorder_columns_with_douyin(preview_df)
    column_config = {}
    if link_col and link_col in preview_df.columns:
        column_config[link_col] = st.column_config.LinkColumn("视频链接", width="small", help="点击跳转")
    show_all = st.checkbox("显示全部数据（默认仅显示前100行）")
    if show_all:
        st.dataframe(preview_df, column_config=column_config, width='stretch')
    else:
        st.dataframe(preview_df.head(100), column_config=column_config, width='stretch')

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info(
        "**姓名/工号/抖音号**：从Excel中读取，若缺失则显示空白。\n\n"
        "**无抖音号员工**：当「作者昵称」为空时，该员工所有指标自动为0，且发布数量不计入任何统计。\n\n"
        "**发布数量定义**：有效抖音号的视频ID计数。\n\n"
        "**发布监控**：支持单天或范围模式，默认显示昨日数据。\n\n"
        "**视频链接**：若Excel中包含‘视频链接’等列，自动显示可点击链接。"
    )

if __name__ == "__main__":
    main()