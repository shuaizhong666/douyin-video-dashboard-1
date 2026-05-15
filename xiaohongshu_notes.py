"""
小红书笔记数据可视化看板（单数据源版）
数据源：GitHub 仓库中的 Excel 文件（仅小红书笔记数据.xlsx）
运行命令：streamlit run xiaohongshu_dashboard.py
字段说明：获赞与收藏数（Excel 中作者主页总获赞收藏数，聚合取 max）
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
    page_title="小红书笔记数据分析看板",
    page_icon="📕",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================== 数据源配置 ========================
GITHUB_BASE_URL = "https://raw.githubusercontent.com/shuaizhong666/douyin-video-dashboard-1/main/小红书笔记数据.xlsx"

# ======================== 加载函数 ========================
@st.cache_data(ttl=180, show_spinner="正在从 GitHub 加载数据...")
def load_single_excel_from_github(url, max_retries=3):
    """加载 Excel 文件（带重试机制）"""
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
                return df
            except Exception as e1:
                # 降级：通过 openpyxl 直接读取
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
                return df
        except Exception as e:
            st.warning(f"加载数据失败 (尝试 {attempt+1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return pd.DataFrame()
    return pd.DataFrame()

@st.cache_data(ttl=180, show_spinner="正在加载小红书数据...")
def load_data():
    """
    加载小红书笔记数据
    返回 DataFrame
    """
    df = load_single_excel_from_github(GITHUB_BASE_URL)
    if df.empty:
        st.error("❌ 无法加载数据，请检查网络或文件地址。")
        return pd.DataFrame()
    st.success(f"✅ 数据加载成功：{len(df)} 条记录")
    return df

# ======================== 预处理函数 ========================
def preprocess_for_analysis(df):
    """
    预处理：转换数值类型、解析日期、处理用户昵称/小红书号空值。
    获赞与收藏数直接从Excel读取，视为作者静态属性。
    """
    if df.empty:
        return df
    df_work = df.copy()

    # 姓名、工号列存在性
    if '姓名' not in df_work.columns:
        df_work['姓名'] = ''
    if '工号' not in df_work.columns:
        df_work['工号'] = ''

    # 小红书号：尝试常见列名
    xhs_id_col = None
    possible_xhs_names = ['小红书号', '小红书ID', 'xhs_id', '小红书账号', 'account_id']
    for col in possible_xhs_names:
        if col in df_work.columns:
            xhs_id_col = col
            break
    if xhs_id_col is None:
        df_work['小红书号'] = ''
        st.warning("原始数据中未找到「小红书号」相关列，将显示为空。")
    else:
        df_work.rename(columns={xhs_id_col: '小红书号'}, inplace=True)
        df_work['小红书号'] = df_work['小红书号'].astype(str).str.strip()

    # 数值列转换
    num_cols_cn = ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '关注数', '获赞与收藏数']
    for col in num_cols_cn:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0)
        else:
            if col == '获赞与收藏数':
                df_work[col] = 0
                st.warning("原始数据中缺少「获赞与收藏数」列，已自动填充0。")
            elif col == '关注数':
                df_work[col] = 0

    # 处理用户昵称
    if '用户昵称' in df_work.columns:
        df_work['用户昵称'] = df_work['用户昵称'].astype(str).str.strip()
        df_work['has_valid_nickname'] = (df_work['用户昵称'] != '') & (df_work['用户昵称'] != 'nan')
    else:
        df_work['用户昵称'] = ''
        df_work['has_valid_nickname'] = False
        st.warning("原始数据中缺少「用户昵称」列，将无法正确识别小红书账号。")

    # 无效昵称行强制归零
    invalid_mask = ~df_work['has_valid_nickname']
    if invalid_mask.any():
        for col in ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '关注数', '获赞与收藏数']:
            if col in df_work.columns:
                df_work.loc[invalid_mask, col] = 0
        if '笔记ID' in df_work.columns:
            df_work.loc[invalid_mask, '笔记ID'] = np.nan
        df_work.loc[invalid_mask, '用户昵称'] = '(无小红书号)'

    # 分组键
    df_work['author_group_key'] = df_work['用户昵称']
    if invalid_mask.any():
        name_series = df_work.loc[invalid_mask, '姓名'].astype(str).fillna('')
        id_series = df_work.loc[invalid_mask, '工号'].astype(str).fillna('')
        df_work.loc[invalid_mask, 'author_group_key'] = name_series + '_' + id_series
        empty_key_mask = (df_work['author_group_key'] == '_') & invalid_mask
        df_work.loc[empty_key_mask, 'author_group_key'] = df_work.loc[empty_key_mask].index.astype(str)

    # 发布日期
    if '发布日期' in df_work.columns:
        df_work['publish_date'] = pd.to_datetime(df_work['发布日期'], errors='coerce')
    else:
        df_work['publish_date'] = pd.NaT

    return df_work

def get_author_aggregation(df):
    """
    聚合统计：发布数量（笔记ID计数），
    静态属性（姓名、工号、小红书号、用户昵称、粉丝数、关注数、获赞与收藏数）取 first/max。
    注意：获赞与收藏数是作者主页总数据，因此取 max 而非 sum。
    """
    if df.empty or 'author_group_key' not in df.columns:
        return pd.DataFrame()

    agg_dict = {
        '用户昵称': 'first',
        '姓名': 'first',
        '工号': 'first',
        '小红书号': 'first',
        '笔记ID': 'count',          # 发布数量
        '粉丝数': 'max',            # 静态属性取最大
        '关注数': 'max',            # 静态属性取最大
        '获赞与收藏数': 'max',      # 博主主页总获赞与收藏数，取最大
    }
    # 可选的其他笔记级指标（不用于静态展示，但保留以防后面要用）
    if '点赞数' in df.columns:
        agg_dict['点赞数'] = 'sum'
    if '评论数' in df.columns:
        agg_dict['评论数'] = 'sum'
    if '收藏数' in df.columns:
        agg_dict['收藏数'] = 'sum'

    author_stats = df.groupby('author_group_key').agg(agg_dict).reset_index(drop=False)
    rename_map = {
        'author_group_key': '分组键',
        '用户昵称': '用户昵称',
        '姓名': '姓名',
        '工号': '工号',
        '小红书号': '小红书号',
        '笔记ID': '发布数量',
        '点赞数': '总点赞数',
        '评论数': '总评论数',
        '收藏数': '总收藏数',
        '粉丝数': '粉丝数',
        '关注数': '关注数',
        '获赞与收藏数': '获赞与收藏数'
    }
    author_stats.rename(columns=rename_map, inplace=True)

    # 补零并确保整数类型
    for col in ['发布数量', '粉丝数', '关注数', '获赞与收藏数']:
        if col in author_stats.columns:
            author_stats[col] = author_stats[col].fillna(0).astype(int)
        else:
            author_stats[col] = 0

    return author_stats

def find_note_link_column(df):
    possible_names = ['笔记链接', '链接', '分享链接', 'note_link', 'url', 'link']
    for col in df.columns:
        if col in possible_names:
            return col
    return None

# ======================== 主界面 ========================
def main():
    st.title("📕 小红书笔记数据可视化看板")
    st.markdown(f"**数据源**：GitHub 仓库（小红书笔记数据.xlsx） | 获赞与收藏数为作者主页总获赞收藏数")

    # 加载数据
    raw_df = load_data()
    if raw_df.empty:
        st.stop()

    # 预处理
    df_ana = preprocess_for_analysis(raw_df)
    link_col = find_note_link_column(raw_df)

    # 侧边栏全局筛选
    st.sidebar.header("🔍 全局数据筛选")
    st.sidebar.markdown(f"**原始笔记数**: {len(df_ana)}")
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
            st.sidebar.info(f"筛选后笔记数: {filter_mask.sum()}")

    for col_cn in ['点赞数', '评论数', '分享数', '收藏数', '获赞与收藏数']:
        if col_cn in df_ana.columns and not df_ana[col_cn].isna().all():
            min_val = float(df_ana[col_cn].min())
            max_val = float(df_ana[col_cn].max())
            if min_val < max_val:
                selected = st.sidebar.slider(f"{col_cn} 范围", min_val, max_val, (min_val, max_val))
                mask = (df_ana[col_cn] >= selected[0]) & (df_ana[col_cn] <= selected[1])
                filter_mask &= mask

    df_filtered = df_ana[filter_mask].copy()
    raw_filtered = raw_df.loc[filter_mask] if filter_mask.dtype == bool else raw_df.iloc[filter_mask]

    # ======================== 作者排行榜 ========================
    st.subheader("👥 作者维度综合排行榜")
    st.caption("以下统计基于左侧全局筛选后的数据 | 发布数量 = 有效笔记ID计数 | 获赞与收藏数为作者主页总获赞收藏数")
    author_df = get_author_aggregation(df_filtered)
    if not author_df.empty:
        display_cols_order = ['姓名', '工号', '小红书号', '用户昵称', '获赞与收藏数', '粉丝数']
        display_cols_order = [c for c in display_cols_order if c in author_df.columns]

        # 两个 Tab：🔥 获赞与收藏数 Top10，👥 粉丝数 Top10
        tab1, tab2 = st.tabs(["🔥 获赞与收藏数 Top10", "👥 粉丝数 Top10"])

        with tab1:
            if '获赞与收藏数' in author_df.columns:
                top_likes = author_df.sort_values('获赞与收藏数', ascending=False).head(10)
                top_likes_display = top_likes[display_cols_order]
                st.dataframe(top_likes_display, width='stretch')
                fig_likes = px.bar(top_likes, x='用户昵称', y='获赞与收藏数',
                                   title="获赞与收藏数 Top10（博主主页总获赞收藏）",
                                   text_auto=True, color='获赞与收藏数',
                                   hover_data=['姓名', '工号', '小红书号', '发布数量', '粉丝数'])
                st.plotly_chart(fig_likes, use_container_width=True)
            else:
                st.info("数据中不含获赞与收藏数，无法展示该排行榜。")

        with tab2:
            if '粉丝数' in author_df.columns:
                top_fans = author_df.sort_values('粉丝数', ascending=False).head(10)
                top_fans_display = top_fans[display_cols_order]
                st.dataframe(top_fans_display, width='stretch')
                fig_fans = px.bar(top_fans, x='用户昵称', y='粉丝数',
                                  title="粉丝数 Top10",
                                  text_auto=True, color='粉丝数',
                                  hover_data=['姓名', '工号', '小红书号', '发布数量', '获赞与收藏数'])
                st.plotly_chart(fig_fans, use_container_width=True)
            else:
                st.info("数据中不含粉丝数，无法展示粉丝榜。")
    else:
        st.info("未找到作者数据，无法进行作者排行榜分析。")

    # ======================== 发布监控 ========================
    st.subheader("🔍 发布监控（支持单日或日期范围）")
    st.caption("基于全量数据统计，不受左侧筛选影响。展示发布数量、发布状态等（不展示粉丝数、关注数）。")

    df_monitor = preprocess_for_analysis(raw_df.copy())  # 使用全量数据
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
                count_label = "当天发布笔记数"
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
                count_label = "范围内发布笔记数"
                title_suffix = f"（{date_desc}）"

            selected_notes = df_monitor[mask_selected].copy()

            # 所有作者基本信息
            info_cols = ['author_group_key', '姓名', '工号', '小红书号', '用户昵称']
            all_authors_info = df_monitor[info_cols].drop_duplicates('author_group_key')

            if not selected_notes.empty:
                daily_stats = selected_notes.groupby('author_group_key').agg(
                    发布数量=('笔记ID', 'count')
                ).reset_index()
                author_status = all_authors_info.merge(daily_stats, on='author_group_key', how='left')
            else:
                author_status = all_authors_info.copy()
                author_status['发布数量'] = 0

            author_status['发布数量'] = author_status['发布数量'].fillna(0).astype(int)
            author_status['发布状态'] = author_status['发布数量'].apply(lambda x: '✅ 有发布' if x > 0 else '❌ 无发布')
            author_status['备注'] = author_status['用户昵称'].apply(
                lambda x: '⚠️ 该用户小红书号存在问题，请核查是否正确' if x == '(无小红书号)' else ''
            )

            display_cols = ['姓名', '工号', '小红书号', '用户昵称', '发布数量', '发布状态', '备注']
            display_cols = [c for c in display_cols if c in author_status.columns]
            author_status = author_status[display_cols]

            filter_option = st.radio(
                "筛选作者：",
                ["全部作者", f"仅无发布作者 ({count_label}=0)", f"仅有发布作者 ({count_label}>0)"],
                horizontal=True
            )
            if filter_option == f"仅无发布作者 ({count_label}=0)":
                author_status = author_status[author_status['发布数量'] == 0]
            elif filter_option == f"仅有发布作者 ({count_label}>0)":
                author_status = author_status[author_status['发布数量'] > 0]

            st.write(f"### 作者发布统计 {title_suffix}")
            st.dataframe(author_status, width='stretch')

            total_authors = len(all_authors_info)
            published_authors = (author_status['发布数量'] > 0).sum() if not author_status.empty else 0
            total_notes = author_status['发布数量'].sum() if not author_status.empty else 0
            st.info(f"📊 **摘要**：共 {total_authors} 位作者，其中 {published_authors} 位有发布，合计发布 {total_notes} 篇笔记。")

            if not selected_notes.empty:
                with st.expander("📹 点击查看笔记详情（验证发布数量）"):
                    author_list = sorted(selected_notes['author_group_key'].unique())
                    if author_list:
                        author_display = {}
                        for key in author_list:
                            row = selected_notes[selected_notes['author_group_key'] == key].iloc[0]
                            author_display[key] = f"{row['姓名']}({row['工号']}) - {row['用户昵称']}"
                        selected_display = st.selectbox("选择作者查看其发布的笔记", options=author_list, format_func=lambda x: author_display[x], key="author_detail")
                        author_notes = selected_notes[selected_notes['author_group_key'] == selected_display].copy()
                        author_notes = author_notes.sort_values('publish_date', ascending=False)
                        st.write(f"**{author_display[selected_display]}** 在 {date_desc} 发布了 {len(author_notes)} 篇笔记：")

                        display_cols_details = []
                        if '笔记ID' in author_notes.columns:
                            display_cols_details.append('笔记ID')
                        if '笔记标题' in author_notes.columns:
                            display_cols_details.append('笔记标题')
                        if '发布日期' in author_notes.columns:
                            display_cols_details.append('发布日期')
                        if link_col and link_col in author_notes.columns:
                            display_cols_details.append(link_col)
                        if not display_cols_details:
                            display_cols_details = author_notes.columns.tolist()
                        for prefix in ['姓名', '工号', '小红书号']:
                            if prefix in author_notes.columns and prefix not in display_cols_details:
                                display_cols_details.insert(0, prefix)
                        display_cols_details = list(dict.fromkeys(display_cols_details))

                        column_config = {}
                        if link_col and link_col in author_notes.columns:
                            column_config[link_col] = st.column_config.LinkColumn("笔记链接", width="small", help="点击跳转")
                        st.dataframe(author_notes[display_cols_details], column_config=column_config, width='stretch')
            else:
                st.info("所选日期范围内没有作者发布笔记。")
        else:
            st.info("原始数据中无有效发布日期，无法进行发布监控。")
    else:
        st.info("原始数据缺少'发布日期'或'用户昵称'列，无法进行发布监控。")

    # ======================== 原始数据预览 ========================
    st.subheader("📄 原始数据预览（应用全局筛选后）")
    preview_df = raw_filtered.copy()
    if '小红书号' not in preview_df.columns:
        preview_df['小红书号'] = ''
    if '获赞与收藏数' not in preview_df.columns:
        preview_df['获赞与收藏数'] = 0
    if '关注数' not in preview_df.columns:
        preview_df['关注数'] = 0

    base_cols = ['姓名', '工号', '小红书号', '用户昵称', '笔记标题', '点赞数','收藏数', '评论数', '分享数',
                 '获赞与收藏数', '粉丝数', '发布日期']
    if link_col:
        base_cols.append(link_col)
    existing_cols = [c for c in base_cols if c in preview_df.columns]
    preview_show = preview_df[existing_cols] if existing_cols else preview_df.copy()
    column_config = {}
    if link_col and link_col in preview_show.columns:
        column_config[link_col] = st.column_config.LinkColumn("笔记链接", width="small", help="点击跳转")
    show_all = st.checkbox("显示全部数据（默认仅显示前100行）")
    if show_all:
        st.dataframe(preview_show, column_config=column_config, width='stretch')
    else:
        st.dataframe(preview_show.head(100), column_config=column_config, width='stretch')

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info(
        "**数据源**：小红书笔记数据.xlsx\n\n"
        "**字段说明**：\n"
        "- 小红书号：直接从Excel读取，独立于用户昵称。\n"
        "- 获赞与收藏数：作者主页总获赞与收藏数（取最大值，不累加）。\n"
        "- 关注数：从Excel读取，若无此列则默认为0。\n\n"
        "**姓名/工号**：从Excel中读取，若缺失显示空白。\n\n"
        "**无小红书号用户**：当「用户昵称」为空时，所有指标自动归零，且不计入发布数量。\n\n"
        "**排行榜**：两个Tab分别为「获赞与收藏数 Top10」和「粉丝数 Top10」。\n\n"
        "**发布监控**：只展示作者、小红书号、发布数量及状态，不展示粉丝数、关注数。"
    )

if __name__ == "__main__":
    main()