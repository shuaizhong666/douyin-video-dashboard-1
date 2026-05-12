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
            st.error(f"加载失败 (尝试 {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return pd.DataFrame()
    return pd.DataFrame()


def preprocess_for_analysis(df):
    if df.empty:
        return df
    df_work = df.copy()

    # 补全必要列
    for col in ['姓名', '工号', '抖音号']:
        if col not in df_work.columns:
            df_work[col] = ''
            st.info(f"⚠️ 原始数据缺少「{col}」列，已自动补全")

    # 数值列处理
    num_cols = ['点赞数', '评论数', '分享数', '收藏数', '粉丝数', '获赞总数']
    for col in num_cols:
        if col in df_work.columns:
            df_work[col] = pd.to_numeric(df_work[col], errors='coerce').fillna(0)

    # 昵称有效性判断
    if '作者昵称' in df_work.columns:
        df_work['作者昵称'] = df_work['作者昵称'].astype(str).str.strip()
        df_work['has_valid_nickname'] = (df_work['作者昵称'] != '') & (df_work['作者昵称'] != 'nan')
    else:
        df_work['作者昵称'] = ''
        df_work['has_valid_nickname'] = False

    # 无效昵称归零处理
    invalid_mask = ~df_work['has_valid_nickname']
    if invalid_mask.any():
        for col in num_cols:
            if col in df_work.columns:
                df_work.loc[invalid_mask, col] = 0
        if '视频ID' in df_work.columns:
            df_work.loc[invalid_mask, '视频ID'] = np.nan
        df_work.loc[invalid_mask, '作者昵称'] = '(无抖音号)'

    # 分组键（避免错误合并）
    name_series = df_work['姓名'].fillna('').astype(str)
    id_series = df_work['工号'].fillna('').astype(str)
    douyin_series = df_work['抖音号'].fillna('').astype(str)

    valid_mask = df_work['has_valid_nickname']
    df_work.loc[valid_mask, 'author_group_key'] = df_work.loc[valid_mask, '作者昵称']

    if invalid_mask.any():
        combined_key = name_series + '_' + id_series + '_' + douyin_series
        all_empty_mask = (name_series == '') & (id_series == '') & (douyin_series == '') & invalid_mask
        combined_key = combined_key.where(~all_empty_mask, combined_key + '_' + df_work.index.astype(str))
        df_work.loc[invalid_mask, 'author_group_key'] = combined_key.loc[invalid_mask]

    # 日期解析
    if '创建时间' in df_work.columns:
        df_work['publish_date'] = pd.to_datetime(df_work['创建时间'], errors='coerce')
    else:
        df_work['publish_date'] = pd.NaT

    return df_work


def get_author_aggregation(df):
    if df.empty or 'author_group_key' not in df.columns:
        return pd.DataFrame()

    agg_dict = {
        '作者昵称': 'first', '姓名': 'first', '工号': 'first', '抖音号': 'first', '视频ID': 'count'
    }
    for col in ['点赞数', '评论数', '收藏数', '粉丝数']:
        if col in df.columns:
            agg_dict[col] = 'sum' if col != '粉丝数' else 'max'

    author_stats = df.groupby('author_group_key').agg(agg_dict).reset_index()
    rename_map = {'author_group_key': '分组键', '视频ID': '发布数量', '点赞数': '总点赞数', '评论数': '总评论数',
                  '收藏数': '总收藏数'}
    author_stats.rename(columns=rename_map, inplace=True)

    for col in ['发布数量', '总点赞数', '总评论数', '总收藏数', '粉丝数']:
        author_stats[col] = author_stats.get(col, 0).fillna(0).astype(int)

    return author_stats


def find_video_link_column(df):
    possible = ['视频链接', '链接', '分享链接', 'video_link', 'url', 'link']
    for col in possible:
        if col in df.columns:
            return col
    return None


def reorder_columns_with_douyin(df):
    if df.empty or '抖音号' not in df.columns or '工号' not in df.columns:
        return df
    cols = df.columns.tolist()
    cols.remove('抖音号')
    cols.insert(cols.index('工号') + 1, '抖音号')
    return df[cols]


# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown("**数据源**：GitHub 仓库 | 支持姓名/工号/抖音号 | 无昵称自动归零")

    raw_df = load_data_from_github(GITHUB_RAW_URL)
    if raw_df.empty:
        st.stop()

    df_ana = preprocess_for_analysis(raw_df)
    link_col = find_video_link_column(raw_df)

    # 侧边栏筛选
    st.sidebar.header("🔍 全局筛选")
    st.sidebar.markdown(f"**总视频数**: {len(df_ana)}")
    filter_mask = pd.Series([True] * len(df_ana))

    # 修复日期解包报错（核心修复）
    if 'publish_date' in df_ana.columns:
        valid_dates = df_ana['publish_date'].dropna()
        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()
            date_range = st.sidebar.date_input("选择日期范围", value=(min_date, max_date), min_value=min_date)

            if len(date_range) == 1:
                start_date = end_date = date_range[0]
            else:
                start_date, end_date = date_range

            mask = (df_ana['publish_date'].dt.date >= start_date) & (df_ana['publish_date'].dt.date <= end_date)
            filter_mask &= mask
            st.sidebar.info(f"筛选后: {filter_mask.sum()} 条")

    # 数值筛选
    for col in ['点赞数', '评论数', '分享数', '收藏数']:
        if col in df_ana.columns and not df_ana[col].isna().all():
            min_v, max_v = float(df_ana[col].min()), float(df_ana[col].max())
            if min_v < max_v:
                slc = st.sidebar.slider(f"{col} 区间", min_v, max_v, (min_v, max_v))
                filter_mask &= (df_ana[col] >= slc[0]) & (df_ana[col] <= slc[1])

    df_filtered = df_ana[filter_mask].copy()
    raw_filtered = raw_df.loc[filter_mask]

    # 作者排行榜
    st.subheader("👥 作者综合排行榜")
    author_df = get_author_aggregation(df_filtered)
    if not author_df.empty:
        base = ['姓名', '工号', '抖音号', '作者昵称', '发布数量', '总点赞数', '总评论数', '总收藏数', '粉丝数']
        display_cols = [c for c in base if c in author_df.columns]
        tab1, tab2, tab3 = st.tabs(["📦 发布量 Top10", "❤️ 总点赞 Top10", "👥 粉丝数 Top10"])

        with tab1:
            top = author_df.sort_values('发布数量', ascending=False).head(10)
            st.dataframe(reorder_columns_with_douyin(top[display_cols]), width='stretch')
            st.plotly_chart(
                px.bar(top, x='作者昵称', y='发布数量', title='发布量 Top10', text_auto=True, color='发布数量'),
                width='stretch')

        with tab2:
            if '总点赞数' in author_df.columns:
                top = author_df.sort_values('总点赞数', ascending=False).head(10)
                st.dataframe(reorder_columns_with_douyin(top[display_cols]), width='stretch')
                st.plotly_chart(
                    px.bar(top, x='作者昵称', y='总点赞数', title='总点赞 Top10', text_auto=True, color='总点赞数'),
                    width='stretch')

        with tab3:
            if '粉丝数' in author_df.columns:
                top = author_df.sort_values('粉丝数', ascending=False).head(10)
                st.dataframe(reorder_columns_with_douyin(top[display_cols]), width='stretch')
                st.plotly_chart(
                    px.bar(top, x='作者昵称', y='粉丝数', title='粉丝数 Top10', text_auto=True, color='粉丝数'),
                    width='stretch')

    # 发布监控
    st.subheader("🔍 发布监控")
    df_monitor = preprocess_for_analysis(raw_df.copy())
    if 'publish_date' in df_monitor.columns and 'author_group_key' in df_monitor.columns:
        valid_days = df_monitor['publish_date'].dropna()
        if not valid_days.empty:
            min_all, max_all = valid_days.min().date(), valid_days.max().date()
            mode = st.radio("日期模式", ["单天", "范围"], horizontal=True)

            if mode == "单天":
                sel_date = st.date_input("选择日期", value=min(max_all, date.today() - timedelta(days=1)),
                                         min_value=min_all)
                mask_day = df_monitor['publish_date'].dt.date == sel_date
                label = "当日发布数"
            else:
                rng = st.date_input("日期范围", value=(min_all, max_all), min_value=min_all)
                start, end = rng if len(rng) == 2 else (rng[0], rng[0])
                mask_day = (df_monitor['publish_date'].dt.date >= start) & (df_monitor['publish_date'].dt.date <= end)
                label = "区间发布数"

            selected = df_monitor[mask_day]
            info_cols = ['author_group_key', '姓名', '工号', '抖音号', '作者昵称']
            all_authors = df_monitor[info_cols].drop_duplicates('author_group_key')

            if not selected.empty:
                stats = selected.groupby('author_group_key')['视频ID'].count().reset_index(name='发布数量')
                res = all_authors.merge(stats, on='author_group_key', how='left')
            else:
                res = all_authors.assign(发布数量=0)

            res[label] = res['发布数量'].fillna(0).astype(int)
            res['发布状态'] = res[label].apply(lambda x: '✅ 已发布' if x > 0 else '❌ 未发布')
            res['备注'] = res['作者昵称'].apply(lambda x: '⚠️ 无抖音号' if x == '(无抖音号)' else '')

            show_cols = ['姓名', '工号', '抖音号', '作者昵称', label, '发布状态', '备注']
            res = res[[c for c in show_cols if c in res.columns]]

            filt = st.radio("筛选", ["全部", "仅未发布", "仅已发布"], horizontal=True)
            if filt == "仅未发布": res = res[res[label] == 0]
            if filt == "仅已发布": res = res[res[label] > 0]

            st.dataframe(res, width='stretch')

    # 原始数据
    st.subheader("📄 原始数据预览")
    base_cols = ['姓名', '工号', '抖音号', '作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数',
                 '创建时间']
    if link_col: base_cols.append(link_col)
    preview = raw_filtered[[c for c in base_cols if c in raw_filtered.columns]]
    preview = reorder_columns_with_douyin(preview)

    col_cfg = {link_col: st.column_config.LinkColumn("视频链接")} if link_col else {}
    if st.checkbox("显示全部数据"):
        st.dataframe(preview, column_config=col_cfg, width='stretch')
    else:
        st.dataframe(preview.head(100), column_config=col_cfg, width='stretch')


if __name__ == "__main__":
    main()