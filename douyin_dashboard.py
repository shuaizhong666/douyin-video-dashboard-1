"""
抖音视频数据可视化看板（GitHub Raw 版）
数据源：GitHub 仓库中的 Excel 文件（通过 Raw URL 实时读取）
运行命令：streamlit run douyin_dashboard.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO
import warnings
from datetime import date, datetime  # ✅ 修复 datetime 未定义
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

# ======================== 增强的加载函数 ========================
@st.cache_data(ttl=3600, show_spinner="正在从 GitHub 加载数据...")
def load_data_from_github(url, max_retries=3):
    """
    从 GitHub Raw URL 加载 Excel，带重试、诊断和备用读取方案。
    """
    for attempt in range(max_retries):
        try:
            # 添加随机参数避免 CDN 缓存
            random_param = f"?_={int(datetime.now().timestamp())}" if attempt > 0 else ""
            final_url = url + random_param
            response = requests.get(final_url, timeout=30)
            response.raise_for_status()

            # 检查 Content-Length（如果有）
            content_length = response.headers.get('Content-Length')
            if content_length and int(content_length) < 1000:
                st.warning(f"⚠️ 文件很小（{content_length} 字节），可能不是完整的 Excel 文件。")

            content = response.content

            # 方法1：直接使用 pandas + openpyxl
            try:
                df = pd.read_excel(BytesIO(content), engine='openpyxl')
                if df.empty:
                    st.warning("Excel 文件读取成功，但数据为空。")
                else:
                    st.success(f"✅ 数据加载成功，共 {len(df)} 行记录")
                return df
            except Exception as e1:
                st.warning(f"pandas 直接读取失败: {e1}")

                # 方法2：写入临时文件，用 openpyxl 直接加载
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                        tmp.write(content)
                        tmp_path = tmp.name
                    from openpyxl import load_workbook
                    wb = load_workbook(tmp_path, read_only=True)
                    sheet = wb.active
                    # 将数据转为 DataFrame
                    data = sheet.values
                    cols = next(data)
                    df = pd.DataFrame(data, columns=cols)
                    wb.close()
                    os.unlink(tmp_path)
                    st.success("✅ 通过 openpyxl 直接加载成功")
                    return df
                except Exception as e2:
                    st.error(f"临时文件读取也失败: {e2}")
                    # 诊断：检查文件头部是否为 PK 且是有效 zip
                    if content[:4] == b'PK\x03\x04':
                        st.error("文件头正确（PK），但 openpyxl 无法解析。可能是文件损坏，请确认本地 Excel 能否正常打开。")
                        # 显示前 200 字节的十六进制
                        hex_preview = ' '.join(f'{b:02x}' for b in content[:32])
                        st.text(f"文件头 hex: {hex_preview}")
                    else:
                        st.error(f"文件头不是 zip 文件。前20字节: {content[:20]}")
                    raise e2

        except requests.exceptions.RequestException as e:
            st.error(f"网络请求失败 (尝试 {attempt+1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return pd.DataFrame()
        except Exception as e:
            st.error(f"读取 Excel 失败: {e}")
            if attempt == max_retries - 1:
                return pd.DataFrame()
    return pd.DataFrame()

def preprocess_for_analysis(df):
    """预处理数据：转换数值列和日期列"""
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
    """作者维度聚合"""
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

# ======================== 主界面 ========================
def main():
    st.title("🎵 抖音视频数据可视化看板")
    st.markdown(f"**数据源**：GitHub 仓库（实时同步）")

    raw_df = load_data_from_github(GITHUB_RAW_URL)
    if raw_df.empty:
        st.stop()

    df_ana = preprocess_for_analysis(raw_df)

    # 侧边栏筛选
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

    # 单日发布监控
    st.subheader("🔍 单日发布监控（支持筛选未发布作者）")
    if '创建时间' in raw_df.columns and '作者昵称' in raw_df.columns:
        raw_df_date = raw_df.copy()
        raw_df_date['publish_date'] = pd.to_datetime(raw_df_date['创建时间'], errors='coerce')
        valid_pub = raw_df_date['publish_date'].dropna()
        if not valid_pub.empty:
            min_date_all = valid_pub.min().date()
            max_date_all = valid_pub.max().date()
            default_date = max_date_all
            selected_date = st.date_input("📅 选择检查日期", value=default_date,
                                          min_value=min_date_all, max_value=None)
            mask_today = (raw_df_date['publish_date'].dt.date == selected_date)
            daily_stats = raw_df_date[mask_today].groupby('作者昵称').size().reset_index(name='当天发布数')
            all_authors = sorted(raw_df_date['作者昵称'].dropna().unique())
            author_status = pd.DataFrame({'作者昵称': all_authors})
            author_status = author_status.merge(daily_stats, on='作者昵称', how='left')
            author_status['当天发布数'] = author_status['当天发布数'].fillna(0).astype(int)
            author_status['发布状态'] = author_status['当天发布数'].apply(lambda x: '✅ 已发布' if x > 0 else '❌ 未发布')
            filter_option = st.radio("筛选作者：",
                ["全部作者", "仅未发布作者 (当天发布数为0)", "仅已发布作者 (当天发布数>0)"],
                horizontal=True)
            if filter_option == "仅未发布作者 (当天发布数为0)":
                author_status = author_status[author_status['当天发布数'] == 0]
            elif filter_option == "仅已发布作者 (当天发布数>0)":
                author_status = author_status[author_status['当天发布数'] > 0]
            st.write(f"### {selected_date} 作者发布情况 (共 {len(author_status)} 位)")
            st.dataframe(author_status, use_container_width=True)
            total_authors = len(all_authors)
            published_authors = len(daily_stats)
            total_videos_today = daily_stats['当天发布数'].sum()
            st.info(f"📊 **摘要**：共 {total_authors} 位作者，其中 {published_authors} 位当天发布了视频，合计发布 {total_videos_today} 个视频。")
        else:
            st.info("原始数据中无有效发布日期，无法进行单日监控。")
    else:
        st.info("原始数据缺少'创建时间'或'作者昵称'列，无法进行单日监控。")

    # 原始数据预览
    st.subheader("📄 原始数据预览（筛选后）")
    display_cols_cn = ['作者昵称', '视频描述', '点赞数', '评论数', '分享数', '收藏数', '粉丝数', '创建时间']
    existing_cn = [c for c in display_cols_cn if c in raw_filtered.columns]
    preview_df = raw_filtered[existing_cn] if existing_cn else raw_filtered
    show_all = st.checkbox("显示全部数据（默认仅显示前100行）")
    if show_all:
        st.dataframe(preview_df, use_container_width=True)
    else:
        st.dataframe(preview_df.head(100), use_container_width=True)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 说明")
    st.sidebar.info("看板支持单日发布监控、作者双排行榜。数据源为 GitHub Raw URL，每次刷新页面最多延迟 1 小时缓存。")

if __name__ == "__main__":
    main()