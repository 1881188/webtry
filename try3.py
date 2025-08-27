import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import io
import os
import platform

# 确保页面配置在所有操作前设置
st.set_page_config(
    page_title="广告指标分析系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 必要列定义
UPLOAD_REQUIRED_COLUMNS = [
    '活动', '活动第几天', '渠道', '广告系列ID_h', '广告组ID_h',
    '业绩', '订单', '花费', '曝光', '点击'
]

FINAL_REQUIRED_COLUMNS = UPLOAD_REQUIRED_COLUMNS + [
    '覆盖', 'ROI', 'ROI_数值', 'CTR', 'CTR_数值',
    'CPC', 'CPC_数值', 'CPA', 'CPA_数值', 'CPM', 'CPM_数值', 'CVR', 'CVR_数值'
]


@st.cache_data(ttl=3600)  # 增加缓存时间，减少重计算
def load_ad_data(file_path):
    """读取生成的含计算指标的文件，返回清洗后的数据和日志"""
    try:
        if not file_path.exists():
            st.error(f'❌ 文件不存在：{file_path.absolute()}')
            return None, None

        # 检查文件是否为空
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        if len(excel_file.sheet_names) == 0:
            st.error('❌ Excel文件不包含任何工作表')
            return None, None

        # 读取数据
        df = pd.read_excel(file_path, sheet_name='Sheet1', engine='openpyxl')

        # 检查数据是否为空
        if df.empty:
            st.error('❌ 数据文件为空，没有可分析的数据')
            return None, None

        # 筛选必要列
        available_cols = [col for col in FINAL_REQUIRED_COLUMNS if col in df.columns]
        if not available_cols:
            st.error('❌ 数据文件中未找到任何有效列')
            return None, None

        df = df[available_cols].copy()
        st.success(f'✅ 数据读取成功！共{len(df)}条数据，{len(df.columns)}列指标')

        # 处理重复列
        if df.columns.duplicated().any():
            duplicate_cols = df.columns[df.columns.duplicated()].tolist()
            st.warning(f'⚠️ 发现重复列：{duplicate_cols}，已保留第一列')
            df = df.loc[:, ~df.columns.duplicated()]

        # 数据清洗
        clean_log = {}

        # 处理百分比指标
        for col in ['ROI', 'CTR', 'CVR']:
            if f'{col}_数值' not in df.columns:
                if col in df.columns:
                    df[f'{col}_数值'] = df[col].astype(str).str.strip('%').apply(
                        pd.to_numeric, errors='coerce'
                    ).fillna(0)
                else:
                    st.warning(f'⚠️ 未找到{col}列，已自动创建并填充为0')
                    df[col] = '0%'
                    df[f'{col}_数值'] = 0

            clean_log[col] = f'有效数据（>0）：{len(df[df[f"{col}_数值"] > 0])}条'

        # 处理成本类指标
        for col in ['CPC', 'CPA', 'CPM']:
            if f'{col}_数值' not in df.columns:
                if col in df.columns:
                    df[f'{col}_数值'] = df[col].apply(
                        pd.to_numeric, errors='coerce'
                    ).replace([np.inf, -np.inf], 99999.99).fillna(99999.99)
                else:
                    st.warning(f'⚠️ 未找到{col}列，已自动创建并填充为0')
                    df[col] = 0
                    df[f'{col}_数值'] = 0

            clean_log[col] = f'有效数据（<99999）：{len(df[df[f"{col}_数值"] < 99999])}条'

        # 处理基础指标
        for col in ['曝光', '点击', '订单', '覆盖', '业绩', '花费']:
            if col not in df.columns:
                st.warning(f'⚠️ 未找到{col}列，已自动创建并填充为0')
                df[col] = 0

            df[col] = df[col].apply(pd.to_numeric, errors='coerce').fillna(0)
            clean_log[col] = f'有效数据（>0）：{len(df[df[col] > 0])}条'

        return df, clean_log

    except Exception as e:
        st.error(f'❌ 数据处理错误：{str(e)}')
        return None, None


def calculate_ad_indicators(raw_df):
    """计算广告指标并处理可能的空值问题"""
    if raw_df.empty:
        st.error('❌ 原始数据为空，无法计算指标')
        return None

    calculated_df = raw_df.copy()

    # 确保必要数值列存在
    numeric_cols = ['业绩', '订单', '花费', '曝光', '点击']
    for col in numeric_cols:
        if col not in calculated_df.columns:
            st.warning(f'⚠️ 原始数据缺少{col}列，已填充为0')
            calculated_df[col] = 0
        calculated_df[col] = pd.to_numeric(calculated_df[col], errors='coerce').fillna(0)

    # 安全除法函数
    def safe_divide(numerator, denominator, default=0):
        return np.where(denominator == 0, default, numerator / denominator)

    # 计算各指标
    try:
        # 覆盖（使用曝光替代）
        calculated_df['覆盖'] = calculated_df['曝光']

        # ROI（投资回报率）
        calculated_df['ROI_数值'] = safe_divide(
            (calculated_df['业绩'] - calculated_df['花费']),
            calculated_df['花费'].replace(0, 1),
            default=0
        ) * 100
        calculated_df['ROI'] = calculated_df['ROI_数值'].round(2).astype(str) + '%'

        # CTR（点击通过率）
        calculated_df['CTR_数值'] = safe_divide(
            calculated_df['点击'],
            calculated_df['曝光'].replace(0, 1),
            default=0
        ) * 100
        calculated_df['CTR'] = calculated_df['CTR_数值'].round(2).astype(str) + '%'

        # CPC（每次点击成本）
        calculated_df['CPC_数值'] = safe_divide(
            calculated_df['花费'],
            calculated_df['点击'].replace(0, 1),
            default=0
        )
        calculated_df['CPC'] = calculated_df['CPC_数值'].round(2)

        # CPA（每次订单成本）
        calculated_df['CPA_数值'] = safe_divide(
            calculated_df['花费'],
            calculated_df['订单'].replace(0, 1),
            default=0
        )
        calculated_df['CPA'] = calculated_df['CPA_数值'].round(2)

        # CPM（每千次曝光成本）
        calculated_df['CPM_数值'] = safe_divide(
            calculated_df['花费'],
            calculated_df['曝光'].replace(0, 1),
            default=0
        ) * 1000
        calculated_df['CPM'] = calculated_df['CPM_数值'].round(2)

        # CVR（订单转化率）
        calculated_df['CVR_数值'] = safe_divide(
            calculated_df['订单'],
            calculated_df['点击'].replace(0, 1),
            default=0
        ) * 100
        calculated_df['CVR'] = calculated_df['CVR_数值'].round(2).astype(str) + '%'

        return calculated_df

    except Exception as e:
        st.error(f'❌ 指标计算失败：{str(e)}')
        return None


def calculate_single_ranking(filtered_df):
    indicator_config = {
        'ROI_数值': (False, 'ROI'),
        '覆盖': (False, '覆盖'),
        '点击': (False, '点击'),
        '订单': (False, '订单'),
        'CTR_数值': (False, 'CTR'),
        'CPC_数值': (True, 'CPC'),
        'CPA_数值': (True, 'CPA'),
        'CPM_数值': (True, 'CPM'),
        'CVR_数值': (False, 'CVR')
    }

    # 确保必要列存在
    for col, (_, name) in indicator_config.items():
        if col not in filtered_df.columns:
            st.warning(f'⚠️ 缺少{name}计算所需数据，已自动填充为0')
            filtered_df[col] = 0

    for col, (ascending, name) in indicator_config.items():
        if isinstance(filtered_df[col], pd.DataFrame):
            filtered_df[col] = filtered_df[col].iloc[:, 0]

        filtered_df[col] = filtered_df[col].replace([np.inf, -np.inf], 99999.99 if ascending else 0)
        filtered_df[col] = filtered_df[col].fillna(99999.99 if ascending else 0)

        rank_series = filtered_df[col].rank(method='min', ascending=ascending)
        if isinstance(rank_series, pd.DataFrame):
            rank_series = rank_series.iloc[:, 0]

        filtered_df[f'{name}_排名'] = rank_series.fillna(0).astype('Int64')

    return filtered_df


def get_desktop_path():
    """获取当前用户的桌面路径（适配Windows/macOS/Linux）"""
    try:
        if platform.system() == 'Windows':
            return Path(os.path.expanduser("~")) / "Desktop"
        elif platform.system() == 'Darwin':  # macOS
            return Path(os.path.expanduser("~")) / "Desktop"
        else:  # Linux
            return Path(os.path.expanduser("~")) / "Desktop"
    except Exception as e:
        st.warning(f'⚠️ 获取桌面路径失败：{str(e)}，将使用当前工作目录')
        return Path.cwd()


def main():
    st.title('广告指标计算与综合排名系统')
    st.divider()

    # 使用session_state保存状态，减少不必要的重渲染
    if 'stage' not in st.session_state:
        st.session_state.stage = 0

    # 步骤1：文件上传
    st.subheader('📤 步骤1：上传原始广告数据')
    uploaded_file = st.file_uploader(
        label='请上传Excel文件（需包含以下列）',
        type=['xlsx', 'xls'],
        help=f'必填列：{", ".join(UPLOAD_REQUIRED_COLUMNS)}'
    )

    if not uploaded_file:
        st.info('ℹ️ 请先上传符合要求的Excel文件，再进行后续操作')
        st.session_state.stage = 0
        return
    else:
        st.session_state.stage = 1

    # 步骤2：验证原始数据列
    try:
        # 显式指定引擎，避免依赖问题
        raw_df = pd.read_excel(uploaded_file, sheet_name='Sheet1', engine='openpyxl')

        if raw_df.empty:
            st.error('❌ 上传的文件为空，没有可处理的数据')
            return

        st.success('✅ 原始文件读取成功！')
    except Exception as e:
        st.error(f'❌ 读取文件失败：{str(e)}（请确保文件格式为Excel）')
        return

    # 检查必要列
    missing_cols = [col for col in UPLOAD_REQUIRED_COLUMNS if col not in raw_df.columns]
    if missing_cols:
        st.error(f'❌ 原始文件缺少必要列：{", ".join(missing_cols)}')
        st.info(f'✅ 正确列名清单：{", ".join(UPLOAD_REQUIRED_COLUMNS)}')
        return

    # 步骤3：计算广告指标
    st.subheader('📊 步骤2：自动计算广告指标')
    with st.spinner('正在计算ROI、CTR、CPC等指标...'):
        calculated_df = calculate_ad_indicators(raw_df)

    if calculated_df is None or calculated_df.empty:
        st.error('❌ 无法生成有效指标数据，排名功能无法使用')
        return

    st.success('✅ 所有指标计算完成！')
    st.session_state.stage = 2

    # 预览计算结果 - 限制显示行数，避免大量数据导致渲染问题
    st.subheader('📄 计算结果预览')
    preview_cols = ['活动', '活动第几天', '渠道', '业绩', '订单', 'ROI', 'CTR', 'CPC', 'CPA']
    preview_cols = [col for col in preview_cols if col in calculated_df.columns]

    # 限制预览数据量，解决渲染问题
    max_preview_rows = 100
    if len(calculated_df) > max_preview_rows:
        st.info(f'⚠️ 数据量较大，仅显示前{max_preview_rows}行预览')
        st.dataframe(
            calculated_df[preview_cols].head(max_preview_rows),
            use_container_width=True,
            height=300  # 固定高度，避免动态变化
        )
    else:
        st.dataframe(
            calculated_df[preview_cols],
            use_container_width=True,
            height=300
        )

    # 步骤4：保存新文件和下载功能
    st.subheader('💾 步骤3：生成含指标的新文件')
    output_filename = '广告数据_含计算指标.xlsx'
    local_save_success = False
    local_file_path = None

    # 尝试保存到桌面目录
    try:
        desktop_path = get_desktop_path()
        if not desktop_path.exists():
            desktop_path.mkdir(parents=True, exist_ok=True)
        local_file_path = desktop_path / output_filename

        calculated_df.to_excel(local_file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
        local_save_success = True
        st.success(f'✅ 文件已保存到桌面！路径：{local_file_path.absolute()}')
        st.info('💡 提示：可直接在桌面找到文件，或复制上方路径到文件管理器打开')
    except PermissionError:
        st.warning('⚠️ 桌面目录无写入权限，将尝试保存到文档目录')
        try:
            docs_path = Path(os.path.expanduser("~")) / "Documents"
            docs_path.mkdir(parents=True, exist_ok=True)
            local_file_path = docs_path / output_filename
            calculated_df.to_excel(local_file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            local_save_success = True
            st.success(f'✅ 文件已保存到文档目录！路径：{local_file_path.absolute()}')
        except Exception as e:
            st.error(f'❌ 本地保存失败：{str(e)}（权限不足）')
    except Exception as e:
        st.error(f'❌ 本地保存失败：{str(e)}')

    # 提供文件下载按钮
    buffer = io.BytesIO()
    calculated_df.to_excel(buffer, index=False, sheet_name='Sheet1', engine='openpyxl')
    buffer.seek(0)

    st.download_button(
        label='📥 点击下载含指标的Excel文件（保底方案）',
        data=buffer,
        file_name=output_filename,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        help='若本地保存失败，点击此按钮手动下载文件'
    )

    # 权限问题解决提示
    if not local_save_success:
        st.info('🔧 权限问题解决方法：')
        st.write('1. 右键点击生成的Excel文件下载按钮，选择"另存为"')
        st.write('2. 手动选择一个有权限的文件夹（如桌面、文档）')
        st.write('3. 保存后即可使用完整的排名功能')

    # 步骤5：加载数据执行排名功能
    if local_save_success and local_file_path and local_file_path.exists():
        st.subheader('🏆 步骤4：广告指标综合排名')
        st.divider()

        # 使用容器隔离排名部分，减少渲染冲突
        with st.container():
            df, clean_log = load_ad_data(local_file_path)
            if df is None or df.empty:
                st.error('❌ 无法加载新生成的文件，排名功能无法使用')
                return

            # 数据概况
            st.info('📈 数据概况：')
            st.write(f'- 总数据量：{len(df)} 条')
            st.write(f'- 活动类型：{df["活动"].nunique()} 种（{", ".join(sorted(df["活动"].unique())[:3])}...）')
            st.write(f'- 天数范围：{df["活动第几天"].min()} - {df["活动第几天"].max()} 天')
            for ind, log in clean_log.items():
                st.write(f'- {ind}：{log}')
            st.divider()

            # 筛选条件（侧边栏）
            st.sidebar.header('1. 筛选条件')
            selected_activity = st.sidebar.selectbox(
                '选择活动',
                options=['全部活动'] + sorted(df['活动'].unique()),
                index=0,
                key='activity_select'  # 增加唯一key，避免组件冲突
            )

            min_day, max_day = df['活动第几天'].min(), df['活动第几天'].max()
            selected_day = st.sidebar.number_input(
                f'选择活动天数（{min_day}-{max_day}）',
                min_value=min_day, max_value=max_day, value=min_day, step=1,
                key='day_input'  # 增加唯一key
            )

            # 执行筛选
            filtered_df = df[df['活动第几天'] == selected_day].copy()
            if selected_activity != '全部活动':
                filtered_df = filtered_df[filtered_df['活动'] == selected_activity].copy()
            filtered_df = filtered_df.reset_index(drop=True)

            if filtered_df.empty:
                st.warning(f'⚠️ 未找到「{selected_activity} - 第{selected_day}天」的数据，请更换筛选条件')
                return
            st.success(f'✅ 筛选到 {len(filtered_df)} 条数据，可设置权重计算排名')

            # 权重设置（侧边栏）
            st.sidebar.header('2. 指标权重（总和需100%）')
            weights = {}
            default_weight = round(100 / 9, 2)

            st.sidebar.subheader('📈 收益/效率类')
            weights['ROI'] = st.sidebar.number_input('ROI 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_roi')
            weights['覆盖'] = st.sidebar.number_input('覆盖 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_cover')
            weights['点击'] = st.sidebar.number_input('点击 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_click')
            weights['订单'] = st.sidebar.number_input('订单 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_order')
            weights['CTR'] = st.sidebar.number_input('CTR 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_ctr')
            weights['CVR'] = st.sidebar.number_input('CVR 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_cvr')

            st.sidebar.subheader('💰 成本类')
            weights['CPC'] = st.sidebar.number_input('CPC 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_cpc')
            weights['CPA'] = st.sidebar.number_input('CPA 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_cpa')
            weights['CPM'] = st.sidebar.number_input('CPM 权重(%)', 0.0, 100.0, default_weight, 0.1, key='w_cpm')

            # 权重总和校验
            total_weight = round(sum(weights.values()), 1)
            st.sidebar.info(f'当前总权重：{total_weight}%')
            if total_weight != 100.0:
                st.warning(f'⚠️ 请调整权重至100%（当前{total_weight}%），否则无法计算排名')
                return

            # 综合排名计算
            with st.spinner('正在计算综合排名...'):
                ranked_df = calculate_single_ranking(filtered_df)
                total_count = len(ranked_df)

                ranked_df['综合得分'] = 0.0
                for ind, w in weights.items():
                    rank_col = f'{ind}_排名'
                    ranked_df['综合得分'] += (total_count - ranked_df[rank_col] + 1) * (w / 100)

                ranked_df['综合排名'] = ranked_df['综合得分'].rank(method='min', ascending=False).astype('Int64')
                final_df = ranked_df.sort_values('综合排名', ascending=True).reset_index(drop=True)

            # 排名结果展示 - 限制显示行数
            st.subheader(f'📊 综合排名结果：{selected_activity} - 第{selected_day}天（共{total_count}条）')
            st.divider()

            display_cols = [
                '综合排名', '综合得分', '活动', '活动第几天', '渠道', '广告系列ID_h', '广告组ID_h',
                '业绩', '订单', '花费', '曝光', '点击', '覆盖', 'ROI', 'CTR', 'CPC', 'CPA', 'CPM', 'CVR',
                'ROI_排名', '覆盖_排名', '点击_排名', '订单_排名', 'CTR_排名', 'CVR_排名',
                'CPC_排名', 'CPA_排名', 'CPM_排名'
            ]
            show_cols = [col for col in display_cols if col in final_df.columns]
            show_df = final_df[show_cols].copy()

            # 格式化数值显示
            num_cols = ['业绩', '花费', 'CPC', 'CPA', 'CPM', '综合得分', '曝光', '点击', '订单', '覆盖']
            for col in num_cols:
                if col in show_df.columns:
                    show_df[col] = show_df[col].round(2)

            # 限制显示行数，防止渲染错误
            max_display_rows = 200
            if len(show_df) > max_display_rows:
                st.info(f'⚠️ 数据量较大，仅显示前{max_display_rows}行结果')
                display_df = show_df.head(max_display_rows)
            else:
                display_df = show_df

            st.dataframe(
                display_df,
                column_config={
                    '综合排名': st.column_config.NumberColumn('综合排名', width='small'),
                    '综合得分': st.column_config.NumberColumn('综合得分', width='small'),
                    **{f'{ind}_排名': st.column_config.NumberColumn(f'{ind}排名', width='small') for ind in
                       weights.keys()}
                },
                use_container_width=True,
                height=400  # 固定高度，避免动态变化
            )

            # 展示Top3
            st.divider()
            st.write('🏆 综合排名Top3：')
            top3_cols = ['综合排名', '广告系列ID_h', '广告组ID_h', 'ROI', 'CPA', 'CPC', '综合得分']
            top3 = final_df[final_df['综合排名'] <= 3][[col for col in top3_cols if col in final_df.columns]]

            for _, row in top3.iterrows():
                st.write(f"""
                **第{row['综合排名']}名** | 系列ID：{row['广告系列ID_h']} | 组ID：{row['广告组ID_h']}
                - ROI：{row['ROI']} | CPA：{row['CPA']:.2f} | CPC：{row['CPC']:.2f}
                - 综合得分：{row['综合得分']:.2f}
                """)
                st.divider()
    else:
        st.info('ℹ️ 排名功能使用说明：')
        st.write('1. 点击上方"下载文件"按钮，将文件保存到本地（如桌面）')
        st.write('2. 确保保存路径无中文特殊字符，且有写入权限')
        st.write('3. 重新运行程序，即可自动加载文件并使用排名功能')


if __name__ == '__main__':
    main()
