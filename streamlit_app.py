import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import io
import os
import requests

# --- 1. 字体配置 (解决Streamlit Cloud中文乱码问题) ---
@st.cache_resource
def configure_font():
    """
    下载并配置中文字体（SimHei 或 Noto Sans SC）。
    """
    font_url = "https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC-Regular.ttf"
    font_path = "NotoSansSC-Regular.ttf"

    if not os.path.exists(font_path):
        with st.spinner("正在下载中文字体，请稍候..."):
            try:
                response = requests.get(font_url)
                with open(font_path, "wb") as f:
                    f.write(response.content)
            except Exception as e:
                st.error(f"字体下载失败: {e}")
                return None

    # 添加字体到 Matplotlib
    fm.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = 'Noto Sans SC' # 设置全局字体
    plt.rcParams['axes.unicode_minus'] = False   # 解决负号显示问题
    return font_path

# --- 2. 数据处理函数 ---
def process_data(uploaded_file):
    """
    读取并清洗数据：处理多层表头，填充合并单元格
    """
    # 1. 读取所有数据，不做表头解析
    df_raw = pd.read_csv(uploaded_file, header=None, dtype=str)
    
    # 2. 提取表头行（根据您的描述，第3行是指标，第4行是分子分母）
    # Python索引从0开始，所以是 index 2 和 3
    # 注意：CSV如果前两行被忽略，通常pandas读进来时前两行可能已经是数据了
    # 这里我们假设用户上传的文件包含那两行被忽略的行
    
    # 获取指标名称行 (第3行)
    metric_names = df_raw.iloc[2].fillna(method='ffill') # 向前填充指标名称
    
    # 获取子列名行 (第4行)
    sub_cols = df_raw.iloc[3]
    
    # 3. 构建新的列名
    # 组合两行表头，例如: "DCC首呼_指标"
    new_columns = []
    for m, s in zip(metric_names, sub_cols):
        m = str(m).strip()
        s = str(s).strip()
        if m == "nan" or m == "":
            new_columns.append(s) # 如果第一行是空的（如代理商列），只取第二行
        elif s == "nan" or s == "":
            new_columns.append(m)
        else:
            new_columns.append(f"{m}\n{s}") # 使用换行符分隔，方便绘图

    # 4. 处理数据体 (第5行及之后)
    df_data = df_raw.iloc[4:].copy()
    df_data.columns = new_columns
    
    # 重命名固定列，防止乱码或不一致
    # 假设第一列是代理商，第二列是管家
    cols = list(df_data.columns)
    cols[0] = "代理商"
    cols[1] = "管家"
    df_data.columns = cols
    
    # 5. 填充“代理商”列（处理合并单元格）
    df_data['代理商'] = df_data['代理商'].fillna(method='ffill')
    
    # 6. 过滤掉完全为空的行
    df_data = df_data.dropna(how='all')
    
    return df_data

# --- 3. 图片生成函数 ---
def generate_long_image(agent_name, agent_data):
    """
    使用 Matplotlib 绘制表格长图
    """
    # 配置字体
    configure_font()
    
    # 准备绘图数据
    # 只需要展示的列：管家 + 所有指标列（排除代理商列）
    plot_df = agent_data.drop(columns=['代理商'])
    
    # 计算图片尺寸
    # 高度 = (行数 * 0.5) + 表头高度
    # 宽度 = 列数 * 1.2
    num_rows, num_cols = plot_df.shape
    fig_width = max(10, num_cols * 1.5)
    fig_height = max(4, num_rows * 0.8 + 2)
    
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    
    # 隐藏坐标轴
    ax.axis('off')
    ax.axis('tight')
    
    # 绘制表格
    table = ax.table(
        cellText=plot_df.values,
        colLabels=plot_df.columns,
        cellLoc='center',
        loc='center',
        bbox=[0, 0, 1, 1] # 表格占满整个图
    )
    
    # 美化表格
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    
    # 遍历表格单元格进行样式设置
    cells = table.get_celld()
    for (row, col), cell in cells.items():
        cell.set_text_props(padding=10)
        
        # 表头样式 (row == 0)
        if row == 0:
            cell.set_facecolor('#40466e') # 深蓝色背景
            cell.set_text_props(color='white', weight='bold', fontsize=12)
            cell.set_height(0.15) # 表头高一点
        
        # 数据行样式
        else:
            # 斑马纹背景
            if row % 2 == 0:
                cell.set_facecolor('#f2f2f2')
            else:
                cell.set_facecolor('white')
            
            # 特殊处理：如果是“小计”行，加粗并换个背景色
            # 注意：plot_df的数据行索引从0开始，但table的row从1开始(0是表头)
            # 获取当前行的管家名字
            butler_name = plot_df.iloc[row-1]['管家']
            if '小计' in str(butler_name):
                cell.set_facecolor('#fff3cd') # 浅黄色
                cell.set_text_props(weight='bold')

    # 添加标题
    plt.title(f"{agent_name} - 考核指标详情", fontsize=18, pad=20, fontfamily='Noto Sans SC')
    
    # 将图片保存到内存
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

# --- 4. Streamlit 主界面 ---
st.set_page_config(page_title="代理商报表生成器", layout="wide")

st.title("📊 代理商考核指标长图生成器")
st.markdown("""
上传您的CSV数据文件，系统将自动清洗数据，并按**代理商**生成可视化的考核长图。
""")

# 文件上传
uploaded_file = st.file_uploader("请上传 CSV 文件", type=['csv'])

if uploaded_file is not None:
    try:
        # 处理数据
        df = process_data(uploaded_file)
        
        st.success("数据读取成功！")
        
        # 展示部分预览
        with st.expander("点击查看清洗后的原始数据预览"):
            st.dataframe(df.head(10))
        
        st.divider()
        
        # 获取所有代理商列表
        agents = df['代理商'].unique()
        
        # 选择代理商
        col1, col2 = st.columns([1, 2])
        with col1:
            selected_agent = st.selectbox("选择要生成图片的代理商/门店:", agents)
        
        if selected_agent:
            # 筛选该代理商的数据
            agent_data = df[df['代理商'] == selected_agent]
            
            with col2:
                st.info(f"当前选中: {selected_agent} (共 {len(agent_data)} 行数据)")
            
            # 生成按钮
            if st.button(f"生成 {selected_agent} 的报表图片"):
                with st.spinner("正在绘图..."):
                    img_buffer = generate_long_image(selected_agent, agent_data)
                    
                    # 展示图片
                    st.image(img_buffer, caption=f"{selected_agent} 考核报表", use_container_width=True)
                    
                    # 下载按钮
                    st.download_button(
                        label=f"📥 下载 {selected_agent} 的报表图片",
                        data=img_buffer,
                        file_name=f"{selected_agent}_考核报表.png",
                        mime="image/png"
                    )

    except Exception as e:
        st.error(f"处理文件时发生错误: {e}")
        st.warning("请确保上传的文件格式与描述一致（前两行忽略，第三行指标，第四行列名）。")
