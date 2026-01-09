import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.table import Table
import io
import os
import requests

# --- 1. 字体配置 (修复版) ---
@st.cache_resource
def get_font_name():
    """下载中文字体，注册到 Matplotlib，并返回字体名称"""
    font_url = "https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC-Regular.ttf"
    font_path = "NotoSansSC-Regular.ttf"
    
    # 1. 下载字体
    if not os.path.exists(font_path):
        with st.spinner("正在下载中文字体..."):
            try:
                r = requests.get(font_url)
                with open(font_path, "wb") as f:
                    f.write(r.content)
            except Exception as e:
                st.error(f"字体下载失败: {e}")
                return "sans-serif" # 失败回退

    # 2. 注册字体并获取名称
    try:
        fm.fontManager.addfont(font_path)
        prop = fm.FontProperties(fname=font_path)
        return prop.get_name() # 返回 'Noto Sans SC'
    except Exception as e:
        st.error(f"字体注册警告: {e}")
        return "sans-serif"

# --- 2. 考核配置 ---
TARGETS = {
    "DCC首呼": 0.95, "DCC二呼": 0.90, "邀约开口率": 80.0, "加微开口率": 80.0,
    "试乘试驾满意度": 4.80, "试驾排程率": 0.90, "试驾后次日回访率": 0.90,
    "试乘试驾满意度4.5分问卷占比": 0.90, "交易协助满意度": 4.80, "车辆交付满意度": 4.80
}

def get_target(col_name):
    """根据大指标名称匹配目标值"""
    if not col_name: return None, None
    target_val, target_name = None, ""
    for k, v in TARGETS.items():
        if k in str(col_name):
            if target_name == "" or len(k) > len(target_name):
                target_val, target_name = v, k
    return target_val, target_name

def parse_val(v):
    """转数值"""
    try:
        if pd.isna(v) or str(v).strip() in ["-", ""]: return None
        return float(str(v).replace('%', '').strip())
    except: return None

# --- 3. 数据处理 (保留表头结构) ---
def process_data(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=None, dtype=str)
    else:
        df = pd.read_excel(file, header=None, dtype=str, engine='openpyxl')
    
    # 提取表头结构
    # header_L1: 第一行表头 (指标名称)
    # header_L2: 第二行表头 (分子/分母)
    header_L1 = df.iloc[2].fillna(method='ffill').tolist()
    header_L2 = df.iloc[3].tolist()
    
    # 清洗表头
    clean_L1, clean_L2, unique_cols = [], [], []
    for i, (h1, h2) in enumerate(zip(header_L1, header_L2)):
        h1 = str(h1).strip() if pd.notna(h1) else ""
        h2 = str(h2).strip() if pd.notna(h2) else ""
        
        # 修复空值逻辑
        if h1 == "" or h1.lower() == "nan": h1 = h2
        if h2 == "" or h2.lower() == "nan": h2 = h1
        
        clean_L1.append(h1)
        clean_L2.append(h2)
        # 创建唯一列名用于DataFrame索引
        unique_cols.append(f"{i}_{h1}_{h2}")

    # 处理数据体
    data = df.iloc[4:].copy()
    data.columns = unique_cols
    
    # 标准化前两列
    cols = list(data.columns)
    if len(cols) > 0: cols[0] = "base_代理商"
    if len(cols) > 1: cols[1] = "base_管家"
    data.columns = cols
    
    data['base_代理商'] = data['base_代理商'].fillna(method='ffill')
    data = data.dropna(how='all')
    
    # 将表头结构存入 attrs 供绘图使用
    # 结构: [(H1, H2, ColKey), ...]
    headers_struct = list(zip(clean_L1, clean_L2, unique_cols))
    data.attrs['headers'] = headers_struct
    
    return data

# --- 4. 生成考核结果 ---
def calc_status(row, headers_map):
    failures = []
    # 遍历所有列，找到指标列进行判断
    for h1, h2, col_key in headers_map:
        if "指标" in h2: # 只看叫"指标"的列
            target, t_name = get_target(h1)
            if target is not None:
                val = parse_val(row.get(col_key))
                if val is not None:
                    # 量级对齐
                    comp_val = val
                    if target <= 1.0 and val > 1.0: comp_val = val / 100.0
                    
                    if comp_val < target:
                        # 格式化
                        t_str = f"{target:.0%}" if target <=1.0 else f"{target}"
                        a_str = f"{comp_val:.1%}" if target <=1.0 else f"{val}"
                        failures.append(f"{t_name}:\n{a_str} / {t_str}")
    
    return "👍 全部合格" if not failures else "\n".join(failures)

# --- 5. 绘图 (双层表头核心) ---
def generate_complex_image(agent_name, agent_data):
    # 修改处：获取字体名称字符串，而非对象
    font_family = get_font_name()
    
    # 1. 准备数据和表头
    headers_all = agent_data.attrs['headers'] # [(H1, H2, Key), ...]
    
    # --- 过滤逻辑 (修改处) ---
    headers_plot = []
    for i, (h1, h2, key) in enumerate(headers_all):
        if i == 0: continue # 去掉代理商列 (index 0)
        
        # 核心过滤：如果第二行表头是 "分子" 或 "分母"，则跳过
        if h2 in ["分子", "分母"]:
            continue
            
        headers_plot.append((h1, h2, key))
    
    # 增加“考核结果”列
    # 在 headers_plot 末尾追加
    headers_plot.append(("考核结论", "结果", "calc_status"))
    
    # 计算每一行的数据显示矩阵
    plot_data = [] # 二维列表
    
    for _, row in agent_data.iterrows():
        row_vals = []
        # 计算状态
        status_txt = calc_status(row, headers_all) # 注意：计算状态还是用全量数据
        
        for h1, h2, key in headers_plot:
            if key == "calc_status":
                row_vals.append(status_txt)
            else:
                val = row.get(key, "")
                row_vals.append(val)
        plot_data.append(row_vals)

    # 2. 构建绘图用的全表内容 (Header Rows + Data Rows)
    # Row 0: H1 (Metric Names)
    # Row 1: H2 (Sub Columns)
    # Row 2+: Data
    
    table_content = []
    
    # Row 0 & 1
    row0 = [x[0] for x in headers_plot]
    row1 = [x[1] for x in headers_plot]
    table_content.append(row0)
    table_content.append(row1)
    # Data
    table_content.extend(plot_data)
    
    # 3. 尺寸计算
    num_cols = len(headers_plot)
    num_rows = len(table_content)
    
    # 计算行高：扫描数据行，看换行符数量
    row_heights = []
    # Header rows 固定高度
    row_heights.extend([1.2, 1.0]) 
    
    for r_idx in range(2, num_rows):
        # 这一行所有单元格中最大的换行数
        max_newlines = 0
        for c_val in table_content[r_idx]:
            max_newlines = max(max_newlines, str(c_val).count('\n'))
        # 基础高度 1.0，每多一行文字增加 0.4
        row_heights.append(1.0 + max_newlines * 0.45)
        
    total_h = sum(row_heights) * 0.5 + 2
    total_w = max(16, num_cols * 1.5 + 3) # 稍微宽一点
    
    fig, ax = plt.subplots(figsize=(total_w, total_h))
    ax.axis('off')
    
    # 4. 绘制表格
    # bbox=[0, 0, 1, 1] 让表格充满画布
    table = ax.table(cellText=table_content, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    
    # 5. 精细化样式调整 (Merging & Colors)
    cells = table.get_celld()
    
    # Helper: Check if adjacent cells in Row 0 are same
    def is_same_as_prev(c_idx):
        if c_idx == 0: return False
        return headers_plot[c_idx][0] == headers_plot[c_idx-1][0]

    for (row, col), cell in cells.items():
        # 修改处：使用 fontfamily 参数，而不是 fontproperties 对象
        cell.set_text_props(fontfamily=font_family, padding=10)
        
        # --- Row 0: Metric Headers (Top Level) ---
        if row == 0:
            cell.set_facecolor('#40466e') # 深蓝
            cell.set_text_props(color='white', weight='bold', size=13, fontfamily=font_family)
            cell.set_height(row_heights[row] * 0.04) # 归一化高度调整
            
            # 视觉合并逻辑
            if is_same_as_prev(col):
                # 简单合并视觉效果
                pass
            
        # --- Row 1: Sub Headers (Second Level) ---
        elif row == 1:
            cell.set_facecolor('#5a629e') # 浅一点的蓝
            cell.set_text_props(color='white', weight='bold', size=11, fontfamily=font_family)
            cell.set_height(row_heights[row] * 0.04)

        # --- Data Rows ---
        else:
            # 原始数据索引
            data_row_idx = row - 2
            
            # 斑马纹
            bg = '#f2f2f2' if row % 2 == 0 else 'white'
            
            # 小计行高亮
            # 这里的 col=0 对应的是 headers_plot[0]，即“管家”
            butler_name = str(table_content[row][0])
            if '小计' in butler_name:
                bg = '#fff3cd'
                font_weight = 'bold'
            else:
                font_weight = 'normal'
            
            cell.set_facecolor(bg)
            
            # 字体颜色逻辑
            txt_color = 'black'
            
            # 1. 考核结果列 (最后一列)
            if col == num_cols - 1:
                cell_text = cell.get_text().get_text()
                if "全部合格" in cell_text:
                    txt_color = '#2e7d32' # 深绿
                    font_weight = 'bold'
                else:
                    txt_color = '#c62828' # 深红
                    cell.set_text_props(ha='left') # 左对齐
            
            # 2. 普通数据列标红
            else:
                h1, h2, _ = headers_plot[col]
                cell_val = table_content[row][col]
                
                # 判断是否红字
                if "指标" in h2:
                    t_val, _ = get_target(h1)
                    if t_val is not None:
                        v_num = parse_val(cell_val)
                        if v_num is not None:
                            c_v = v_num if (t_val > 1.0 or v_num <= 1.0) else v_num/100.0
                            if c_v < t_val:
                                txt_color = '#d32f2f'
            
            cell.set_text_props(color=txt_color, weight=font_weight, fontfamily=font_family)
            
            # 动态高度
            cell.set_height(row_heights[row] * 0.05)

    # 标题
    plt.title(f"{agent_name} - 门店考核报表", fontsize=20, pad=30, fontfamily=font_family, color='#333333')
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=200) # 提高DPI使文字更清晰
    plt.close(fig)
    buf.seek(0)
    return buf

# --- 6. Streamlit App ---
st.set_page_config(page_title="门店考核报表V2", layout="wide")
st.title("📊 门店考核报表生成器 (专业版)")
st.markdown("""
上传数据文件，生成带有**双层表头**和**智能考核判定**的专业报表。
(已自动隐藏分子、分母列，只显示核心指标)
""")

f = st.file_uploader("上传 Excel/CSV", type=['xlsx', 'csv'])

if f:
    try:
        df = process_data(f)
        st.success("数据加载成功")
        
        agents = df['base_代理商'].unique()
        sel = st.selectbox("选择门店:", agents)
        
        if sel and st.button("生成报表"):
            with st.spinner("正在生成高清长图..."):
                sub_df = df[df['base_代理商'] == sel]
                img = generate_complex_image(sel, sub_df)
                st.image(img, use_container_width=True)
                st.download_button("下载图片", img, f"{sel}_考核报表.png", "image/png")
                
    except Exception as e:
        st.error(f"出错: {e}")
