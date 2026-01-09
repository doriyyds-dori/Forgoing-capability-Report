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
    font_path = "NotoSansSC-Regular. ttf"
    
    if not os.path.exists(font_path):
        with st.spinner("正在下载中文字体..."):
            try:
                r = requests.get(font_url)
                with open(font_path, "wb") as f:
                    f.write(r.content)
            except Exception as e:
                st. error(f"字体下载失败:  {e}")
                return "sans-serif"

    try: 
        fm.fontManager.addfont(font_path)
        return "Noto Sans SC"
    except Exception as e:
        st. error(f"字体注册警告: {e}")
        return "sans-serif"

# --- 2. 考核配置 ---
TARGETS = {
    "DCC首呼": 0.95, "DCC二呼": 0.90, "邀约开口率": 80.0, "加微开口率": 80.0,
    "试乘试驾满意度":  4.80, "试��排程率": 0.90, "试驾后次日回访率": 0.90,
    "试乘试驾满意度4.5分问卷占比": 0.90, "交易协助满意度": 4.80, "车辆交付满意度": 4.80
}

def get_target(col_name):
    """根据大指标名称匹配目标值"""
    if not col_name: 
        return None, None
    target_val, target_name = None, ""
    for k, v in TARGETS.items():
        if k in str(col_name):
            if target_name == "" or len(k) > len(target_name):
                target_val, target_name = v, k
    return target_val, target_name

def parse_val(v):
    """转数值"""
    try:
        if pd.isna(v) or str(v).strip() in ["-", ""]:
            return None
        return float(str(v).replace('%', '').strip())
    except:
        return None

# --- 3. 数据处理 (兼容 openpyxl 错误) ---
def process_data(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=None, dtype=str)
    else:
        # 尝试多种方式读取 Excel
        try:
            # 方法1: 使用 openpyxl (默认)
            df = pd. read_excel(file, header=None, dtype=str, engine='openpyxl')
        except TypeError as e:
            if "InlineFont" in str(e):
                # openpyxl 版本兼容性问题，尝试其他引擎
                st.warning("⚠️ 检测到 Excel 文件格式兼容性问题，尝试使用备用方式读取...")
                try:
                    # 方法2: 尝试 xlrd (适用于 . xls)
                    df = pd.read_excel(file, header=None, dtype=str, engine='xlrd')
                except:
                    # 方法3: 提示用户转换格式
                    st.error("""
                    ❌ **Excel 文件读取失败！**
                    
                    **原因：** 您的 Excel 文件格式与当前环境不兼容（openpyxl 库版本问题）
                    
                    **解决方案：**
                    1. 在 Excel 中打开文件，另存为 `.csv` 格式后重新上传
                    2. 或者在 Excel 中"另存为" → 选择 "Excel 工作簿 (. xlsx)" 重新保存
                    3. 或者使用 WPS/LibreOffice 打开并重新保存
                    """)
                    raise
            else:
                raise
    
    # 提取表头结构
    header_L1 = df.iloc[2]. ffill().tolist()
    header_L2 = df.iloc[3]. tolist()
    
    # 清洗表头
    clean_L1, clean_L2, unique_cols = [], [], []
    for i, (h1, h2) in enumerate(zip(header_L1, header_L2)):
        h1 = str(h1).strip() if pd.notna(h1) else ""
        h2 = str(h2).strip() if pd.notna(h2) else ""
        
        if h1 == "" or h1. lower() == "nan":
            h1 = h2
        if h2 == "" or h2.lower() == "nan":
            h2 = h1
        
        clean_L1.append(h1)
        clean_L2.append(h2)
        unique_cols.append(f"{i}_{h1}_{h2}")

    # 处理数据体
    data = df.iloc[4:].copy()
    data.columns = unique_cols
    
    # 标准化前两列
    cols = list(data.columns)
    if len(cols) > 0:
        cols[0] = "base_代理商"
    if len(cols) > 1:
        cols[1] = "base_管家"
    data. columns = cols
    
    data['base_代理商'] = data['base_代理商']. ffill()
    data = data.dropna(how='all')
    
    headers_struct = list(zip(clean_L1, clean_L2, unique_cols))
    data. attrs['headers'] = headers_struct
    
    return data

# --- 4. 生成考核结果 ---
def calc_status(row, headers_map):
    failures = []
    for h1, h2, col_key in headers_map:
        if "指标" in h2:
            target, t_name = get_target(h1)
            if target is not None: 
                val = parse_val(row. get(col_key))
                if val is not None: 
                    comp_val = val
                    if target <= 1.0 and val > 1.0:
                        comp_val = val / 100.0
                    
                    if comp_val < target:
                        t_str = f"{target:.0%}" if target <= 1.0 else f"{target}"
                        a_str = f"{comp_val:.1%}" if target <= 1.0 else f"{val}"
                        failures.append(f"{t_name}:\n{a_str} / {t_str}")
    
    return "👍 全部合格" if not failures else "\n".join(failures)

# --- 5. 绘图 (双层表头核心) ---
def generate_complex_image(agent_name, agent_data):
    font_family = get_font_name()
    
    # 全局设置字体
    plt.rcParams['font.family'] = font_family
    plt.rcParams['font.sans-serif'] = [font_family]
    
    headers_all = agent_data.attrs['headers']
    
    # 过滤逻辑
    headers_plot = []
    for i, (h1, h2, key) in enumerate(headers_all):
        if i == 0:
            continue
        if h2 in ["分子", "分母"]: 
            continue
        headers_plot. append((h1, h2, key))
    
    headers_plot.append(("考核结论", "结果", "calc_status"))
    
    # 计算每一行的数据
    plot_data = []
    for _, row in agent_data.iterrows():
        row_vals = []
        status_txt = calc_status(row, headers_all)
        
        for h1, h2, key in headers_plot:
            if key == "calc_status":
                row_vals.append(status_txt)
            else:
                val = row.get(key, "")
                row_vals.append(val)
        plot_data.append(row_vals)

    # 构建表格内容
    table_content = []
    row0 = [x[0] for x in headers_plot]
    row1 = [x[1] for x in headers_plot]
    table_content.append(row0)
    table_content.append(row1)
    table_content.extend(plot_data)
    
    # 尺寸计算
    num_cols = len(headers_plot)
    num_rows = len(table_content)
    
    row_heights = [1.2, 1.0]
    for r_idx in range(2, num_rows):
        max_newlines = 0
        for c_val in table_content[r_idx]:
            max_newlines = max(max_newlines, str(c_val).count('\n'))
        row_heights.append(1.0 + max_newlines * 0.45)
        
    total_h = sum(row_heights) * 0.5 + 2
    total_w = max(16, num_cols * 1.5 + 3)
    
    fig, ax = plt.subplots(figsize=(total_w, total_h))
    ax.axis('off')
    
    # 绘制表格
    table = ax.table(cellText=table_content, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    
    # 样式调整
    cells = table.get_celld()

    for (row, col), cell in cells.items():
        # Row 0: 第一层表头
        if row == 0:
            cell.set_facecolor('#40466e')
            cell.set_text_props(color='white', weight='bold', size=13)
            cell.set_height(row_heights[row] * 0.04)
            
        # Row 1: 第二层表头
        elif row == 1:
            cell. set_facecolor('#5a629e')
            cell.set_text_props(color='white', weight='bold', size=11)
            cell.set_height(row_heights[row] * 0.04)

        # 数据行
        else:
            bg = '#f2f2f2' if row % 2 == 0 else 'white'
            
            butler_name = str(table_content[row][0])
            if '小计' in butler_name:
                bg = '#fff3cd'
                font_weight = 'bold'
            else:
                font_weight = 'normal'
            
            cell.set_facecolor(bg)
            txt_color = 'black'
            
            # 考核结果列
            if col == num_cols - 1:
                cell_text = cell.get_text().get_text()
                if "全部合格" in cell_text:
                    txt_color = '#2e7d32'
                    font_weight = 'bold'
                else:
                    txt_color = '#c62828'
                    cell. set_text_props(ha='left')
            
            # 普通数据列标红逻辑
            else:
                h1, h2, _ = headers_plot[col]
                cell_val = table_content[row][col]
                
                if "指标" in h2:
                    t_val, _ = get_target(h1)
                    if t_val is not None:
                        v_num = parse_val(cell_val)
                        if v_num is not None:
                            c_v = v_num if (t_val > 1.0 or v_num <= 1.0) else v_num / 100.0
                            if c_v < t_val: 
                                txt_color = '#d32f2f'
            
            cell.set_text_props(color=txt_color, weight=font_weight)
            cell.set_height(row_heights[row] * 0.05)

    plt.title(f"{agent_name} - 门店考核报表", fontsize=20, pad=30, color='#333333')
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf

# --- 6. Streamlit App ---
st.set_page_config(page_title="门店考核报表V2", layout="wide")
st.title("📊 门店考核报表生成器 (专业版)")
st.markdown("""
上传数据文件，生成带有**双层表头**和**智能考核判定**的专业报表。
(已自动隐藏分子、分母列，只显示核心指标)

⚠️ **如果 Excel 文件上传失败，请：**
- 将文件另存为 CSV 格式后重新上传
- 或使用 Excel 重新保存为 .xlsx 格式
""")

f = st.file_uploader("上传 Excel/CSV", type=['xlsx', 'xls', 'csv'])

if f:
    try:
        df = process_data(f)
        st.success("✅ 数据加载成功")
        
        agents = df['base_代理商'].unique()
        sel = st.selectbox("选择门店:", agents)
        
        if sel and st.button("生成报表"):
            with st.spinner("正在生成高清长图..."):
                sub_df = df[df['base_代理商'] == sel]
                img = generate_complex_image(sel, sub_df)
                st.image(img, use_container_width=True)
                st.download_button("📥 下载图片", img, f"{sel}_考核报表.png", "image/png")
                
    except Exception as e:
        st.error(f"❌ 出错:  {e}")
        import traceback
        st.code(traceback.format_exc())
