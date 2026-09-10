import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import numpy as np
import io
import os
import base64
from openpyxl import Workbook
from openpyxl.drawing.image import Image as xlImage

# 맑은 고딕 폰트 및 마이너스 기호 깨짐 방지 전역 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 웹 페이지 기본 설정
st.set_page_config(
    page_title="상용성그래프 만들기", 
    layout="wide"
)

# ==========================================
# 1. 로고 이미지를 Base64로 인코딩 (표준 UI 적용)
# ==========================================
try:
    with open("logo.png", "rb") as image_file:
        logo_base64 = base64.b64encode(image_file.read()).decode()
except Exception:
    logo_base64 = ""

# ==========================================
# 2. 고정 메뉴바 및 UI 커스텀 CSS 주입
# ==========================================
st.markdown(f"""
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet" />
<style>
    /* 1. Streamlit 기본 상단 헤더, 메뉴, 푸터 숨기기 */
    [data-testid="stHeader"], #MainMenu, footer {{
        display: none !important;
    }}
    
    /* 2. 🌟 새로운 상단 고정 메뉴바 디자인 🌟 */
    .fixed-header {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 60px;
        background-color: #ffffff;
        box-shadow: 0px 2px 10px rgba(0,0,0,0.1);
        z-index: 999998;
        display: flex;
        align-items: center;
        padding-left: 20px;
        border-bottom: 1px solid #eaeaea;
    }}
    .fixed-header img {{
        width: 45px;
        margin-right: 12px;
    }}
    .fixed-header h2 {{
        margin: 0;
        padding: 0;
        font-size: 24px;
        font-weight: 700;
        color: #31333F;
    }}

    /* 3. 본문 상단 여백 설정 (상단바에 가려지지 않도록) */
    .block-container {{
        padding-top: 80px !important; 
        padding-bottom: 0rem !important;
    }}
    
    /* 머티리얼 아이콘 정렬 */
    .material-symbols-outlined {{
        line-height: 1 !important;
        vertical-align: middle;
    }}
</style>

<!-- 상단 메뉴바 HTML 렌더링 -->
<div class="fixed-header">
    <img src="data:image/png;base64,{logo_base64}" onerror="this.style.display='none'">
    <h2>T/S 상용성 그래프 만들기</h2>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 3. 메인 화면 시작 
# ==========================================

# 화면을 좌/우 두 개의 열로 나눔 (비율 1.1 : 1)
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:
    # 그래프 기본 정보 및 염료 정보 입력 섹션
    header_sub_col, btn_sub_col = st.columns([0.65, 0.35])
    
    with header_sub_col:
        st.markdown("<h3 style='display: flex; align-items: center; margin-top: 0;'><span class='material-symbols-outlined' style='margin-right:8px;'>edit_document</span>그래프 및 염료 정보 입력</h3>", unsafe_allow_html=True)
    
    with btn_sub_col:
        st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)
        download_btn_placeholder = st.empty()

    graph_title = st.text_input("Graph Title:", value="", placeholder="예: HP Combi.")

    sub_col1, sub_col2 = st.columns(2)
    with sub_col1:
        dye_name_1 = st.text_input("Dye 1 Name:", value="", placeholder="예: Sunfix Amber HP")
        dye_name_2 = st.text_input("Dye 2 Name:", value="", placeholder="예: Sunfix Red HP")
        dye_name_3 = st.text_input("Dye 3 Name:", value="", placeholder="예: Sunfix Space HP")
    with sub_col2:
        dye_amount_1_raw = st.text_input("Dye 1 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.82")
        dye_amount_2_raw = st.text_input("Dye 2 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.90")
        dye_amount_3_raw = st.text_input("Dye 3 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.70")

    def format_dye_amount(amount_str):
        val = amount_str.strip()
        if not val: return ""
        if "%" in val:
            if "o.w.f" in val: return val
            return f"{val} o.w.f"
        return f"{val}% o.w.f"

    dye_amount_1 = format_dye_amount(dye_amount_1_raw)
    dye_amount_2 = format_dye_amount(dye_amount_2_raw)
    dye_amount_3 = format_dye_amount(dye_amount_3_raw)

    name_lengths = [len(dye_name_1), len(dye_name_2), len(dye_name_3)]
    max_len = max(name_lengths) + 3 

    def align_label(name, amount, default_name):
        if not name: return default_name
        return f"{name.ljust(max_len)}{amount}"

    label_1 = align_label(dye_name_1, dye_amount_1, "Dye 1")
    label_2 = align_label(dye_name_2, dye_amount_2, "Dye 2")
    label_3 = align_label(dye_name_3, dye_amount_3, "Dye 3")

    # 1. 폰트 크기 최적화 (상수를 600 -> 480으로 낮춰 표 밖으로 나가는 현상 방지)
    max_label_len = max(len(label_1), len(label_2), len(label_3))
    
    calculated_font_size = int(440 / max_label_len)
    
    # 폰트 최대 크기를 18로 제한하여 너무 꽉 차는 것을 방지
    font_size = max(10, min(18, calculated_font_size))

    st.markdown("<h3 style='display: flex; align-items: center;'><span class='material-symbols-outlined' style='margin-right:8px;'>settings</span>데이터 입력</h3>", unsafe_allow_html=True)
    
    # ... (중략: 데이터프레임 처리 및 그래프 그리는 부분은 기존과 동일하게 유지) ...

    raw_data = None
    default_x_visible = [5.0, 10.0, 20.0, np.nan, 22.0, 25.0, 30.0, 40.0, 60.0, 80.0, 100.0]
    blank_columns = ["\u200b" * (i + 1) for i in range(len(default_x_visible))]

    template_df = pd.DataFrame([
        default_x_visible,
        [np.nan] * len(default_x_visible),
        [np.nan] * len(default_x_visible),
        [np.nan] * len(default_x_visible)
    ], index=['Time', 'Dye 1', 'Dye 2', 'Dye 3'], columns=blank_columns)

    col_config = {col: st.column_config.NumberColumn(width="small") for col in blank_columns}
    edited_df = st.data_editor(template_df, use_container_width=True, column_config=col_config)

    point_1_df = pd.DataFrame([[0.0], [0.0], [0.0], [0.0]], index=['Time', 'Dye 1', 'Dye 2', 'Dye 3'], columns=["Hidden_P1"])
    full_df = pd.concat([point_1_df, edited_df], axis=1)

    transposed_df = full_df.transpose()
    transposed_df.columns = ['Time', 'Dye 1', 'Dye 2', 'Dye 3']
    transposed_df.reset_index(drop=True, inplace=True)

    if not transposed_df.iloc[1:][['Dye 1', 'Dye 2', 'Dye 3']].dropna(how='all').empty:
        raw_data = transposed_df

with col_right:
    st.markdown("<h3 style='display: flex; align-items: center;'><span class='material-symbols-outlined' style='margin-right:8px;'>show_chart</span>그래프 결과</h3>", unsafe_allow_html=True)
    
    def plot_smooth_segment(ax, x, y, color, label=None, linewidth=4.0):
        mask = ~np.isnan(x) & ~np.isnan(y)
        x = x[mask]
        y = y[mask]
        
        if len(x) >= 3:
            x_new = np.linspace(x.min(), x.max(), 300)
            f = interp1d(x, y, kind='quadratic')
            y_new = f(x_new)
            y_new = np.clip(y_new, 0, 100)
            ax.plot(x_new, y_new, color=color, label=label, linewidth=linewidth)
        elif len(x) > 0:
            ax.plot(x, y, color=color, label=label, linewidth=linewidth)

    if raw_data is not None:
        try:
            df = raw_data.copy()
            df.columns = ['X'] + list(df.columns[1:])
            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            df = df.dropna(subset=['X'])
            
            df1 = df[df['X'] <= 20]
            df2 = df[df['X'] >= 22]

            if not df1.empty or not df2.empty:
                fig, ax = plt.subplots(figsize=(8, 6.2))
                fig.subplots_adjust(top=0.80)

                color_yellow = '#FFB300' 
                color_red = '#E53935'    
                color_blue = '#1E88E5'   
                colors = [color_yellow, color_red, color_blue]

                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[1]].values, color=color_yellow, label=label_1)
                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[2]].values, color=color_red, label=label_2)
                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[3]].values, color=color_blue, label=label_3)

                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[1]].values, color=color_yellow)
                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[2]].values, color=color_red)
                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[3]].values, color=color_blue)

                title_text = graph_title if graph_title.strip() else "Dye Compatibility Graph"
                ax.set_title(title_text, fontsize=24, fontweight='bold', pad=40)
                
                ax.set_xlabel('')
                ax.set_ylabel('')
                ax.grid(True, linestyle='--', alpha=0.7)
                
                leg = ax.legend(
                    loc='lower left',             
                    bbox_to_anchor=(0.25, 0.03),  
                    framealpha=0.9,               
                    edgecolor='#CCCCCC',
                    prop={'family': 'monospace', 'size': font_size},
                    borderpad=0.6,                
                    labelspacing=0.4              
                )
                
                for i, text in enumerate(leg.get_texts()):
                    text.set_color(colors[i % len(colors)])
                    text.set_weight('bold')

                ax.set_ylim(0, 120)
                ax.set_xlim(0, 120)

                ax.annotate(
                    'Alkali\nDosing', 
                    xy=(20, 120), 
                    xytext=(20, 122),  # 기존 127에서 122로 낮춰 제목과 거리를 둠
                    arrowprops=dict(arrowstyle="->", color='black', lw=1.5), # 화살표 두께도 1.5로 살짝 얇게 조절
                    fontsize=11,       # 기존 14에서 11로 글씨 크기 축소
                    fontweight='bold', 
                    color='black',
                    horizontalalignment='center', 
                    verticalalignment='bottom',
                    annotation_clip=False
                )

                fig.savefig('temp_graph.png', dpi=600, bbox_inches='tight')
                st.image('temp_graph.png')

                wb = Workbook()
                ws = wb.active
                ws.title = "Graph Data"

                ws.append(["그래프 제목", title_text])
                ws.append([])
                ws.append(["염료 구분", "염료명", "함량 (% o.w.f)"])
                ws.append(["Dye 1", dye_name_1, dye_amount_1_raw])
                ws.append(["Dye 2", dye_name_2, dye_amount_2_raw])
                ws.append(["Dye 3", dye_name_3, dye_amount_3_raw])
                ws.append([])

                ws.append(["[입력된 데이터 (Time & Dyes)]"])
                for index, row in full_df.iterrows():
                    ws.append([index] + row.tolist())

                img_for_excel = xlImage('temp_graph.png')
                img_for_excel.width = img_for_excel.width * 0.5   
                img_for_excel.height = img_for_excel.height * 0.5
                ws.add_image(img_for_excel, 'B15')

                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                download_filename = f"{graph_title if graph_title else 'Graph_Data'}.xlsx"

                download_btn_placeholder.download_button(
                    label="엑셀 파일 다운로드",
                    data=excel_buffer,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    icon=":material/download:"
                )

        except Exception as e:
            st.error(f"그래프를 생성하는 중 오류가 발생했습니다. 오류: {e}", icon=":material/error:")
    else:
        st.info("왼쪽 표에 데이터를 붙여넣으면 여기에 그래프가 표시됩니다.", icon=":material/arrow_back:")