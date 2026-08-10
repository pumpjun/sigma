import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image as xlImage

# 맑은 고딕 폰트 및 마이너스 기호 깨짐 방지 전역 설정 (제목 및 주석 등에 적용)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 화면의 상하 여백을 확 줄여서 스크롤을 방지하고 Material Icon을 로드하는 CSS 추가
hide_streamlit_style = """
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet" />
<style>
/* Material Icon 자체 여백 및 수직 정렬 보정 */
.material-symbols-outlined {
    line-height: 1 !important;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 0rem;
}
</style>
"""

# 웹 페이지 기본 설정 (이모티콘 탭 아이콘 제거)
st.set_page_config(
    page_title="상용성그래프 만들기", 
    layout="wide"
)

st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 메인 타이틀 (수직 정렬 적용)
st.markdown(
    "<h1 style='display: flex; align-items: center; margin-top: 0;'><span class='material-symbols-outlined' style='font-size:36px; margin-right:12px;'>bar_chart</span>T/S 상용성 그래프 만들기</h1>", 
    unsafe_allow_html=True
)
st.write("엑셀 데이터를 표에 바로 붙여넣어 부드러운 상용성 그래프를 생성합니다.")

# 화면을 좌/우 두 개의 열로 나눔 (비율 1.1 : 1)
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:
    # 1. 그래프 기본 정보 및 염료 정보 입력 섹션
    # 제목과 다운로드 버튼을 나란히 배치하기 위해 열을 나눔
    header_col, btn_col = st.columns([0.65, 0.35])
    
    with header_col:
        st.markdown("<h3 style='display: flex; align-items: center; margin-top: 0;'><span class='material-symbols-outlined' style='margin-right:8px;'>edit_document</span>그래프 및 염료 정보 입력</h3>", unsafe_allow_html=True)
    
    with btn_col:
        # 제목과 높이를 맞추기 위한 여백 추가
        st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)
        # 나중에 엑셀이 생성된 뒤 버튼을 넣을 빈 공간(placeholder) 미리 생성
        download_btn_placeholder = st.empty()

    # 그래프 제목 입력 필드
    graph_title = st.text_input("Graph Title:", value="", placeholder="예: HP Combi.")

    # 염료명과 함량을 입력하는 세션
    sub_col1, sub_col2 = st.columns(2)
    with sub_col1:
        dye_name_1 = st.text_input("Dye 1 Name:", value="", placeholder="예: Sunfix Amber HP")
        dye_name_2 = st.text_input("Dye 2 Name:", value="", placeholder="예: Sunfix Red HP")
        dye_name_3 = st.text_input("Dye 3 Name:", value="", placeholder="예: Sunfix Space HP")
    with sub_col2:
        dye_amount_1_raw = st.text_input("Dye 1 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.82")
        dye_amount_2_raw = st.text_input("Dye 2 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.90")
        dye_amount_3_raw = st.text_input("Dye 3 % o.w.f (숫자만 입력):", value="", placeholder="예: 0.70")

    # 함량 뒤에 '% o.w.f'를 자동으로 붙여주는 편리한 변환 함수
    def format_dye_amount(amount_str):
        val = amount_str.strip()
        if not val:
            return ""
        if "%" in val:
            if "o.w.f" in val:
                return val
            return f"{val} o.w.f"
        return f"{val}% o.w.f"

    # 범례에 표시할 텍스트 가공
    dye_amount_1 = format_dye_amount(dye_amount_1_raw)
    dye_amount_2 = format_dye_amount(dye_amount_2_raw)
    dye_amount_3 = format_dye_amount(dye_amount_3_raw)

    # 이름 길이를 맞춰서 줄맞춤하기
    name_lengths = [len(dye_name_1), len(dye_name_2), len(dye_name_3)]
    max_len = max(name_lengths) + 3 

    # 지정된 길이(max_len)만큼 공간을 확보하고 왼쪽 정렬해주는 함수
    def align_label(name, amount, default_name):
        if not name:
            return default_name
        return f"{name.ljust(max_len)}{amount}"

    label_1 = align_label(dye_name_1, dye_amount_1, "Dye 1")
    label_2 = align_label(dye_name_2, dye_amount_2, "Dye 3")
    label_3 = align_label(dye_name_3, dye_amount_3, "Dye 3")

    # 글자 길이에 따라 폰트 크기 자동 조절
    max_label_len = max(len(label_1), len(label_2), len(label_3))

    if max_label_len > 40:      
        font_size = 10
    elif max_label_len > 30:    
        font_size = 12
    elif max_label_len > 22:    
        font_size = 14
    else:                       
        font_size = 17

    # 2. 데이터 입력 (가로형 표, Point 1 숨김)
    st.markdown("<h3 style='display: flex; align-items: center;'><span class='material-symbols-outlined' style='margin-right:8px;'>settings</span>데이터 입력</h3>", unsafe_allow_html=True)

    raw_data = None

    # 기본 시간축 값 정의 (Point 1인 0.0을 제외한 나머지 포인트)
    default_x_visible = [5.0, 10.0, 20.0, np.nan, 22.0, 25.0, 30.0, 40.0, 60.0, 80.0, 100.0]

    # 폭이 없는 공백 문자(\u200b)를 사용하여 열 제목을 완전히 보이지 않게 하고, 너비 계산의 차이를 없앰
    blank_columns = ["\u200b" * (i + 1) for i in range(len(default_x_visible))]

    # 눈에 보이는 표 템플릿 생성
    template_df = pd.DataFrame([
        default_x_visible,
        [np.nan] * len(default_x_visible),
        [np.nan] * len(default_x_visible),
        [np.nan] * len(default_x_visible)
    ], index=['Time', 'Dye 1', 'Dye 2', 'Dye 3'], columns=blank_columns)

    # 모든 열의 너비를 "small"로 통일되게 강제 적용
    col_config = {col: st.column_config.NumberColumn(width="small") for col in blank_columns}

    # 대화형 표 생성 (왼쪽 화면에 배치)
    edited_df = st.data_editor(template_df, use_container_width=True, column_config=col_config)

    # 숨겨진 Point 1 데이터 생성 (무조건 0으로 고정)
    point_1_df = pd.DataFrame([
        [0.0],
        [0.0],
        [0.0],
        [0.0]
    ], index=['Time', 'Dye 1', 'Dye 2', 'Dye 3'], columns=["Hidden_P1"])

    # Point 1과 사용자가 입력한 데이터를 하나로 합침
    full_df = pd.concat([point_1_df, edited_df], axis=1)

    # 그래프를 그리기 위해 세로형(행=시간, 열=염료)으로 변환
    transposed_df = full_df.transpose()
    transposed_df.columns = ['Time', 'Dye 1', 'Dye 2', 'Dye 3']
    transposed_df.reset_index(drop=True, inplace=True)

    # 실제 염료 데이터(Dye 1, 2, 3)가 입력되었는지 확인 (Point 1의 0값 제외)
    if not transposed_df.iloc[1:][['Dye 1', 'Dye 2', 'Dye 3']].dropna(how='all').empty:
        raw_data = transposed_df


# 오른쪽 화면: 그래프 출력 및 엑셀 다운로드
with col_right:
    st.markdown("<h3 style='display: flex; align-items: center;'><span class='material-symbols-outlined' style='margin-right:8px;'>show_chart</span>그래프 결과</h3>", unsafe_allow_html=True)
    
    # 데이터를 부드러운 곡선으로 보간하여 그려주는 함수
    def plot_smooth_segment(ax, x, y, color, label=None, linewidth=2.5):
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

    # 데이터가 성공적으로 로드되었을 때 그래프 그리기 시작
    if raw_data is not None:
        try:
            df = raw_data.copy()
            df.columns = ['X'] + list(df.columns[1:])
            
            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            df = df.dropna(subset=['X'])
            
            # 20과 22 구간 끊기 설정
            df1 = df[df['X'] <= 20]
            df2 = df[df['X'] >= 22]

            if not df1.empty or not df2.empty:
                # 그래프 사이즈를 조정
                fig, ax = plt.subplots(figsize=(8, 6.2))
                fig.subplots_adjust(top=0.80)

                # 지정된 가독성 높은 색상 코드
                color_yellow = '#FFB300'  # 노란색
                color_red = '#E53935'     # 빨간색
                color_blue = '#1E88E5'    # 파란색
                colors = [color_yellow, color_red, color_blue]

                # 세그먼트 1 (X <= 20) 그리기
                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[1]].values, color=color_yellow, label=label_1)
                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[2]].values, color=color_red, label=label_2)
                plot_smooth_segment(ax, df1['X'].values, df1[df.columns[3]].values, color=color_blue, label=label_3)

                # 세그먼트 2 (X >= 22) 그리기
                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[1]].values, color=color_yellow)
                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[2]].values, color=color_red)
                plot_smooth_segment(ax, df2['X'].values, df2[df.columns[3]].values, color=color_blue)

                # 메인 타이틀 설정 (맑은 고딕 적용됨)
                title_text = graph_title if graph_title.strip() else "Dye Compatibility Graph"
                ax.set_title(title_text, fontsize=20, fontweight='bold', pad=40)
                
                # 축 이름(글자) 삭제
                ax.set_xlabel('')
                ax.set_ylabel('')
                
                ax.grid(True, linestyle='--', alpha=0.7)
                
                # 범례 폰트를 다시 고정폭(monospace)으로 변경하여 줄맞춤 복구
                leg = ax.legend(
                    loc='lower right', 
                    framealpha=1.0, 
                    edgecolor='#CCCCCC',
                    prop={'family': 'monospace', 'size': font_size}
                )
                
                # 범례 내의 텍스트 색상을 각 그래프 선 색상과 일치
                for i, text in enumerate(leg.get_texts()):
                    text.set_color(colors[i % len(colors)])
                    text.set_weight('bold')

                # Y축 및 X축 범위를 0부터 120까지 고정
                ax.set_ylim(0, 120)
                ax.set_xlim(0, 120)

                # Alkali Dosing 주석 배치 (맑은 고딕 적용됨)
                ax.annotate(
                    'Alkali\nDosing', 
                    xy=(20, 120), 
                    xytext=(20, 127),
                    arrowprops=dict(arrowstyle="->", color='black', lw=1.5),
                    fontsize=11, 
                    fontweight='bold', 
                    color='black',
                    horizontalalignment='center', 
                    verticalalignment='bottom',
                    annotation_clip=False
                )

                # 웹 화면 출력 및 고해상도 저장 처리
                fig.savefig('temp_graph.png', dpi=300, bbox_inches='tight')
                st.image('temp_graph.png')

                # ==========================================
                # 엑셀 파일 생성 및 다운로드 로직
                # ==========================================
                wb = Workbook()
                ws = wb.active
                ws.title = "Graph Data"

                # 1. 엑셀에 그래프 기본 정보 및 염료 정보 쓰기
                ws.append(["그래프 제목", title_text])
                ws.append([])
                ws.append(["염료 구분", "염료명", "함량 (% o.w.f)"])
                ws.append(["Dye 1", dye_name_1, dye_amount_1_raw])
                ws.append(["Dye 2", dye_name_2, dye_amount_2_raw])
                ws.append(["Dye 3", dye_name_3, dye_amount_3_raw])
                ws.append([])

                # 2. 엑셀에 표 데이터 쓰기 (웹 화면에 보이는 가로형 그대로)
                ws.append(["[입력된 데이터 (Time & Dyes)]"])
                
                # full_df의 인덱스(Time, Dye 1 등)를 첫 번째 열로, 나머지 데이터를 가로로 엑셀에 기록
                for index, row in full_df.iterrows():
                    ws.append([index] + row.tolist())

                # 3. 엑셀에 생성된 그래프 이미지 삽입 
                # 가로로 길어진 데이터 아래쪽(B15 근처)에 깔끔하게 배치합니다.
                img_for_excel = xlImage('temp_graph.png')
                img_for_excel.width = img_for_excel.width * 0.5   # 엑셀 내 이미지 크기를 절반(50%)으로 축소
                img_for_excel.height = img_for_excel.height * 0.5
                ws.add_image(img_for_excel, 'B15')

                # 메모리 버퍼에 엑셀 저장
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                
                # 파일명 생성 (제목이 없으면 기본 이름 설정)
                download_filename = f"{graph_title if graph_title else 'Graph_Data'}.xlsx"

                # 스트림릿 다운로드 버튼 생성 (이모지 제거, 구글 Material 아이콘 적용)
                # 이 코드는 위에서 만든 왼쪽 제목 옆 빈 공간(placeholder)에 렌더링됩니다.
                download_btn_placeholder.download_button(
                    label="엑셀 파일 다운로드",
                    data=excel_buffer,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    icon=":material/download:"
                )

        except Exception as e:
            # 에러 아이콘도 구글 Material 아이콘 적용
            st.error(f"그래프를 생성하는 중 오류가 발생했습니다. 표에 문자가 섞여있는지 확인해 주세요. 오류: {e}", icon=":material/error:")
    else:
        # 안내 아이콘도 구글 Material 아이콘 적용
        st.info("왼쪽 표에 데이터를 붙여넣으면 여기에 그래프가 표시됩니다.", icon=":material/arrow_back:")