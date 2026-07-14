import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import numpy as np

# 웹 페이지 기본 설정
st.set_page_config(
    page_title="📊 DYE COMPATIBILITY GRAPH VIEWER", 
    page_icon="📊",  # <- 여기에 원하는 이모지를 직접 넣으시면 됩니다.
    layout="centered"
)

st.title("📊 T/S 상용성 그래프 만들기")
st.write("엑셀 데이터를 표에 바로 붙여넣거나 파일을 업로드하여 부드러운 상용성 그래프를 생성합니다.")

# 1. 그래프 기본 정보 및 염료 정보 입력 섹션
st.subheader("📝 그래프 및 염료 정보 입력")

# 그래프 제목 입력 필드
graph_title = st.text_input("Enter Graph Title:", value="", placeholder="예: HP Combi.")

# 염료명과 함량을 입력하는 세션
col1, col2 = st.columns(2)
with col1:
    dye_name_1 = st.text_input("Enter Dye 1 Name:", value="", placeholder="예: Sunfix Amber HP")
    dye_name_2 = st.text_input("Enter Dye 2 Name:", value="", placeholder="예: Sunfix Red HP")
    dye_name_3 = st.text_input("Enter Dye 3 Name:", value="", placeholder="예: Sunfix Space HP")
with col2:
    dye_amount_1_raw = st.text_input("Enter Dye 1 Amount (숫자만 입력):", value="", placeholder="예: 0.82")
    dye_amount_2_raw = st.text_input("Enter Dye 2 Amount (숫자만 입력):", value="", placeholder="예: 0.90")
    dye_amount_3_raw = st.text_input("Enter Dye 3 Amount (숫자만 입력):", value="", placeholder="예: 0.70")

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

label_1 = f"{dye_name_1} {dye_amount_1}".strip() if dye_name_1 else "Dye 1"
label_2 = f"{dye_name_2} {dye_amount_2}".strip() if dye_name_2 else "Dye 2"
label_3 = f"{dye_name_3} {dye_amount_3}".strip() if dye_name_3 else "Dye 3"


# 2. 데이터 입력
st.subheader("⚙️ 데이터 입력")
input_method = st.radio(
    "데이터 입력 방식을 선택해 주세요:",
    ["웹 표(Grid)에 직접 붙여넣기 (추천)", "엑셀 파일 업로드 (.xlsx)"]
)

raw_data = None

# [방식 1] 웹 표(Grid)에 붙여넣기
if "표" in input_method:
    st.write("👇 **기본 입력된 시간/온도 축을 확인하시고, 오른쪽 염료 칸(Dye 1 ~ 3)에 데이터를 입력해 주세요.**")
    
    # 💡 요청하신 기본 시간/온도 축 값 정의 (20과 22 사이는 빈 칸으로 띄움)
    default_x = [0.0, 5.0, 10.0, 20.0, np.nan, 22.0, 25.0, 30.0, 40.0, 60.0, 80.0, 100.0]
    
    # 시간/온도 축이 미리 들어가 있는 템플릿 표 생성
    template_df = pd.DataFrame({
        'Time/Temp': default_x,
        'Dye 1': [np.nan] * len(default_x),
        'Dye 2': [np.nan] * len(default_x),
        'Dye 3': [np.nan] * len(default_x)
    })
    
    # 대화형 표 생성 (행 추가/삭제 가능)
    edited_df = st.data_editor(template_df, num_rows="dynamic", use_container_width=True)
    
    # 기본 시간축 외에 실제 염료 데이터(Dye 1, 2, 3)가 최소 한 칸이라도 입력되었을 때만 그래프 그리기 작동
    if not edited_df.iloc[:, 1:].dropna(how='all').empty:
        raw_data = edited_df

# [방식 2] 파일 업로드
else:
    uploaded_file = st.file_uploader("엑셀 파일을 선택해주세요 (.xlsx)", type=['xlsx'])
    if uploaded_file is not None:
        try:
            raw_data = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다. 오류: {e}")

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
            fig, ax = plt.subplots(figsize=(8, 6.5))
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

            # 메인 타이틀 및 축 설정
            title_text = graph_title if graph_title.strip() else "Dye Compatibility Graph"
            ax.set_title(title_text, fontsize=20, fontweight='bold', pad=50)
            ax.set_xlabel('Time / Temp', fontsize=12, labelpad=10)
            ax.set_ylabel('Exhaustion (%)', fontsize=12, labelpad=10)
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # 범례 생성 (글자 크기 15, 테두리 설정)
            leg = ax.legend(fontsize=15, loc='lower right', framealpha=1.0, edgecolor='#CCCCCC')
            
            # 범례 내의 텍스트 색상을 각 그래프 선 색상과 일치
            for i, text in enumerate(leg.get_texts()):
                text.set_color(colors[i % len(colors)])
                text.set_weight('bold')

            # Y축 및 X축 범위를 0부터 120까지 고정
            ax.set_ylim(0, 120)
            ax.set_xlim(0, 120)

            # Alkali Dosing 주석 배치
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

            # 웹 화면 출력
            st.pyplot(fig)
            st.info("💡 팁: 완료된 그래프를 마우스 우클릭하여 바로 '이미지 복사'할 수 있습니다.")

    except Exception as e:
        st.error(f"그래프를 생성하는 중 오류가 발생했습니다. 표에 문자가 섞여있는지 확인해 주세요. 오류: {e}")