import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import numpy as np
import io
import os
import re
import base64
from openpyxl import Workbook
from openpyxl.drawing.image import Image as xlImage

# 웹 페이지 기본 설정
st.set_page_config(
    page_title="T/S 상용성 그래프",
    page_icon="logo.png" if os.path.exists("logo.png") else "📈",
    layout="wide"
)

# ==========================================
# 0. 글꼴: 영어·숫자 = Outfit, 한글 = 나눔스퀘어 네오 (static/fonts 폴더의 글꼴 파일)
#    - Streamlit Cloud / config.toml 이 있으면: Streamlit이 static 폴더를 직접 내보냄
#    - 내 PC에서 config 없이 실행하면: PC 안에서만 쓰는 통로(127.0.0.1)로 내보냄
#    - 글꼴 파일이 없으면 예전처럼 Pretendard(인터넷) → 맑은 고딕
# ==========================================
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_FONT_DIR = os.path.join(_APP_DIR, "static", "fonts")
_OY_NEO = [("aLt", "100 349"), ("bRg", "350 549"), ("cBd", "550 749"), ("dEb", "750 849"), ("eHv", "850 900")]
_OY_OUTFIT = "outfit-latin-wght-normal.woff2"
_OY_UR = ("U+0000-001F,U+0021-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,"
          "U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD")
_FONT_STACK = "'OY Outfit', 'OY NanumSquare Neo', 'Pretendard', -apple-system, 'Malgun Gothic', sans-serif"

@st.cache_resource
def _oy_font_server():
    """config 없이 내 PC에서 실행할 때만 쓰는 글꼴 통로 (.woff2 파일만 내보냄)"""
    import threading, functools
    from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
    class _H(SimpleHTTPRequestHandler):
        def do_GET(self):
            _name = self.path.split("?")[0].lstrip("/")
            if "/" in _name or "\\" in _name or not _name.lower().endswith(".woff2"):
                self.send_error(404); return
            return super().do_GET()
        def end_headers(self):
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Cache-Control", "max-age=31536000")
            super().end_headers()
        def log_message(self, *a):
            pass
    srv = ThreadingHTTPServer(("127.0.0.1", 0), functools.partial(_H, directory=_FONT_DIR))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{srv.server_address[1]}"

def _oy_font_base():
    if not os.path.exists(os.path.join(_FONT_DIR, "NanumSquareNeoTTF-bRg.woff2")):
        return ""
    try:
        if st.get_option("server.enableStaticServing"):
            return "app/static/fonts"
    except Exception:
        pass
    try:
        return _oy_font_server()
    except Exception:
        return ""

_FONT_BASE = _oy_font_base()
_FONT_FACES = ""
if _FONT_BASE:
    _FONT_FACES = (f"@font-face {{ font-family: 'OY Outfit'; font-style: normal; font-weight: 100 900; font-display: swap; "
                   f"src: url('{_FONT_BASE}/{_OY_OUTFIT}') format('woff2-variations'), url('{_FONT_BASE}/{_OY_OUTFIT}') format('woff2'); unicode-range: {_OY_UR}; }}\n"
                   + "\n".join(f"@font-face {{ font-family: 'OY NanumSquare Neo'; font-style: normal; font-weight: {_w}; font-display: swap; "
                               f"src: url('{_FONT_BASE}/NanumSquareNeoTTF-{_c}.woff2') format('woff2'); }}" for _c, _w in _OY_NEO))
    # 표(데이터 입력 칸) · 입력칸 글자도 같은 글꼴로 (config.toml에 이미 있으면 그대로, 없으면 처음 한 번 설정하고 다시 그림)
    try:
        from streamlit import config as _st_config
        _THEME_FONT = "OY Outfit, OY NanumSquare Neo, sans-serif"
        if _st_config.get_option("theme.font") != _THEME_FONT or _FONT_BASE not in str(_st_config.get_option("theme.fontFaces") or ""):
            _st_config.set_option("theme.fontFaces",
                [{"family": "OY Outfit", "url": f"{_FONT_BASE}/{_OY_OUTFIT}", "weight": "100 900", "unicodeRange": _OY_UR}]
                + [{"family": "OY NanumSquare Neo", "url": f"{_FONT_BASE}/NanumSquareNeoTTF-{_c}.woff2", "weight": _w} for _c, _w in _OY_NEO])
            _st_config.set_option("theme.font", _THEME_FONT)
            _st_config.set_option("theme.primaryColor", "#1b489d")
            st.rerun()
    except Exception:
        pass

# 그래프(matplotlib) 글꼴: 영어·숫자 Outfit → 한글 나눔스퀘어 네오 (글자마다 자동으로 나눠 씀), 파일이 없으면 맑은 고딕
from matplotlib import font_manager as _fm
_mpl_fams = []
for _fam, _files in (("Outfit", ("Outfit-Regular.ttf", "Outfit-Bold.ttf")), ("NanumSquare Neo", ("NanumSquareNeo-bRg.ttf", "NanumSquareNeo-cBd.ttf"))):
    _ok = False
    for _f in _files:
        _p = os.path.join(_FONT_DIR, _f)
        if os.path.exists(_p):
            try:
                _fm.fontManager.addfont(_p); _ok = True
            except Exception:
                pass
    if _ok: _mpl_fams.append(_fam)
plt.rcParams['font.family'] = _mpl_fams + ['Malgun Gothic', 'AppleGothic', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False   # 마이너스 기호 깨짐 방지

# ==========================================
# 1. 상단 바 로고 (네이비 바탕용 logo1.png → 없으면 logo.png)
# ==========================================
def _b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""
logo_base64 = _b64("logo1.png") or _b64("logo.png")

# ==========================================
# 2. 화면 디자인 (Spectra50 · SpectraMatch · UV-Vis · Dye Finder · ColorMap과 같은 네이비 상단 바)
# ==========================================
_CSS = """
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined" rel="stylesheet" />
<style>
__FONT_FACES__
:root {
    --navy: #1b489d; --navy-2: #2557b8; --navy-d: #133777;
    --soft: #e9eef8; --soft-2: #f3f6fc;
    --bg: #f4f6fa; --card: #ffffff; --ink: #1d2433; --muted: #6b7487; --line: #e2e7f0;
}
.stApp { background: var(--bg) !important; }
.stApp, .stApp p, .stApp label, .stApp button, .stApp input, .stApp textarea,
.stApp h1, .stApp h2, .stApp h3, .stApp li, .stApp [data-testid="stMarkdownContainer"] {
    font-family: __FONT_STACK__;
    color: var(--ink);
}
.material-symbols-outlined { font-family: 'Material Symbols Rounded', 'Material Symbols Outlined' !important; font-weight: normal; font-style: normal;
    letter-spacing: normal; text-transform: none; white-space: nowrap; direction: ltr; font-feature-settings: 'liga'; -webkit-font-feature-settings: 'liga'; line-height: 1; vertical-align: middle; }
header[data-testid="stHeader"], #MainMenu, footer { display: none !important; }
[data-testid="stToolbar"], [data-testid="stAppDeployButton"], .stDeployButton { display: none !important; }
.block-container { padding: 76px 1.75rem 1.5rem 1.75rem !important; max-width: none !important; }
/* 화면에 안 보이는 칸(스타일만 있는 칸 · 고정 상단 바 칸) 세로 간격 제거 */
.block-container [data-testid="stElementContainer"]:has(> [data-testid="stMarkdown"] style):not(:has(p, img, span, table)),
.block-container [data-testid="stElementContainer"]:has(.fixed-header) { margin-bottom: -1rem !important; }

/* ---------- 상단 바 ---------- */
.fixed-header {
    position: fixed; top: 0; left: 0; right: 0; height: 60px; z-index: 999990;
    background: linear-gradient(100deg, var(--navy) 0%, var(--navy-2) 100%);
    box-shadow: 0 4px 18px rgba(19,55,119,.28);
    display: flex; align-items: center; gap: 12px; padding: 0 20px;
}
.fixed-header img { height: 38px; width: auto; max-width: 160px; object-fit: contain; }
.fixed-header .mark { width: 34px; height: 34px; border-radius: 9px; background: rgba(255,255,255,.15); color: #fff;
    display: flex; align-items: center; justify-content: center; }
.fixed-header .t1 { font-size: 21px; font-weight: 800; letter-spacing: -.3px; line-height: 1; color: #fff; }
.fixed-header .hdr-sub { font-size: 12.5px; font-weight: 600; color: #fff; background: rgba(255,255,255,.16);
    border: 1px solid rgba(255,255,255,.22); padding: 3px 10px; border-radius: 999px; }

/* ---------- 페이지 제목 ---------- */
.page-title { font-size: 24px; font-weight: 800; color: var(--ink); letter-spacing: -.5px; margin: 0; display: flex; align-items: center; gap: 10px; }
.page-title .ico { width: 36px; height: 36px; border-radius: 10px; background: var(--navy); color: #fff;
    display: inline-flex; align-items: center; justify-content: center; box-shadow: 0 3px 10px rgba(27,72,157,.28); }
.page-title .ico .material-symbols-outlined { font-size: 22px; color: #fff; }
.page-sub { color: var(--muted); font-size: 13.5px; margin: 6px 0 14px 0; }

/* ---------- 카드 ---------- */
.st-key-card_input, .st-key-card_graph {
    background: var(--card); border: 1px solid var(--line) !important; border-radius: 14px;
    padding: 18px 20px 16px 20px !important;
    box-shadow: 0 1px 2px rgba(16,24,40,.04), 0 4px 14px rgba(16,24,40,.04);
}
.sec { display: flex; align-items: center; gap: 8px; font-size: 15px; font-weight: 800; color: var(--ink); margin: 2px 0 6px 0; letter-spacing: -.2px; }
.sec .n { width: 22px; height: 22px; border-radius: 7px; background: var(--navy); color: #fff; font-size: 12px;
    display: inline-flex; align-items: center; justify-content: center; font-weight: 800; }
.sec .material-symbols-outlined { font-size: 20px; color: var(--navy); }
.sec .opt { margin-left: auto; font-size: 11.5px; font-weight: 700; color: var(--navy); background: var(--soft); border-radius: 999px; padding: 2px 9px; }
.hint { font-size: 12.5px; color: var(--muted); margin: -2px 0 8px 0; }
.dye-tag { display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 800; color: var(--ink); margin: 6px 0 -6px 0; }
.dye-tag i { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
.sec-gap { height: 10px; }

/* ---------- 입력 · 버튼 ---------- */
.stTextInput label p { font-size: 12.5px !important; font-weight: 700 !important; color: var(--muted) !important; }
[data-baseweb="input"] { border-radius: 10px !important; }
[data-testid="stTextInput"] [data-baseweb="input"]:focus-within { border-color: var(--navy) !important; }
.stDownloadButton > button {
    border-radius: 10px !important; border: 1px solid var(--navy) !important; background: var(--navy) !important; color: #fff !important;
    font-weight: 700 !important; min-height: 40px !important; box-shadow: 0 2px 6px rgba(27,72,157,.25) !important;
}
.stDownloadButton > button p, .stDownloadButton > button span { color: #fff !important; }
.stDownloadButton > button:hover { background: var(--navy-d) !important; border-color: var(--navy-d) !important; }
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { border-radius: 12px !important; overflow: hidden; border: 1px solid var(--line) !important; }
[data-testid="stAlert"] { border-radius: 12px !important; }

/* ---------- 그래프 자리 ---------- */
[data-testid="stImage"] { display: flex; justify-content: center; }
[data-testid="stImage"] img { border-radius: 10px; }
.empty { border: 1.5px dashed #c6cfdf; border-radius: 14px; background: #fafbfd; padding: 46px 20px; text-align: center; color: var(--muted); }
.empty .ic { width: 54px; height: 54px; border-radius: 14px; background: var(--soft); margin: 0 auto 12px; display: flex; align-items: center; justify-content: center; }
.empty .ic .material-symbols-outlined { font-size: 28px; color: var(--navy); }
.empty .tt { font-size: 16px; font-weight: 800; color: var(--ink); }
.empty .ds { font-size: 13px; margin-top: 6px; }
</style>
"""
_CSS = _CSS.replace("__FONT_FACES__", _FONT_FACES).replace("__FONT_STACK__", _FONT_STACK)
st.markdown("\n".join(l for l in re.sub(r"/\*.*?\*/", "", _CSS, flags=re.S).splitlines() if l.strip()), unsafe_allow_html=True)

def icon(name):
    return f"<span class='material-symbols-outlined'>{name}</span>"

_logo = (f"<img src='data:image/png;base64,{logo_base64}'>" if logo_base64 else f"<div class='mark'>{icon('show_chart')}</div>")
st.markdown(f"""
<div class="fixed-header">{_logo}<div class="t1">T/S 상용성 그래프</div><span class="hdr-sub">Compatibility Graph</span></div>
<div class="page-title"><span class="ico">{icon('show_chart')}</span>T/S 상용성 그래프 만들기</div>
<div class="page-sub">염료 정보와 시간별 염착 데이터를 넣으면 상용성 그래프를 그리고, 그래프와 데이터를 엑셀 파일로 내려받을 수 있습니다.</div>
""", unsafe_allow_html=True)

# ==========================================
# 3. 메인 화면 시작 
# ==========================================

# 화면을 좌/우 두 개의 열로 나눔 (비율 1.1 : 1)
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left, st.container(key="card_input"):
    # 그래프 기본 정보 및 염료 정보 입력 섹션
    header_sub_col, btn_sub_col = st.columns([0.65, 0.35])
    
    with header_sub_col:
        st.markdown(f"<div class='sec'><span class='n'>1</span>그래프 · 염료 정보</div>", unsafe_allow_html=True)
    
    with btn_sub_col:
        download_btn_placeholder = st.empty()

    graph_title = st.text_input("Graph Title:", value="", placeholder="예: HP Combi.")

    # 염료마다 한 줄: [염료명 | 함량], 왼쪽 색 점 = 그래프 선 색 (노랑 · 빨강 · 파랑)
    _dye_rows = []
    for _i, (_ph_n, _ph_a, _dot) in enumerate([("예: Sunfix Amber HP", "예: 0.82", "#FFB300"),
                                                ("예: Sunfix Red HP", "예: 0.90", "#E53935"),
                                                ("예: Sunfix Space HP", "예: 0.70", "#1E88E5")], start=1):
        st.markdown(f"<div class='dye-tag'><i style='background:{_dot}'></i>Dye {_i}</div>", unsafe_allow_html=True)
        _cn, _ca = st.columns([1.6, 1])
        with _cn:
            _n = st.text_input(f"Dye {_i} Name:", value="", placeholder=_ph_n)
        with _ca:
            _a = st.text_input(f"Dye {_i} % o.w.f (숫자만 입력):", value="", placeholder=_ph_a)
        _dye_rows.append((_n, _a))
    (dye_name_1, dye_amount_1_raw), (dye_name_2, dye_amount_2_raw), (dye_name_3, dye_amount_3_raw) = _dye_rows

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
    
    calculated_font_size = int(460 / max_label_len)
    
    # 폰트 최대 크기를 18로 제한하여 너무 꽉 차는 것을 방지
    font_size = max(10, min(18, calculated_font_size))

    st.markdown("<div class='sec-gap'></div><div class='sec'><span class='n'>2</span>데이터 입력</div>"
                "<div class='hint'>첫 줄 Time(분), 아래 세 줄에 Dye 1~3 염착 값을 넣거나 엑셀에서 복사해 붙여넣으세요. 비어 있는 칸(20분과 22분 사이)이 알칼리 투입 구간입니다.</div>", unsafe_allow_html=True)
    
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

with col_right, st.container(key="card_graph"):
    st.markdown(f"<div class='sec'>{icon('show_chart')}그래프 결과<span class='opt'>엑셀에 그래프 · 데이터 함께 저장</span></div>", unsafe_allow_html=True)
    
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
                    bbox_to_anchor=(0.22, 0.03),  
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
        st.markdown(f"<div class='empty'><div class='ic'>{icon('show_chart')}</div><div class='tt'>그래프가 여기에 표시됩니다</div>"
                    f"<div class='ds'>왼쪽 표에 데이터를 넣거나 붙여넣으면 바로 그려집니다.</div></div>", unsafe_allow_html=True)