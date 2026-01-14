#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
발파설계 웹앱 - Streamlit 버전
"""
import streamlit as st
import streamlit.components.v1 as components
import math
import os
import io
from datetime import datetime

# 페이지 설정
st.set_page_config(
    page_title="Smart Stem",
    page_icon="https://raw.githubusercontent.com/stark0112/blasting/main/apple-touch-icon.png",
    layout="centered"
)

# iOS/Android 홈화면 아이콘 및 PWA 설정 (JavaScript로 <head>에 직접 주입)
components.html("""
<script>
(function() {
    var iconUrl = 'https://raw.githubusercontent.com/stark0112/blasting/main/apple-touch-icon.png';
    var manifestUrl = 'https://raw.githubusercontent.com/stark0112/blasting/main/manifest.json';

    // 기존 apple-touch-icon 제거
    var existing = document.querySelectorAll('link[rel="apple-touch-icon"], link[rel="apple-touch-icon-precomposed"]');
    existing.forEach(function(el) { el.remove(); });

    // apple-touch-icon 추가
    var link1 = document.createElement('link');
    link1.rel = 'apple-touch-icon';
    link1.href = iconUrl;
    document.head.appendChild(link1);

    var link2 = document.createElement('link');
    link2.rel = 'apple-touch-icon-precomposed';
    link2.href = iconUrl;
    document.head.appendChild(link2);

    // iOS PWA 메타 태그
    var meta1 = document.createElement('meta');
    meta1.name = 'apple-mobile-web-app-capable';
    meta1.content = 'yes';
    document.head.appendChild(meta1);

    var meta2 = document.createElement('meta');
    meta2.name = 'apple-mobile-web-app-status-bar-style';
    meta2.content = 'default';
    document.head.appendChild(meta2);

    var meta3 = document.createElement('meta');
    meta3.name = 'apple-mobile-web-app-title';
    meta3.content = 'Smart Stem';
    document.head.appendChild(meta3);

    // Android manifest
    var manifest = document.createElement('link');
    manifest.rel = 'manifest';
    manifest.href = manifestUrl;
    document.head.appendChild(manifest);

    // theme-color
    var theme = document.createElement('meta');
    theme.name = 'theme-color';
    theme.content = '#1f2937';
    document.head.appendChild(theme);
})();
</script>
""", height=0)

# 심플한 CSS + 인쇄용 CSS
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 900px; }
    h1 { text-align: center; margin-bottom: 2rem; }
    .result-box {
        background: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .pa-title {
        color: #d63031;
        font-size: 1.4rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }

    /* 인쇄용 CSS */
    @media print {
        /* 배경색 흰색으로 */
        * {
            background: white !important;
            color: black !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }

        /* 숨길 요소들 - no-print 클래스 */
        .no-print,
        header,
        footer,
        iframe,
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        [data-testid="stHeader"],
        .stDeployButton,
        #MainMenu {
            display: none !important;
        }

        /* 결과 영역 표시 */
        .print-only {
            display: block !important;
        }

        /* 테이블 스타일 */
        table {
            border-collapse: collapse !important;
        }
        th, td {
            border: 1px solid #333 !important;
            padding: 8px !important;
        }
    }
</style>
""", unsafe_allow_html=True)


# ================= 계산 로직 =================
def compute(K=None, n=None, Vel=None, D=None, Q1=None, C=0.33, V=1.2,
            pd_choice=None, pd_text=None, k1=0.7, V1_theory=1.2):
    def rnd(x, n): return round(x, n)

    have_all = all([K, n, Vel, D])
    Q2 = rnd((D**2) * ((Vel/K)**(2/(-n))), 2) if have_all else None

    if Q1 is None:
        if not have_all:
            raise ValueError("Q1이 비어있을 때는 K, n, Vel, D를 모두 입력해야 합니다.")
        Q3 = Q2
    else:
        Q3 = rnd(min(Q1, Q2), 2) if have_all else rnd(Q1, 2)

    Pa = 1 if Q3 < 0.125 else 2 if Q3 < 0.5 else 3 if Q3 < 1.6 else 4 if Q3 < 5 else 5 if Q3 < 15 else 6

    pd = None
    pd_from_custom = False
    if pd_text:
        try:
            pd = float(pd_text)
            pd_from_custom = True
        except: pass
    if pd is None and pd_choice:
        pd = float(pd_choice)
    if pd is None:
        pd = 0.032 if Pa in [1,2,3] else 0.050 if Pa in [4,5] else 0.076
    pd = rnd(pd, 3)

    pd_msg = None
    if Pa in (1,2) and (pd_from_custom or pd_choice) and pd > 0.032:
        pd = 0.032
        pd_msg = "폭약경이 적합하지 않아 0.032m로 조정되었습니다."

    def anfo(p): return (1000*0.815*3.1415*(p**2))/4.0, 1.0, 0.1
    tol = 1e-9

    W1, h1, nu = {
        1: (0.12, 0.2, 0.5), 2: (0.25, 0.295, 0.5)
    }.get(Pa, (0.25, 0.295, 0.5))

    if Pa >= 3:
        if pd_from_custom:
            W1, h1, nu = anfo(pd)
        elif abs(pd-0.050) < tol:
            W1, h1, nu = 1.0, 0.42, 0.5
        elif abs(pd-0.065) < tol:
            W1, h1, nu = 2.0, 0.52, 0.5 if Pa < 6 else 1.0

    if pd_from_custom and Q3 >= 0.5:
        Q = float(Q3)
        h = h1 * (Q3/W1)
    else:
        Q4 = int((Q3/W1)*2.0) if W1 <= 2.0 else int(Q3)
        Q = (Q4/2.0)*W1 if W1 <= 2.0 else float(Q4)
        h = 0.95 * h1 * Q / W1

    denom = C * V1_theory * (0.7*h + 0.77*(Q**(1/3)) + 10*pd)
    if denom <= 0: raise ValueError("계산 오류")

    B1 = 0.94 * math.sqrt(Q/denom)
    S1 = V1_theory * B1

    if abs(V-1.2) < 1e-12:
        B, S = rnd(B1, 2), rnd(S1, 2)
    else:
        B = math.sqrt((B1*S1)/V)
        S = V * B
        B, S = rnd(B, 2), rnd(S, 2)

    T = rnd((k1*(pd**-0.25) if Pa==1 else k1*(pd**-0.18)) * math.sqrt(B*S), 2)
    H = rnd(T + h, 2)
    K_step = rnd(H - 0.2*B, 2)
    c1 = rnd(Q/(B*S*K_step) if B*S*K_step else 0, 2)

    return {"B": B, "S": S, "T": T, "h": rnd(H-T, 2), "H": H, "Q": Q,
            "c1": c1, "K_step": K_step, "Pa": Pa, "pd": pd, "_msg": pd_msg}


def get_pattern_path(result):
    # 모든 결과에 exam.jpg 사용
    path = os.path.join(os.path.dirname(__file__), "exam.jpg")
    return (path, 1) if os.path.exists(path) else (None, 1)


def make_pdf(result, img_path):
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except: return None

    font = None
    font_paths = [
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf",
        "/usr/share/fonts/opentype/nanum/NanumGothic.otf",
        "C:\\Windows\\Fonts\\malgun.ttf",
        "C:\\Windows\\Fonts\\NanumGothic.ttf",
    ]
    for p in font_paths:
        if os.path.isfile(p):
            try:
                pdfmetrics.registerFont(TTFont("KOR", p))
                font = "KOR"
                break
            except: pass

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    mm = lambda x: x * 72.0 / 25.4

    def setf(s):
        try: c.setFont(font or "Helvetica", s)
        except: c.setFont("Helvetica", s)

    # 타이틀
    setf(16)
    c.drawCentredString(W/2, H-mm(25), "스마트스템 발파설계")

    # 출력날짜 (우측 정렬)
    setf(10)
    output_date = datetime.now().strftime("%Y-%m-%d %H:%M")
    c.drawRightString(W - mm(15), H-mm(35), f"출력날짜: {output_date}")

    # Pa 이름
    pa_names = {1:"미진동발파패턴", 2:"정밀진동제어발파", 3:"소규모진동제어발파",
                4:"중규모진동제어발파", 5:"일반발파", 6:"대규모발파"}
    setf(12)
    c.drawString(mm(25), H-mm(40), pa_names.get(result['Pa'], '일반발파'))

    # 테이블 그리기
    table_x = mm(25)
    table_y = H - mm(50)
    col1_w = mm(35)  # 항목 열 너비
    col2_w = mm(30)  # 값 열 너비
    row_h = mm(8)    # 행 높이

    rows = [
        ("항목", "값"),
        ("저항선 (B)", f"{result['B']:.2f} m"),
        ("공간격 (S)", f"{result['S']:.2f} m"),
        ("전색장 (T)", f"{result['T']:.2f} m"),
        ("장약장 (h)", f"{result['h']:.2f} m"),
        ("천공장 (H)", f"{result['H']:.2f} m"),
        ("계단높이", f"{result['K_step']:.2f} m"),
        ("장약량/공 (Q)", f"{result['Q']} kg"),
        ("비장약량 (c1)", f"{result['c1']} kg/m³"),
        ("폭약경 (pd)", f"{result['pd']} m"),
    ]

    # 테이블 테두리 및 텍스트
    c.setStrokeColorRGB(0.3, 0.3, 0.3)
    c.setLineWidth(0.5)
    for i, (label, value) in enumerate(rows):
        y = table_y - i * row_h
        # 셀 테두리
        c.rect(table_x, y - row_h, col1_w, row_h)
        c.rect(table_x + col1_w, y - row_h, col2_w, row_h)
        # 헤더 배경
        if i == 0:
            c.setFillColorRGB(0.94, 0.94, 0.94)
            c.rect(table_x, y - row_h, col1_w + col2_w, row_h, fill=1)
            c.setFillColorRGB(0, 0, 0)
        # 텍스트
        setf(9 if i == 0 else 9)
        c.drawString(table_x + mm(2), y - row_h + mm(2.5), label)
        c.drawString(table_x + col1_w + mm(2), y - row_h + mm(2.5), value)

    # 이미지 (오른쪽)
    if img_path and os.path.isfile(img_path):
        try:
            ir = ImageReader(img_path)
            iw, ih = ir.getSize()
            img_x = table_x + col1_w + col2_w + mm(15)
            img_max_w = W - img_x - mm(15)
            img_max_h = mm(100)
            scale = min(img_max_w/iw, img_max_h/ih)
            c.drawImage(ir, img_x, H - mm(50) - ih*scale, iw*scale, ih*scale)
        except: pass

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.getvalue()


# ================= UI =================
# 타이틀 (인쇄시 숨김)
st.markdown('<div class="no-print">', unsafe_allow_html=True)
st.title("Smart Stem")
st.caption("Smart Stem v1")
st.markdown('</div>', unsafe_allow_html=True)

# 입력 폼 (인쇄시 숨김)
st.markdown('<div class="no-print">', unsafe_allow_html=True)
with st.form("calc_form"):
    st.subheader("입력값")

    c1, c2 = st.columns(2)
    with c1:
        Q1_in = st.text_input("공당장약량 (kg)", placeholder="미입력시 자동계산",
                              help="입력하지 않으면 이격거리에 따라 산출")
        K_in = st.number_input("K값", value=200.0,
                               help="시험발파추정식 변경 가능")
        n_in = st.number_input("n값", value=-1.60, format="%.2f",
                               help="시험발파추정식 변경 가능")
        Vel_in = st.number_input("허용진동기준치 (cm/sec)", value=0.30, format="%.2f",
                                 help="보안물건의 허용기준치 입력")

    with c2:
        D_in = st.text_input("보안물건 거리 (m)", placeholder="미입력시 무시",
                             help="진동을 고려하고 싶은 경우 입력")
        C_in = st.number_input("발파계수", value=0.33, format="%.2f",
                               help="암질에 따라 풍화암 0.25 ~ 경암 0.5")
        V_in = st.number_input("공간격비율", value=1.2, format="%.2f",
                               help="보통 1.0 ~ 1.25 범위 설정함")
        pd_sel = st.selectbox("폭약직경", ["자동", "0.032", "0.050", "0.065", "직접입력"],
                              help="단위(m), 선택하지 않으면 자동선택")
        pd_custom = st.text_input("폭약직경 직접입력 (m)", placeholder="직접입력 선택시 입력",
                                  help="위에서 '직접입력' 선택 시 이 값이 사용됩니다")

    k1_sel = st.radio("목적", ["비산제어(0.7)", "파쇄도개선(0.55)", "광산채석장(0.5)"], horizontal=True)

    submitted = st.form_submit_button("계산", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# 계산 실행
if submitted:
    try:
        Q1 = float(Q1_in) if Q1_in.strip() else None
        D = float(D_in) if D_in.strip() else None
        pd_choice = pd_sel if pd_sel in ["0.032", "0.050", "0.065"] else None
        k1 = {"비산제어(0.7)": 0.7, "파쇄도개선(0.55)": 0.55, "광산채석장(0.5)": 0.5}[k1_sel]

        result = compute(K=K_in, n=n_in, Vel=Vel_in, D=D, Q1=Q1, C=C_in, V=V_in,
                        pd_choice=pd_choice, pd_text=pd_custom if pd_sel=="직접입력" else None, k1=k1)

        st.session_state["result"] = result
        st.session_state["img_path"], st.session_state["idx"] = get_pattern_path(result)

        if result.get("_msg"):
            st.warning(result["_msg"])

    except Exception as e:
        st.error(f"오류: {e}")

# 결과 표시
if "result" in st.session_state:
    r = st.session_state["result"]
    img = st.session_state.get("img_path")

    st.divider()

    pa_names = {1:"미진동발파패턴", 2:"정밀진동제어발파", 3:"소규모진동제어발파",
                4:"중규모진동제어발파", 5:"일반발파", 6:"대규모발파"}

    st.markdown(f"### {pa_names.get(r['Pa'], '일반발파')}")

    col1, col2 = st.columns([1, 1.8], vertical_alignment="top")

    with col1:
        st.markdown("""
        <style>
        [data-testid="stMarkdownContainer"] table td,
        [data-testid="stMarkdownContainer"] table th {
            padding: 11px 10px !important;
        }
        </style>
        """, unsafe_allow_html=True)
        st.markdown(f"""
| 항목 | 값 |
|------|-----|
| 저항선 (B) | **{r['B']:.2f} m** |
| 공간격 (S) | **{r['S']:.2f} m** |
| 전색장 (T) | **{r['T']:.2f} m** |
| 장약장 (h) | **{r['h']:.2f} m** |
| 천공장 (H) | **{r['H']:.2f} m** |
| 계단높이 | **{r['K_step']:.2f} m** |
| 장약량/공 (Q) | **{r['Q']} kg** |
| 비장약량 (c1) | **{r['c1']} kg/m³** |
| 폭약경 (pd) | **{r['pd']} m** |
""")

    with col2:
        if img and os.path.exists(img):
            st.image(img, use_container_width=True)
        else:
            st.info("패턴 이미지 없음")

    # 버튼
    st.divider()
    b1, b2, _ = st.columns([1, 1, 2])
    pdf = make_pdf(r, img)

    with b1:
        if pdf:
            st.download_button("PDF 저장", pdf, "발파설계결과.pdf", "application/pdf", use_container_width=True)

    with b2:
        # 인쇄용 HTML 생성
        import base64
        img_b64 = ""
        if img and os.path.exists(img):
            with open(img, "rb") as f:
                img_b64 = base64.b64encode(f.read()).decode()

        pa_name = pa_names.get(r['Pa'], '일반발파')
        output_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        print_html = f'''<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>스마트스템 발파설계</title>
<style>
body {{ font-family: 'Malgun Gothic', sans-serif; padding: 30px; }}
h2 {{ text-align: center; margin-bottom: 5px; }}
.date-line {{ text-align: right; margin-bottom: 15px; font-size: 12px; color: #555; }}
h3 {{ margin-bottom: 20px; }}
.container {{ display: flex; gap: 30px; align-items: flex-start; }}
.left {{ flex: 0 0 auto; }}
.right {{ flex: 1; display: flex; align-items: flex-start; }}
table {{ border-collapse: collapse; }}
th, td {{ border: 1px solid #333; padding: 8px 15px; text-align: left; }}
th {{ background: #f0f0f0; }}
.right img {{ height: 380px; width: auto; object-fit: contain; }}
</style>
</head><body>
<h2>스마트스템 발파설계</h2>
<div class="date-line">출력날짜: {output_date}</div>
<h3>{pa_name}</h3>
<div class="container">
<div class="left">
<table>
<tr><th>항목</th><th>값</th></tr>
<tr><td>저항선 (B)</td><td>{r['B']:.2f} m</td></tr>
<tr><td>공간격 (S)</td><td>{r['S']:.2f} m</td></tr>
<tr><td>전색장 (T)</td><td>{r['T']:.2f} m</td></tr>
<tr><td>장약장 (h)</td><td>{r['h']:.2f} m</td></tr>
<tr><td>천공장 (H)</td><td>{r['H']:.2f} m</td></tr>
<tr><td>계단높이</td><td>{r['K_step']:.2f} m</td></tr>
<tr><td>장약량/공 (Q)</td><td>{r['Q']} kg</td></tr>
<tr><td>비장약량 (c1)</td><td>{r['c1']} kg/m³</td></tr>
<tr><td>폭약경 (pd)</td><td>{r['pd']} m</td></tr>
</table>
</div>
<div class="right">
{"<img src='data:image/jpeg;base64," + img_b64 + "'/>" if img_b64 else "<p>패턴 이미지 없음</p>"}
</div>
</div>
<script>window.onload=function(){{window.print();}}</script>
</body></html>'''

        print_html_b64 = base64.b64encode(print_html.encode('utf-8')).decode('ascii')

        # 인쇄 버튼을 HTML로 직접 렌더링 (PDF 버튼 스타일과 동일)
        components.html(f'''
        <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ margin: 0; padding: 0; }}
        button {{
            display: flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            padding: 0.25rem 0.75rem;
            min-height: 38.4px;
            background-color: rgb(19, 23, 32);
            color: rgb(250, 250, 250);
            border: 1px solid rgba(250, 250, 250, 0.2);
            border-radius: 0.5rem;
            cursor: pointer;
            font-size: 1rem;
            font-weight: 400;
            font-family: "Source Sans Pro", sans-serif;
        }}
        button:hover {{
            border-color: rgb(250, 250, 250);
            color: rgb(250, 250, 250);
        }}
        button:active {{
            background-color: rgb(250, 250, 250);
            color: rgb(19, 23, 32);
        }}
        </style>
        <button onclick="openPrint()">인쇄</button>
        <script>
        function openPrint() {{
            var w = window.open('', '_blank', 'width=800,height=600');
            if(w) {{
                // UTF-8 디코딩
                var binary = atob('{print_html_b64}');
                var bytes = new Uint8Array(binary.length);
                for (var i = 0; i < binary.length; i++) {{
                    bytes[i] = binary.charCodeAt(i);
                }}
                var html = new TextDecoder('utf-8').decode(bytes);
                w.document.write(html);
                w.document.close();
            }} else {{
                alert('팝업이 차단되었습니다. 팝업 차단을 해제해주세요.');
            }}
        }}
        </script>
        ''', height=42)

st.markdown('<div class="no-print">', unsafe_allow_html=True)
st.divider()
st.caption("Smart Stem v1 - 발파설계 계산기")
st.markdown('</div>', unsafe_allow_html=True)
