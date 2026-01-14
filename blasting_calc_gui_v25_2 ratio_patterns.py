#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
발파 기본연산 계산기 v25 (화면 자동 축소 + PDF + 패턴 이미지)
- 화면이 작으면 창/글꼴/패널을 비례 축소하여 UI가 잘리지 않게 표시
- 이미지: 크롭 금지. 항상 축소하여 여백 안에 맞춤(상 15mm, 하 15mm 이상, 좌우 30mm)
- 폭약직경 라디오: 기본 미선택, ANFO(직접입력) 시 입력값을 pd로 우선 적용
- [ANFO 분기] 선택 시 Q=Q3, Q>=0.5면 h=h1*(Q3/W1), Q<0.5면 기본(비-ANFO) 경로
- [Pa=1,2 규칙] 사용자가 pd를 입력/선택했고 pd>0.032면 0.032로 강제 + 안내 메시지
- PDF: "발파설계결과" 중앙 정렬, 위 30mm/왼쪽 30mm 여백
      패턴 이미지는 좌로 10mm, 위로 10mm 이동(제목과 충돌 방지)
"""
import os, sys, io, math, base64, tempfile
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont

AUTO_CROP = False  # True로 하면 흰 여백만 살짝 트리밍(콘텐츠 크롭 아님)

# 1x1 PNG placeholder (이미지 미발견 시)
EMBEDDED_PLACEHOLDER_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQV"
    "R4nGMAAQAABQABDQottAAAAABJRU5ErkJggg=="
)

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None


# ================= 계산 로직 =================
def compute_outputs_full_with_inputs(
    K=None, n=None, Vel=None, D=None,
    Q1=None, C=0.33, V=1.2,
    pd_choice=None, pd_text=None,
    k1=0.7, V1_theory=1.2
):
    def round_n(x, n): return round(x, n)

    # --- Q1/Q2/Q3 ---
    have_all = (K is not None and n is not None and Vel is not None and D is not None)
    have_Q1 = (Q1 is not None)
    Q2 = None
    if have_all:
        # Q2 = (D^2) * ((Vel/K)^(2/(-n)))  → 소수 2자리
        Q2 = round_n((D**2) * ((Vel / K) ** (2 / (-n))), 2)
    if not have_Q1:
        if have_all:
            Q3 = Q2
        else:
            raise ValueError("입력오류: Q1이 비어있을 때는 K, n, Vel, D를 모두 입력해야 합니다.")
    else:
        # Q1이 있고 모든 파라미터도 있으면 Q3 = min(Q1, Q2); 아니면 Q3 = Q1
        Q3 = round_n(min(Q1, Q2), 2) if have_all else round_n(Q1, 2)

    # --- Pa ---
    if   Q3 < 0.125: Pa = 1
    elif Q3 < 0.5:   Pa = 2
    elif Q3 < 1.6:   Pa = 3
    elif Q3 < 5:     Pa = 4
    elif Q3 < 15:    Pa = 5
    else:            Pa = 6

    # --- pd 결정: ANFO 직접입력 > 라디오 선택 > Pa 기본 ---
    pd = None; pd_from_custom = False
    if pd_text:
        try:
            val = float(pd_text)
            if val > 0:
                pd = val
                pd_from_custom = True
        except:
            pass
    if pd is None and (pd_choice is not None):
        pd = float(pd_choice)
    if pd is None:
        if Pa in [1,2,3]: pd = 0.032
        elif Pa in [4,5]: pd = 0.050
        else: pd = 0.076
    pd = round_n(pd, 3)

    # --- [Pa=1,2 전용] pd 허용 규칙 ---
    pd_forced_msg = None
    user_pd_supplied = pd_from_custom or (pd_choice is not None)
    if Pa in (1, 2) and user_pd_supplied:
        if pd > 0.032:
            pd = 0.032
            pd_forced_msg = "폭약경이 적합하지 않아 0.032 m로 선택하였습니다."

    # --- ANFO 공식(0.815) ---
    def anfo_formula(pd_val):
        return (1000 * 0.815 * 3.1415 * (pd_val ** 2)) / 4.0, 1.0, 0.1

    # --- W1/h1/nu 테이블 ---
    tol = 1e-9
    if Pa == 1:
        W1, h1, nu = 0.12, 0.2, 0.5
    elif Pa == 2:
        W1, h1, nu = 0.25, 0.295, 0.5
    elif Pa == 3:
        if pd_from_custom:
            W1, h1, nu = anfo_formula(pd)
        elif abs(pd - 0.032) < tol:
            W1, h1, nu = 0.25, 0.295, 0.5
        elif abs(pd - 0.050) < tol:
            W1, h1, nu = 1.0, 0.420, 0.5
        else:
            W1, h1, nu = 0.25, 0.295, 0.5
    elif Pa == 4:
        if pd_from_custom:
            W1, h1, nu = anfo_formula(pd)
        elif abs(pd - 0.032) < tol:
            W1, h1, nu = 0.25, 0.295, 0.5
        elif abs(pd - 0.050) < tol:
            W1, h1, nu = 1.0, 0.42, 0.5
        elif abs(pd - 0.065) < tol:
            W1, h1, nu = 2.0, 0.52, 0.5
        else:
            W1, h1, nu = 1.0, 0.42, 0.5
    elif Pa == 5:
        if pd_from_custom:
            W1, h1, nu = anfo_formula(pd)
        elif abs(pd - 0.032) < tol:
            W1, h1, nu = 0.25, 0.295, 0.5
        elif abs(pd - 0.050) < tol:
            W1, h1, nu = 1.0, 0.42, 0.5
        elif abs(pd - 0.065) < tol:
            W1, h1, nu = 2.0, 0.52, 0.5
        else:
            W1, h1, nu = 1.0, 0.42, 0.5
    elif Pa == 6:
        if abs(pd - 0.065) < tol:
            W1, h1, nu = 2.0, 0.52, 1.0
        elif abs(pd - 0.050) < tol:
            W1, h1, nu = 1.0, 0.42, 0.5
        elif abs(pd - 0.032) < tol:
            W1, h1, nu = 0.25, 0.295, 0.5
        elif pd_from_custom:
            W1, h1, nu = anfo_formula(pd)
        else:
            W1, h1, nu = anfo_formula(pd)
    else:
        W1, h1, nu = 1.0, 0.42, 0.5

    # --- Q4/Q/h  (ANFO 선택 규칙: Q=Q3, h=h1*(Q3/W1) if Q>=0.5, else 기본 경로) ---
    if pd_from_custom:
        Q = float(Q3)  # ANFO 선택 시 Q는 Q3
        if Q >= 0.5:
            Q4 = int((Q / W1) * 2.0) if W1 <= 2.0 else int(Q)
            h  = h1 * (Q3 / W1)
        else:
            # Q<0.5 → ANFO 특수처리 해제(기본 경로)
            Q4 = int((Q3 / W1) * 2.0) if W1 <= 2.0 else int(Q3)
            Q  = (Q4 / 2.0) * W1 if W1 <= 2.0 else float(Q4)
            h  = 0.95 * h1 * Q / W1
    else:
        # 기본(비-ANFO) 경로
        Q4 = int((Q3 / W1) * 2.0) if W1 <= 2.0 else int(Q3)
        Q  = (Q4 / 2.0) * W1 if W1 <= 2.0 else float(Q4)
        h  = 0.95 * h1 * Q / W1

    # --- B,S 산정 ---
    denom = C * V1_theory * (0.7*h + 0.77*(Q**(1/3)) + 10*pd)
    if denom <= 0:
        raise ValueError("분모가 0 이하입니다.")
    B1_raw = 0.94 * math.sqrt(Q / denom)
    S1_raw = V1_theory * B1_raw

    if abs(V - 1.2) < 1e-12:
        B, S = round_n(B1_raw, 2), round_n(S1_raw, 2)
    else:
        # 면적 보존 기반 보정
        B_corr = math.sqrt((B1_raw * S1_raw) / V)
        S_corr = V * B_corr
        B, S = round_n(B_corr, 2), round_n(S_corr, 2)

    # --- T, H, K_step, c1 ---
    T = round_n((k1 * (pd ** -0.25) if Pa == 1 else k1 * (pd ** -0.18)) * math.sqrt(B * S), 2)
    H = round_n(T + h, 2)
    K_step = round_n(H - 0.2 * B, 2)
    denom_c1 = B * S * K_step
    c1_val = 0.0 if denom_c1 == 0 else Q / denom_c1
    c1 = round_n(c1_val, 2)

    return {
        "B": B, "S": S, "T": T, "h": H - T, "H": H, "Q": Q, "c1": c1, "K_step": K_step,
        "Q1": Q1, "Q2": Q2, "Q3": Q3, "Pa": Pa, "pd": pd, "W1": W1, "h1": h1, "nu": nu, "Q4": Q4,
        "B1": round(B1_raw,4), "S1": round(S1_raw,4),
        "_pd_from_custom": pd_from_custom,
        "_pd_forced_msg": pd_forced_msg,
    }


# ================= GUI =================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # 기준(UI 설계) 크기
        self.base_w, self.base_h = 1280, 900

        # 모니터 크기 측정 (작은 여유공간 확보)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        safe_w, safe_h = sw - 40, sh - 80

        # 스케일 = 1.0(그대로) 또는 축소 비율
        self.ui_scale = min(1.0, safe_w / self.base_w, safe_h / self.base_h)
        def S(x):  # 정수 스케일러
            return max(1, int(round(x * self.ui_scale)))
        self.S = S  # 다른 곳에서도 사용

        # 창 크기 적용
        self.title("발파설계앱 (v.1)")
        self.geometry(f"{S(self.base_w)}x{S(self.base_h)}")
        self.resizable(False, False)

        # ===== 글꼴 (스케일 반영) =====
        base_font = tkfont.nametofont("TkDefaultFont")
        fam = base_font.cget("family")
        base_size = int(base_font.cget("size"))
        inc = max(int(round(base_size*1.3)), base_size+2)

        self.font_input       = tkfont.Font(family=fam, size=max(8, int(round(inc           * self.ui_scale))))
        self.font_input_bold  = tkfont.Font(family=fam, size=max(8, int(round(inc           * self.ui_scale))), weight="bold")
        self.font_btn         = tkfont.Font(family=fam, size=max(8, int(round(inc           * self.ui_scale))))
        self.font_res_title   = tkfont.Font(family=fam, size=max(8, int(round(14            * self.ui_scale))), weight="bold")
        self.font_res_label   = tkfont.Font(family=fam, size=max(8, int(round(12            * self.ui_scale))), weight="bold")
        self.font_res_value   = tkfont.Font(family=fam, size=max(8, int(round(12            * self.ui_scale))))
        self.hint_color       = "#b91c1c"

        self.last_result = None
        self.last_pattern_path = None

        frm = ttk.Frame(self, padding=S(10)); frm.pack(fill="both", expand=True)

        # ===== 입력 =====
        self.entries = {}; row = 0
        def add_input_row(label_text, key, default, hint):
            nonlocal row
            ttk.Label(frm, text=label_text, width=22, font=self.font_input_bold)\
                .grid(row=row, column=0, sticky="e", padx=S(4), pady=S(3))
            e = ttk.Entry(frm, width=18, font=self.font_input)
            e.grid(row=row, column=1, sticky="w", padx=S(4), pady=S(3))
            if default != "": e.insert(0, str(default))
            self.entries[key] = e    
            ttk.Label(frm, text=hint, font=self.font_input, foreground=self.hint_color)\
                .grid(row=row, column=2, sticky="w", padx=S(6), pady=S(3))
            row += 1

        add_input_row("공당(지발당)장약량(kg)", "Q1", "", "입력하지 않으면 이격거리에 따라 산출")
        add_input_row("K값(진동추정식)", "K", 200.0, "시험발파추정식 변경 가능")
        add_input_row("n값(진동추정식)", "n", -1.60, "시험발파추정식 변경 가능")
        add_input_row("허용진동기준치(cm/sec)", "Vel", 0.30, "보안물건의 허용기준치 입력")
        add_input_row("보안물건과 거리(m)", "D", "", "진동을 고려하고 싶은 경우 입력 ")
        add_input_row("발파계수", "C", 0.33, "암질에 따라 풍화암 0.25 ~ 경암 0.5")
        add_input_row("공간격비율", "V", 1.2, "보통 1.0 ~ 1.25 범위 설정함")

        # 폭약직경 + ANFO
        ttk.Label(frm, text="폭약직경(m)", width=22, font=self.font_input_bold)\
            .grid(row=row, column=0, sticky="e", padx=S(4), pady=S(3))
        pd_frame = ttk.Frame(frm); pd_frame.grid(row=row, column=1, sticky="w", padx=S(4), pady=S(3))
        self.pd_var = tk.StringVar(value="__none__")
        self.pd_radios = []
        for v in ["0.032","0.050","0.065"]:
            rb = tk.Radiobutton(pd_frame, text=v, value=v, variable=self.pd_var, font=self.font_input)
            rb.pack(side="left", padx=S(4)); self.pd_radios.append(rb)
        rb_custom = tk.Radiobutton(pd_frame, text="ANFO(입력(m))", value="custom", variable=self.pd_var, font=self.font_input)
        rb_custom.pack(side="left", padx=(S(8),S(4))); self.pd_radios.append(rb_custom)
        self.pd_entry = ttk.Entry(pd_frame, width=10, font=self.font_input, state="disabled")
        self.pd_entry.pack(side="left", padx=S(4))
        ttk.Label(frm, text="단위(m), 선택하지 않으면 자동선택", font=self.font_input, foreground=self.hint_color)\
            .grid(row=row, column=2, sticky="w", padx=S(6), pady=S(3))
        self.pd_var.trace_add("write", self._on_pd_change)
        for rb in self.pd_radios:
            try: rb.deselect()
            except: pass
        self.after(0, lambda: [rb.deselect() for rb in self.pd_radios])
        self.pd_var.set("__none__")
        row += 1

        # 목적 선택(k1)
        ttk.Label(frm, text="목적선택", width=22, font=self.font_input_bold)\
            .grid(row=row, column=0, sticky="e", padx=S(4), pady=S(3))
        k1_frame = ttk.Frame(frm); k1_frame.grid(row=row, column=1, sticky="w", padx=S(4), pady=S(3))
        self.k1_var = tk.StringVar(value="0.7")
        for txt,val in [("비산제어 (0.7)","0.7"),("파쇄도개선 (0.55)","0.55"),("광산·채석장 (0.5)","0.5")]:
            tk.Radiobutton(k1_frame, text=txt, value=val, variable=self.k1_var, font=self.font_input)\
                .pack(side="left", padx=S(8))
        row += 1

        # 버튼들
        btn = ttk.Frame(frm); btn.grid(row=row, column=0, columnspan=3, pady=S(8))
        tk.Button(btn, text="계산", width=10, relief="raised",
                  bg="#cfe8ff", activebackground="#a8d1ff", bd=2,
                  font=self.font_btn, command=self.calculate).pack(side="left", padx=S(10))
        tk.Button(btn, text="결과출력", width=10, relief="raised",
                  bg="#ef4444", activebackground="#dc2626", fg="white", bd=2,
                  font=self.font_btn, command=self.export_pdf).pack(side="left", padx=S(10))
        tk.Button(btn, text="인쇄", width=10, relief="raised",
                  bg="#22c55e", activebackground="#16a34a", fg="white", bd=2,
                  font=self.font_btn, command=self.print_result).pack(side="left", padx=S(10))
        row += 1

        # ===== 표시 영역 =====
        display = ttk.Frame(frm); display.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=(S(4),0))
        frm.grid_rowconfigure(row, weight=1)

        self.values_panel = tk.Canvas(display, width=self.S(300), height=self.S(620), bg="#f8fafc",
                                      highlightthickness=1, highlightbackground="#e5e7eb")
        self.values_panel.pack(side="left", padx=(0,self.S(10)), pady=self.S(6), fill="y")
        self.image_canvas = tk.Canvas(display, width=self.S(840), height=self.S(620), bg="white",
                                      highlightthickness=1, highlightbackground="#ddd")
        self.image_canvas.pack(side="left", padx=0, pady=self.S(6))

        self._pil_image = None
        self._tk_image  = None

    # ---------- 이미지 유틸 ----------
    def _mm_to_px(self, mm: float) -> int:
        try:
            dpi_px_per_in = self.winfo_fpixels('1i')
        except Exception:
            dpi_px_per_in = 96.0
        return int(dpi_px_per_in * (mm / 25.4))

    def _autocrop_whitespace(self, img, thr=245):
        if Image is None: return img
        gray = img.convert("L")
        bw = gray.point(lambda p: 255 if p < thr else 0, mode="1")
        bbox = bw.getbbox()
        return img.crop(bbox) if bbox else img

    def _select_pattern_image_by_ratio(self, ratio: float):
        search_dirs = []
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            # 발파프로그램 폴더 우선 검색
            balpa_prog = os.path.join(base_dir, "발파프로그램")
            if os.path.isdir(balpa_prog): search_dirs.append(balpa_prog)
            # patterns 폴더
            preferred = os.path.join(base_dir, "patterns")
            if os.path.isdir(preferred): search_dirs.append(preferred)
            # 기본 디렉토리
            search_dirs.append(base_dir)
        except Exception: pass
        bundle = getattr(sys, "_MEIPASS", None)
        if bundle and os.path.isdir(bundle): search_dirs.append(bundle)
        search_dirs.append(os.getcwd())
        # 현재 디렉토리의 발파프로그램 폴더도 검색
        cwd_balpa = os.path.join(os.getcwd(), "발파프로그램")
        if os.path.isdir(cwd_balpa) and cwd_balpa not in search_dirs:
            search_dirs.insert(0, cwd_balpa)

        if   ratio <= 0.25:  idx = 1
        elif ratio <= 0.375: idx = 2
        elif ratio <= 0.625: idx = 3
        elif ratio <= 0.75:  idx = 4
        else:                idx = 5

        for d in search_dirs:
            try:
                for nm in os.listdir(d):
                    low = nm.lower()
                    if low.endswith((".jpg",".jpeg",".png",".bmp",".gif")) and (
                        f"발파패턴{idx}" in nm or f"pattern{idx}" in low or f"blast{idx}" in low
                    ):
                        return os.path.join(d, nm)
            except Exception: pass
        return None

    def _load_and_show_image(self, path):
        if Image is None or ImageTk is None:
            messagebox.showwarning("Pillow 필요", "이미지 표시를 위해 Pillow가 필요합니다. 'pip install pillow' 후 다시 실행하세요.")
            return
        try:
            img = Image.open(path)
        except Exception as e:
            messagebox.showerror("이미지 오류", f"이미지를 열 수 없습니다:\n{e}")
            return
        self._show_img_on_canvas(img)

    def _load_embedded_placeholder(self):
        if Image is None or ImageTk is None:
            messagebox.showwarning("Pillow 필요", "이미지 표시를 위해 Pillow가 필요합니다.")
            return
        try:
            raw = base64.b64decode(EMBEDDED_PLACEHOLDER_B64)
            from PIL import Image as PILImage
            img = PILImage.open(io.BytesIO(raw))
        except Exception as e:
            messagebox.showerror("이미지 오류", f"내장 이미지를 열 수 없습니다: {e}")
            return
        self._show_img_on_canvas(img)

    def _show_img_on_canvas(self, img):
        if AUTO_CROP and Image is not None:
            try:
                img = self._autocrop_whitespace(img, thr=245)
            except Exception:
                pass

        margin_top    = self._mm_to_px(15.0)
        margin_bottom = self._mm_to_px(15.0)
        margin_left   = self._mm_to_px(30.0)
        margin_right  = self._mm_to_px(30.0)

        self.image_canvas.update_idletasks()
        cw = self.image_canvas.winfo_width()
        ch = self.image_canvas.winfo_height()

        try:
            hl = int(float(self.image_canvas.cget('highlightthickness')))
        except Exception:
            hl = 0
        try:
            bd = int(float(self.image_canvas.cget('bd')))  # borderwidth(alias)
        except Exception:
            try:
                bd = int(float(self.image_canvas.cget('borderwidth')))
            except Exception:
                bd = 0

        content_w = max(1, cw - 2*(hl + bd))
        content_h = max(1, ch - 2*(hl + bd))
        safety = 2

        target_w = max(1, content_w - (margin_left + margin_right) - safety)
        target_h = max(1, content_h - (margin_top + margin_bottom) - safety)

        scale = min(target_w / img.width, target_h / img.height)
        new_w = max(1, int(round(img.width  * scale)))
        new_h = max(1, int(round(img.height * scale)))

        try:
            from PIL import Image as PILImage
            resample = getattr(PILImage, "LANCZOS", getattr(PILImage, "ANTIALIAS", 1))
        except Exception:
            resample = 1
        if (new_w, new_h) != (img.width, img.height):
            img = img.resize((new_w, new_h), resample)

        self._pil_image = img
        self._tk_image  = ImageTk.PhotoImage(img)
        self.image_canvas.delete("all")

        x0 = hl + bd + margin_left + (target_w - new_w)//2
        y0 = hl + bd + margin_top
        self.image_canvas.create_image(x0, y0, image=self._tk_image, anchor="nw")

    # ---------- 입력/계산 ----------
    @staticmethod
    def _parse_float_or_none(s: str):
        s = s.strip()
        if s == "": return None
        return float(s)

    def _on_pd_change(self, *args):
        v = self.pd_var.get()
        if v == "custom":
            self.pd_entry.config(state="normal"); self.pd_entry.focus_set()
        else:
            self.pd_entry.config(state="disabled"); self.pd_entry.delete(0, tk.END)

    def calculate(self):
        try:
            Q1  = self._parse_float_or_none(self.entries["Q1"].get())
            K   = self._parse_float_or_none(self.entries["K"].get())
            n   = self._parse_float_or_none(self.entries["n"].get())
            Vel = self._parse_float_or_none(self.entries["Vel"].get())
            D   = self._parse_float_or_none(self.entries["D"].get())
            C   = self._parse_float_or_none(self.entries["C"].get())
            V   = self._parse_float_or_none(self.entries["V"].get())
        except ValueError:
            messagebox.showerror("입력 오류", "숫자 형식이 올바르지 않습니다."); return
        if C is None: C = 0.33
        if V is None: V = 1.2

        # pd
        pd_choice = None; pd_text = None
        val = self.pd_var.get()
        if val == "custom":
            raw = self.pd_entry.get().strip()
            if raw == "":
                messagebox.showerror("입력 오류", "ANFO(직접입력)를 선택하셨습니다. 천공경(m)을 입력하세요.")
                return
            try:
                _pd = float(raw)
                if _pd <= 0: raise ValueError
                pd_text = f"{_pd}"
            except Exception:
                messagebox.showerror("입력 오류", "ANFO 천공경은 양의 실수(m)로 입력하세요.")
                return
        elif val in ("0.032","0.050","0.065"):
            pd_choice = float(val)

        # k1
        k1_str = getattr(self, "k1_var", tk.StringVar(value="0.7")).get().strip()
        try: k1_val = float(k1_str) if k1_str else 0.7
        except: k1_val = 0.7

        try:
            res = compute_outputs_full_with_inputs(
                K=K, n=n, Vel=Vel, D=D, Q1=Q1, C=C, V=V,
                pd_choice=pd_choice, pd_text=pd_text, k1=k1_val
            )
        except Exception as e:
            messagebox.showerror("계산 오류", str(e)); return

        self.last_result = res

        # pd 강제 변경 안내
        msg = res.get("_pd_forced_msg")
        if msg:
            messagebox.showinfo("폭약경 조정", msg)

        self._update_values_panel(res)

        H = float(res["H"]); h_val = float(res["h"])
        ratio = 0.0 if H == 0 else (h_val / H)
        path = self._select_pattern_image_by_ratio(ratio)
        self.last_pattern_path = path
        if path: self._load_and_show_image(path)
        else:    self._load_embedded_placeholder()

    def _update_values_panel(self, res: dict):
        c = self.values_panel; c.delete("all")
        S = self.S
        x0, y0 = S(18), S(16)
        title = {
            1:"미진동발파패턴",2:"정밀진동제어발파",3:"소규모진동제어발파",
            4:"중규모진동제어발파",5:"일반발파",6:"대규모발파"
        }.get(res.get("Pa",5),"일반발파")
        c.create_text(x0, y0, text=title, anchor="nw", font=self.font_res_title)

        y = y0 + S(36)
        lines = [
            ("B(저항선)", f"{res['B']:.2f} m"),
            ("공간격",    f"{res['S']:.2f} m"),
            ("전색장",    f"{res['T']:.2f} m"),
            ("장약장",    f"{res['h']:.2f} m"),
            ("천공장",    f"{res['H']:.2f} m"),
            ("계단높이",   f"{res['K_step']:.2f} m"),
            ("장약량/공",  f"{res['Q']} kg"),
            ("비장약량",   f"{res['c1']} kg/m³"),
            ("폭약경",     f"{res['pd']} m"),
        ]
        for label,val in lines:
            lbl = c.create_text(x0, y, text=f"{label} : ", anchor="nw", font=self.font_res_label)
            bx = c.bbox(lbl); x_val = (bx[2] + S(4)) if bx else (x0 + S(120))
            c.create_text(x_val, y, text=val, anchor="nw", font=self.font_res_value)
            y += S(32)

    # ---------- PDF 내보내기 ----------
    def export_pdf(self):
        # 1) 결과가 없으면 계산 먼저 시도
        if not self.last_result:
            self.calculate()
            if not self.last_result:
                return

        # 2) 경로 선택
        save_path = filedialog.asksaveasfilename(
            title="결과 PDF 저장",
            defaultextension=".pdf",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        )
        if not save_path:
            return

        # 3) reportlab 로드
        try:
            from reportlab.pdfgen import canvas as pdfcanvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.utils import ImageReader
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
        except Exception:
            messagebox.showerror("모듈 필요", "PDF 저장을 위해 reportlab이 필요합니다.\n명령프롬프트에서:\n  pip install reportlab")
            return

        # 4) 한글 폰트 등록 시도
        font_name = None
        candidates = []
        win = os.environ.get("WINDIR")
        if win:
            candidates += [
                os.path.join(win, "Fonts", "malgun.ttf"),
                os.path.join(win, "Fonts", "NanumGothic.ttf"),
            ]
        candidates += [
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        ]
        for p in candidates:
            if os.path.isfile(p):
                try:
                    pdfmetrics.registerFont(TTFont("KOR", p))
                    font_name = "KOR"
                    break
                except Exception:
                    continue

        # 5) 페이지/여백(mm)
        PAGE_W, PAGE_H = A4  # 595x842 pt
        def mm(x): return x * 72.0 / 25.4
        margin_l, margin_r = mm(30), mm(20)    # 왼쪽 30 mm
        margin_t, margin_b = mm(30), mm(15)    # 위쪽 30 mm

        # 6) 캔버스 만들기 & 폰트
        c = pdfcanvas.Canvas(save_path, pagesize=A4)
        def setfont(size):
            try:
                c.setFont(font_name if font_name else "Helvetica", size)
            except Exception:
                c.setFont("Helvetica", size)

        # 7) 타이틀 (가운데 정렬)
        setfont(18)
        c.drawCentredString(PAGE_W / 2.0, PAGE_H - margin_t, "스마트스템 발파설계")

        # 출력날짜 (우측 정렬)
        setfont(10)
        output_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        c.drawRightString(PAGE_W - margin_r, PAGE_H - margin_t - mm(8), f"출력날짜: {output_date}")
        y = PAGE_H - margin_t - mm(16)

        # Pa 라벨
        pa_title = {
            1:"미진동발파패턴",2:"정밀진동제어발파",3:"소규모진동제어발파",
            4:"중규모진동제어발파",5:"일반발파",6:"대규모발파"
        }.get(self.last_result.get("Pa",5),"일반발파")
        setfont(12)
        c.drawString(margin_l, y, f"발파공법 : {pa_title}")
        y -= mm(8)

        # 8) 좌측 정보 블록
        lines = [
            ("저항선(B)", f"{self.last_result['B']:.2f} m"),
            ("공간격(S)", f"{self.last_result['S']:.2f} m"),
            ("전색장(T)", f"{self.last_result['T']:.2f} m"),
            ("장약장(h)", f"{self.last_result['h']:.2f} m"),
            ("천공장(H)", f"{self.last_result['H']:.2f} m"),
            ("계단높이(K_step)", f"{self.last_result['K_step']:.2f} m"),
            ("장약량/공(Q)", f"{self.last_result['Q']} kg"),
            ("비장약량(c1)", f"{self.last_result['c1']} kg/m³"),
            ("폭약경(pd)", f"{self.last_result['pd']} m"),
        ]
        setfont(11)
        col_w = (PAGE_W - margin_l - margin_r) * 0.50  # 왼쪽 영역 폭
        line_h = mm(7)
        y0 = y
        for lab, val in lines:
            c.drawString(margin_l, y, f"{lab} : {val}")
            y -= line_h

        # 9) 우측 패턴 이미지 (← 왼쪽 10mm 이동, ↑ 위쪽 10mm 이동)
        img_path = self.last_pattern_path
        x_img_base = margin_l + col_w + mm(8)
        x_img = max(margin_l + mm(5), x_img_base - mm(10))  # 최소 여백 보호하며 왼쪽으로 10mm
        img_area_w = PAGE_W - margin_r - x_img
        img_area_h = y0 - margin_b
        if img_area_w > 0 and img_area_h > 0:
            try:
                if img_path and os.path.isfile(img_path):
                    ir = ImageReader(img_path)
                    iw, ih = ir.getSize()
                    scale = min(img_area_w/iw, img_area_h/ih)
                    new_w, new_h = iw*scale, ih*scale

                    # top 기준을 y0에서 10mm 올림, 상단여백(30mm) 아래 2mm 버퍼 확보
                    y_top_base   = y0
                    y_top_target = min(y_top_base + mm(10), PAGE_H - margin_t - mm(2))
                    y_img = max(margin_b, y_top_target - new_h)  # bottom y

                    c.drawImage(ir, x_img, y_img,
                                width=new_w, height=new_h,
                                preserveAspectRatio=True, mask='auto')
                else:
                    c.setStrokeColorRGB(0.8,0.8,0.8)
                    c.rect(x_img, margin_b, img_area_w, img_area_h, stroke=1, fill=0)
                    setfont(10)
                    c.drawString(x_img+mm(5), margin_b+img_area_h/2, "패턴 이미지 없음")
            except Exception as e:
                setfont(10)
                c.drawString(x_img, margin_b+img_area_h/2, f"이미지 오류: {e}")

        # 10) 폰트 안내(옵션)
        if font_name is None:
            c.setFillColorRGB(1,0,0)
            setfont(9)
            c.drawString(margin_l, margin_b/2,
                         "참고: 시스템 한글 폰트를 찾지 못했습니다. 글자가 깨지면 malgun.ttf 또는 NanumGothic.ttf를 설치해 주세요.")

        # 11) 저장
        c.showPage()
        c.save()
        messagebox.showinfo("완료", f"PDF 저장 완료:\n{save_path}")

    # ---------- 인쇄 기능 ----------
    def print_result(self):
        # 1) 결과가 없으면 계산 먼저 시도
        if not self.last_result:
            self.calculate()
            if not self.last_result:
                return

        # 2) reportlab 로드
        try:
            from reportlab.pdfgen import canvas as pdfcanvas
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.utils import ImageReader
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
        except Exception:
            messagebox.showerror("모듈 필요", "인쇄를 위해 reportlab이 필요합니다.\n명령프롬프트에서:\n  pip install reportlab")
            return

        # 3) 임시 PDF 파일 생성
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, "발파설계결과_print.pdf")

        # 4) 한글 폰트 등록 시도
        font_name = None
        candidates = []
        win = os.environ.get("WINDIR")
        if win:
            candidates += [
                os.path.join(win, "Fonts", "malgun.ttf"),
                os.path.join(win, "Fonts", "NanumGothic.ttf"),
            ]
        candidates += [
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        ]
        for p in candidates:
            if os.path.isfile(p):
                try:
                    pdfmetrics.registerFont(TTFont("KOR", p))
                    font_name = "KOR"
                    break
                except Exception:
                    continue

        # 5) 페이지/여백(mm)
        PAGE_W, PAGE_H = A4
        def mm(x): return x * 72.0 / 25.4
        margin_l, margin_r = mm(30), mm(20)
        margin_t, margin_b = mm(30), mm(15)

        # 6) 캔버스 만들기 & 폰트
        c = pdfcanvas.Canvas(temp_path, pagesize=A4)
        def setfont(size):
            try:
                c.setFont(font_name if font_name else "Helvetica", size)
            except Exception:
                c.setFont("Helvetica", size)

        # 7) 타이틀 (가운데 정렬)
        setfont(18)
        c.drawCentredString(PAGE_W / 2.0, PAGE_H - margin_t, "스마트스템 발파설계")

        # 출력날짜 (우측 정렬)
        setfont(10)
        output_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        c.drawRightString(PAGE_W - margin_r, PAGE_H - margin_t - mm(8), f"출력날짜: {output_date}")
        y = PAGE_H - margin_t - mm(16)

        # Pa 라벨
        pa_title = {
            1:"미진동발파패턴",2:"정밀진동제어발파",3:"소규모진동제어발파",
            4:"중규모진동제어발파",5:"일반발파",6:"대규모발파"
        }.get(self.last_result.get("Pa",5),"일반발파")
        setfont(12)
        c.drawString(margin_l, y, f"발파공법 : {pa_title}")
        y -= mm(8)

        # 8) 좌측 정보 블록
        lines = [
            ("저항선(B)", f"{self.last_result['B']:.2f} m"),
            ("공간격(S)", f"{self.last_result['S']:.2f} m"),
            ("전색장(T)", f"{self.last_result['T']:.2f} m"),
            ("장약장(h)", f"{self.last_result['h']:.2f} m"),
            ("천공장(H)", f"{self.last_result['H']:.2f} m"),
            ("계단높이(K_step)", f"{self.last_result['K_step']:.2f} m"),
            ("장약량/공(Q)", f"{self.last_result['Q']} kg"),
            ("비장약량(c1)", f"{self.last_result['c1']} kg/m³"),
            ("폭약경(pd)", f"{self.last_result['pd']} m"),
        ]
        setfont(11)
        col_w = (PAGE_W - margin_l - margin_r) * 0.50
        line_h = mm(7)
        y0 = y
        for lab, val in lines:
            c.drawString(margin_l, y, f"{lab} : {val}")
            y -= line_h

        # 9) 우측 패턴 이미지
        img_path = self.last_pattern_path
        x_img_base = margin_l + col_w + mm(8)
        x_img = max(margin_l + mm(5), x_img_base - mm(10))
        img_area_w = PAGE_W - margin_r - x_img
        img_area_h = y0 - margin_b
        if img_area_w > 0 and img_area_h > 0:
            try:
                if img_path and os.path.isfile(img_path):
                    # Pillow로 이미지를 RGB로 변환하여 검정색 문제 해결
                    from PIL import Image as PILImage
                    pil_img = PILImage.open(img_path)
                    if pil_img.mode in ('RGBA', 'LA', 'P'):
                        # 투명 배경을 흰색으로 변환
                        background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
                        if pil_img.mode == 'P':
                            pil_img = pil_img.convert('RGBA')
                        background.paste(pil_img, mask=pil_img.split()[-1] if pil_img.mode == 'RGBA' else None)
                        pil_img = background
                    elif pil_img.mode != 'RGB':
                        pil_img = pil_img.convert('RGB')

                    # 임시 파일로 저장 후 사용
                    temp_img_path = os.path.join(temp_dir, "temp_pattern_print.jpg")
                    pil_img.save(temp_img_path, 'JPEG', quality=95)

                    ir = ImageReader(temp_img_path)
                    iw, ih = ir.getSize()
                    scale = min(img_area_w/iw, img_area_h/ih)
                    new_w, new_h = iw*scale, ih*scale
                    y_top_base   = y0
                    y_top_target = min(y_top_base + mm(10), PAGE_H - margin_t - mm(2))
                    y_img = max(margin_b, y_top_target - new_h)
                    c.drawImage(ir, x_img, y_img,
                                width=new_w, height=new_h,
                                preserveAspectRatio=True)
                else:
                    c.setStrokeColorRGB(0.8,0.8,0.8)
                    c.rect(x_img, margin_b, img_area_w, img_area_h, stroke=1, fill=0)
                    setfont(10)
                    c.drawString(x_img+mm(5), margin_b+img_area_h/2, "패턴 이미지 없음")
            except Exception as e:
                setfont(10)
                c.drawString(x_img, margin_b+img_area_h/2, f"이미지 오류: {e}")

        # 10) 저장
        c.showPage()
        c.save()

        # 11) Windows 직접 인쇄 - 프린터 선택 후 인쇄
        self._print_with_dialog()

    def _print_with_dialog(self):
        """프린터 선택 대화상자를 표시하고 인쇄"""
        try:
            import win32print
            import win32ui
            import win32con
            from PIL import Image as PILImage, ImageDraw, ImageFont
        except ImportError:
            messagebox.showerror("모듈 필요", "인쇄를 위해 pywin32가 필요합니다.\n명령프롬프트에서:\n  pip install pywin32")
            return

        # 프린터 목록 가져오기
        printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        default_printer = win32print.GetDefaultPrinter()

        # 프린터 선택 대화상자
        select_win = tk.Toplevel(self)
        select_win.title("프린터 선택")
        select_win.geometry("350x150")
        select_win.resizable(False, False)
        select_win.transient(self)
        select_win.grab_set()

        # 중앙 배치
        select_win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 350) // 2
        y = self.winfo_y() + (self.winfo_height() - 150) // 2
        select_win.geometry(f"+{x}+{y}")

        ttk.Label(select_win, text="프린터 선택:", font=self.font_input_bold).pack(pady=(15, 5))

        printer_var = tk.StringVar(value=default_printer)
        printer_combo = ttk.Combobox(select_win, textvariable=printer_var, values=printers, width=40, state="readonly")
        printer_combo.pack(pady=5)

        def do_print():
            selected_printer = printer_var.get()
            select_win.destroy()
            self._do_actual_print(selected_printer)

        def cancel():
            select_win.destroy()

        btn_frame = ttk.Frame(select_win)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="인쇄", width=10, bg="#22c55e", fg="white", command=do_print).pack(side="left", padx=10)
        tk.Button(btn_frame, text="취소", width=10, command=cancel).pack(side="left", padx=10)

    def _do_actual_print(self, printer_name):
        """선택한 프린터로 실제 인쇄 수행"""
        try:
            import win32print
            import win32ui
            import win32con
            from PIL import Image as PILImage, ImageDraw, ImageFont

            # A4 크기 이미지 생성 (200 DPI)
            DPI = 200
            A4_WIDTH = int(8.27 * DPI)
            A4_HEIGHT = int(11.69 * DPI)

            img = PILImage.new('RGB', (A4_WIDTH, A4_HEIGHT), 'white')
            draw = ImageDraw.Draw(img)

            # 폰트 설정
            try:
                win_dir = os.environ.get("WINDIR", "C:\\Windows")
                font_path = os.path.join(win_dir, "Fonts", "malgun.ttf")
                font_title = ImageFont.truetype(font_path, 36)
                font_label = ImageFont.truetype(font_path, 24)
                font_value = ImageFont.truetype(font_path, 22)
            except:
                font_title = ImageFont.load_default()
                font_label = font_title
                font_value = font_title

            margin_l = int(30 * DPI / 25.4)
            margin_t = int(30 * DPI / 25.4)
            margin_r = int(20 * DPI / 25.4)

            # 제목
            title = "스마트스템 발파설계"
            try:
                title_bbox = draw.textbbox((0, 0), title, font=font_title)
                title_w = title_bbox[2] - title_bbox[0]
            except:
                title_w = 280
            draw.text(((A4_WIDTH - title_w) // 2, margin_t), title, fill='black', font=font_title)

            # 출력날짜 (우측 정렬)
            output_date = datetime.now().strftime("%Y-%m-%d %H:%M")
            date_text = f"출력날짜: {output_date}"
            try:
                font_date = ImageFont.truetype(font_path, 18)
                date_bbox = draw.textbbox((0, 0), date_text, font=font_date)
                date_w = date_bbox[2] - date_bbox[0]
            except:
                font_date = font_value
                date_w = 180
            draw.text((A4_WIDTH - margin_r - date_w, margin_t + 50), date_text, fill='#555555', font=font_date)

            # Pa 타이틀
            pa_title = {
                1:"미진동발파패턴",2:"정밀진동제어발파",3:"소규모진동제어발파",
                4:"중규모진동제어발파",5:"일반발파",6:"대규모발파"
            }.get(self.last_result.get("Pa",5),"일반발파")

            y = margin_t + 70
            draw.text((margin_l, y), f"발파공법 : {pa_title}", fill='black', font=font_label)
            y += 50

            # 결과 데이터
            lines = [
                ("저항선(B)", f"{self.last_result['B']:.2f} m"),
                ("공간격(S)", f"{self.last_result['S']:.2f} m"),
                ("전색장(T)", f"{self.last_result['T']:.2f} m"),
                ("장약장(h)", f"{self.last_result['h']:.2f} m"),
                ("천공장(H)", f"{self.last_result['H']:.2f} m"),
                ("계단높이(K_step)", f"{self.last_result['K_step']:.2f} m"),
                ("장약량/공(Q)", f"{self.last_result['Q']} kg"),
                ("비장약량(c1)", f"{self.last_result['c1']} kg/m³"),
                ("폭약경(pd)", f"{self.last_result['pd']} m"),
            ]

            for lab, val in lines:
                draw.text((margin_l, y), f"{lab} : {val}", fill='black', font=font_value)
                y += 40

            # 패턴 이미지 삽입
            if self.last_pattern_path and os.path.isfile(self.last_pattern_path):
                try:
                    pattern_img = PILImage.open(self.last_pattern_path)
                    if pattern_img.mode != 'RGB':
                        pattern_img = pattern_img.convert('RGB')
                    max_w = A4_WIDTH // 2 - 40
                    max_h = A4_HEIGHT - margin_t * 4
                    pattern_img.thumbnail((max_w, max_h), PILImage.LANCZOS)
                    x_pos = A4_WIDTH // 2 + 20
                    y_pos = margin_t + 70
                    img.paste(pattern_img, (x_pos, y_pos))
                except:
                    pass

            # 프린터 DC 생성 및 인쇄
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)

            hDC.StartDoc("스마트스템 발파설계")
            hDC.StartPage()

            # 프린터 해상도
            printer_width = hDC.GetDeviceCaps(win32con.HORZRES)
            printer_height = hDC.GetDeviceCaps(win32con.VERTRES)

            # 비율 유지
            scale = min(printer_width / A4_WIDTH, printer_height / A4_HEIGHT)
            new_w = int(A4_WIDTH * scale)
            new_h = int(A4_HEIGHT * scale)

            # 이미지를 BMP로 저장
            temp_bmp = os.path.join(tempfile.gettempdir(), "print_temp.bmp")
            img.save(temp_bmp, "BMP")

            # 비트맵 로드 및 그리기
            import win32gui
            hBitmap = win32gui.LoadImage(0, temp_bmp, win32con.IMAGE_BITMAP, 0, 0, win32con.LR_LOADFROMFILE)

            mem_dc = hDC.CreateCompatibleDC()
            dib = win32ui.CreateBitmapFromHandle(hBitmap)
            mem_dc.SelectObject(dib)
            hDC.StretchBlt((0, 0), (new_w, new_h), mem_dc, (0, 0), (A4_WIDTH, A4_HEIGHT), win32con.SRCCOPY)

            hDC.EndPage()
            hDC.EndDoc()
            mem_dc.DeleteDC()
            hDC.DeleteDC()
            win32gui.DeleteObject(hBitmap)

            messagebox.showinfo("인쇄 완료", f"'{printer_name}'(으)로 인쇄를 전송했습니다.")

        except Exception as e:
            messagebox.showerror("인쇄 오류", f"인쇄 중 오류가 발생했습니다:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
