# -*- coding: utf-8 -*-
"""
ปริ้นบาร์โค้ด / QR Code สำรองสำหรับฉลากขนาด (100 x 50 mm)
ยิง/พิมพ์เลข -> Enter -> ดูตัวอย่าง -> Enter ปริ้น / Ctrl+S บันทึกภาพ / Esc ยกเลิก
ต้องติดตั้ง: pip install pillow pywin32 qrcode
"""
import os
import re
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import qrcode
import win32print
import win32ui
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageWin
from qrcode.constants import ERROR_CORRECT_M

# ================= ตั้งค่า =================
APP_TITLE = "ปริ้นบาร์โค้ดสำรอง (100 x 50 mm)"
# ---- label ----
LABEL_LONG_MM = 100      # ด้านยาวของ label
LABEL_SHORT_MM = 50      # ด้านสั้นของ label
SIDE_MARGIN_MM = 4       # ขอบซ้าย/ขวา
TEXT_HEIGHT_MM = 6       # ขนาดตัวหนังสือใต้รหัส
MIN_TEXT_MM = 2.5        # ข้อความยาวจะย่อตัวหนังสือได้เล็กสุดเท่านี้
GAP_MM = 2               # ระยะห่างบาร์โค้ดกับตัวหนังสือ
# ---- บาร์โค้ด ----
BARCODE_HEIGHT_MM = 22   # ความสูงแท่งบาร์โค้ด
QUIET_MODULES = 10       # พื้นที่ว่างหัว-ท้ายบาร์โค้ด
MAX_MODULE_MM = 0.40     # ความกว้างแท่งเล็กสุดไม่เกินนี้
# ---- QR Code ----
QR_MAX_MM = 38           # ขนาด QR รวมขอบขาว
QR_BORDER = 4            # ขอบขาวรอบ QR (หน่วยเป็นช่อง) ตามมาตรฐาน
QR_MIN_MODULE_PX = 2     # จุด QR เล็กสุด (เล็กกว่านี้สแกนยาก)
# ---- อื่นๆ ----
SAVE_DPI = 600           # ความละเอียดไฟล์ภาพที่บันทึก
DEFAULT_DPI = 300        # ใช้ตอนไม่มีเครื่องพิมพ์
PREVIEW_WIDTH_PX = 600   # ความกว้างภาพตัวอย่างบนจอ
ROTATE = Image.Transpose.ROTATE_90   # ถ้าออกมากลับหัว เปลี่ยนเป็น ROTATE_270

KIND_NAME = {"barcode": "บาร์โค้ด", "qr": "QR Code"}

# ================= แก้แป้นภาษาไทย =================
_THAI_TO_EN = {}
for th, en in (
    ("ๅ/-ภถุึคตจ", "1234567890"),
    ('๐"ฎฑธํ๊ณฯญฤฆฏโฌ็๋ษศ()ฉฮฺ์?', "QWERTYUIOPASDFGHJKLZXCVBNM"),  # Shift
    ("ๆไำพะัีรนยฟหกดเ้่าสผปแอิืท", "QWERTYUIOPASDFGHJKLZXCVBNM"),   # ไม่กด Shift
):
    _THAI_TO_EN.update(zip(th, en))

# ================= ตรวจเลข S10 =================
S10_RE = re.compile(r"^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$")
S10_WEIGHTS = (8, 6, 4, 2, 3, 5, 9, 7)

def s10_check_digit(d8):
    c = 11 - sum(int(d) * w for d, w in zip(d8, S10_WEIGHTS)) % 11
    return 0 if c == 10 else 5 if c == 11 else c

def compact(s):
    return s.replace(" ", "").replace("-", "").upper()

def normalize(raw, kind):
    """คืนค่า (ข้อความที่จะใช้, แปลงจากแป้นไทยหรือไม่)"""
    text = raw.strip()
    converted = False
    if re.search(r"[\u0E00-\u0E7F]", text):
        conv = "".join(_THAI_TO_EN.get(c, c) for c in text)
        if S10_RE.match(compact(conv)):
            return compact(conv), True          # ยิงเลขพัสดุตอนแป้นเป็นไทย
        if kind == "qr":
            return text, False                  # QR ใส่ภาษาไทยได้ ใช้ตามที่พิมพ์
        text, converted = conv, True
    if S10_RE.match(compact(text)):
        return compact(text), converted
    return text, converted

def validate(code):
    m = S10_RE.match(code)
    if not m:
        return "other", "ไม่ใช่รูปแบบเลขพัสดุ (S10) — สร้างเป็นรหัสทั่วไป"
    expected = s10_check_digit(m.group(2))
    if int(m.group(3)) == expected:
        return "ok", "เลขพัสดุถูกต้อง"
    return "bad", f"เลขตรวจสอบไม่ถูก (หลักที่ 11 ควรเป็น {expected}) อาจพิมพ์หรือยิงเลขผิด"

# ================= Code 128 =================
CODE128 = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312",
    "132212", "221213", "221312", "231212", "112232", "122132", "122231", "113222",
    "123122", "123221", "223211", "221132", "221231", "213212", "223112", "312131",
    "311222", "321122", "321221", "312212", "322112", "322211", "212123", "212321",
    "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
    "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121",
    "313121", "211331", "231131", "213113", "213311", "213131", "311123", "311321",
    "331121", "312113", "312311", "332111", "314111", "221411", "431111", "111224",
    "111422", "121124", "121421", "141122", "141221", "112214", "112412", "122114",
    "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
    "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112",
    "421211", "212141", "214121", "412121", "111143", "111341", "131141", "114113",
    "114311", "411113", "411311", "113141", "114131", "311141", "411131", "211412",
    "211214", "211232", "2331112",
]

def code128_modules(text):
    for ch in text:
        if not 32 <= ord(ch) <= 126:
            raise ValueError(f"ตัวอักษร '{ch}' ใช้ในบาร์โค้ดไม่ได้ (ถ้าต้องการภาษาไทย ให้เลือก QR Code)")
    codes = [104] + [ord(c) - 32 for c in text]
    checksum = (104 + sum(i * v for i, v in enumerate(codes[1:], start=1))) % 103
    codes += [checksum, 106]
    modules, bar = [], True
    for w in "".join(CODE128[c] for c in codes):
        modules.extend([bar] * int(w))
        bar = not bar
    return modules

# ================= QR Code =================
def qr_matrix(text):
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_M, border=QR_BORDER)
    qr.add_data(text)
    try:
        qr.make(fit=True)
    except Exception:
        raise ValueError("ข้อความยาวเกินกว่าที่ QR Code จะรองรับได้")
    return qr.get_matrix()

def check_encodable(text, kind):
    if kind == "qr":
        qr_matrix(text)
    else:
        code128_modules(text)

# ================= วาด label =================
def mm2px(mm, dpi):
    return int(round(mm * dpi / 25.4))

def load_font(px, thai=False):
    names = (("tahomabd.ttf", "leelawdb.ttf", "tahoma.ttf") if thai
             else ("arialbd.ttf", "tahomabd.ttf", "arial.ttf"))
    for name in names:
        try:
            return ImageFont.truetype(name, px)
        except OSError:
            pass
    return ImageFont.load_default()

def fit_text(draw, text, max_w, dpi):
    """ย่อตัวหนังสือให้พอดีความกว้าง ถ้ายังไม่พอให้ตัดท้ายเป็น …"""
    thai = any(ord(c) > 126 for c in text)
    size, min_size = mm2px(TEXT_HEIGHT_MM, dpi), max(10, mm2px(MIN_TEXT_MM, dpi))
    while True:
        font = load_font(size, thai)
        if draw.textlength(text, font=font) <= max_w or size <= min_size:
            break
        size -= 2
    if draw.textlength(text, font=font) > max_w:
        while text and draw.textlength(text + "…", font=font) > max_w:
            text = text[:-1]
        text += "…"
    return text, font

def render_label(text, kind, dpi):
    W, H = mm2px(LABEL_LONG_MM, dpi), mm2px(LABEL_SHORT_MM, dpi)
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)
    avail = W - 2 * mm2px(SIDE_MARGIN_MM, dpi)

    shown, font = fit_text(draw, text, avail, dpi)
    l, t, r, b = draw.textbbox((0, 0), shown, font=font)
    tw, th = r - l, b - t

    if kind == "qr":
        matrix = qr_matrix(text)
        n = len(matrix)
        side = min(mm2px(QR_MAX_MM, dpi), H - th - mm2px(GAP_MM, dpi))
        m = side // n
        if m < QR_MIN_MODULE_PX:
            raise ValueError("ข้อความยาวเกินไป QR Code บน label นี้จะเล็กจนสแกนยาก")
        size = n * m
        x0, y0 = (W - size) // 2, (H - (size + th)) // 2
        for y, row in enumerate(matrix):
            for x, dark in enumerate(row):
                if dark:
                    draw.rectangle([x0 + x * m, y0 + y * m,
                                    x0 + (x + 1) * m - 1, y0 + (y + 1) * m - 1], fill=0)
        text_y = y0 + size          # ขอบขาวของ QR ทำหน้าที่เป็นระยะห่างอยู่แล้ว
    else:
        modules = code128_modules(text)
        max_mod = max(1, int(MAX_MODULE_MM * dpi / 25.4))
        m = min(avail // (len(modules) + 2 * QUIET_MODULES), max_mod)
        if m < 1:
            raise ValueError("ข้อความยาวเกินไป บาร์โค้ดใส่ใน label ไม่พอ")
        bc_w, bc_h = len(modules) * m, mm2px(BARCODE_HEIGHT_MM, dpi)
        gap = mm2px(GAP_MM, dpi)
        x0, y0 = (W - bc_w) // 2, (H - (bc_h + gap + th)) // 2
        i = 0
        while i < len(modules):
            if modules[i]:
                j = i
                while j < len(modules) and modules[j]:
                    j += 1
                draw.rectangle([x0 + i * m, y0, x0 + j * m - 1, y0 + bc_h - 1], fill=0)
                i = j
            else:
                i += 1
        text_y = y0 + bc_h + gap

    draw.text(((W - tw) // 2 - l, text_y - t), shown, font=font, fill=0)
    return img

def to_bw(img):
    return img.point(lambda p: 255 if p >= 128 else 0, mode="1")

def save_label_image(path, text, kind):
    img = to_bw(render_label(text, kind, SAVE_DPI))
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        img.save(path, "PDF", resolution=SAVE_DPI)
    elif ext in (".jpg", ".jpeg"):
        img.convert("L").save(path, "JPEG", quality=95, dpi=(SAVE_DPI, SAVE_DPI))
    else:
        img.save(path, "PNG", dpi=(SAVE_DPI, SAVE_DPI))

# ================= เครื่องพิมพ์ =================
LOGPIXELSX = 88
PHYSICALWIDTH, PHYSICALHEIGHT = 110, 111
PHYSICALOFFSETX, PHYSICALOFFSETY = 112, 113

ATTR_WORK_OFFLINE = 0x400
STATUS_PROBLEMS = {
    0x00000001: "เครื่องพิมพ์ถูกหยุดชั่วคราว (Paused)",
    0x00000002: "เครื่องพิมพ์เกิดข้อผิดพลาด",
    0x00000008: "กระดาษติด",
    0x00000010: "กระดาษหมด",
    0x00000080: "เครื่องพิมพ์ออฟไลน์",
    0x00001000: "เครื่องพิมพ์ไม่พร้อมใช้งาน",
    0x00100000: "เครื่องพิมพ์ต้องการให้แก้ไขที่ตัวเครื่อง",
    0x00400000: "ฝาเครื่องพิมพ์เปิดอยู่",
}

def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    return [p[2] for p in win32print.EnumPrinters(flags)]

def pick_printer(printers):
    for p in printers:
        if "zebra" in p.lower() or "zdesigner" in p.lower():
            return p
    try:
        d = win32print.GetDefaultPrinter()
        if d in printers:
            return d
    except Exception:
        pass
    return printers[0] if printers else ""

def check_printer(printer):
    """คืนค่า (รายการปัญหา, จำนวนงานค้างในคิว)"""
    h = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(h, 2)
    finally:
        win32print.ClosePrinter(h)
    problems = [t for bit, t in STATUS_PROBLEMS.items() if info["Status"] & bit]
    if info["Attributes"] & ATTR_WORK_OFFLINE:
        problems.insert(0, "เครื่องพิมพ์ไม่ได้เชื่อมต่อ (Windows ตั้งเป็นใช้งานแบบออฟไลน์)")
    return problems, info.get("cJobs", 0)

def get_printer_info(printer):
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer)
    try:
        dpi = dc.GetDeviceCaps(LOGPIXELSX)
        w = dc.GetDeviceCaps(PHYSICALWIDTH) / dpi * 25.4
        h = dc.GetDeviceCaps(PHYSICALHEIGHT) / dpi * 25.4
        return dpi, w, h
    finally:
        dc.DeleteDC()

def print_label(printer, label, doc_name):
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer)
    started = False
    try:
        pw, ph = dc.GetDeviceCaps(PHYSICALWIDTH), dc.GetDeviceCaps(PHYSICALHEIGHT)
        ox, oy = dc.GetDeviceCaps(PHYSICALOFFSETX), dc.GetDeviceCaps(PHYSICALOFFSETY)
        if pw < ph:                      # driver ตั้งกระดาษแนวตั้ง -> หมุนให้
            label = label.transpose(ROTATE)
        page = Image.new("L", (pw, ph), 255)
        page.paste(label, ((pw - label.width) // 2, (ph - label.height) // 2))
        page = to_bw(page)

        dc.StartDoc(doc_name)
        started = True
        dc.StartPage()
        ImageWin.Dib(page).draw(dc.GetHandleOutput(), (-ox, -oy, pw - ox, ph - oy))
        dc.EndPage()
        dc.EndDoc()
        started = False
    finally:
        if started:
            try:
                dc.AbortDoc()
            except Exception:
                pass
        dc.DeleteDC()

# ================= หน้าตัวอย่าง =================
class PreviewDialog(tk.Toplevel):
    """result = "print" / "save" / None"""
    def __init__(self, master, title, label_img, warnings, print_error):
        super().__init__(master)
        self.result = None
        self.can_print = print_error is None
        self.title(title)
        self.resizable(False, False)
        self.transient(master)

        w = PREVIEW_WIDTH_PX
        h = round(label_img.height * w / label_img.width)
        self.photo = ImageTk.PhotoImage(label_img.resize((w, h), Image.Resampling.LANCZOS))
        tk.Label(self, image=self.photo, bd=1, relief="solid").pack(padx=16, pady=(16, 4))
        tk.Label(self, text="ขนาดจริง 100 x 50 mm", fg="gray").pack()

        for text in warnings:
            tk.Label(self, text="⚠ " + text, fg="#d97706", font=("Tahoma", 11, "bold"),
                     wraplength=w, justify="left").pack(padx=16, pady=2, anchor="w")
        if print_error:
            tk.Label(self, text=f"✖ ปริ้นไม่ได้: {print_error}\nยังบันทึกเป็นภาพได้ (Ctrl+S)",
                     fg="red", font=("Tahoma", 11, "bold"),
                     wraplength=w, justify="left").pack(padx=16, pady=2, anchor="w")

        btns = ttk.Frame(self)
        btns.pack(pady=14)
        print_btn = ttk.Button(btns, text="ปริ้น (Enter)", command=self.do_print)
        print_btn.pack(side="left", padx=6, ipady=4, ipadx=10)
        if not self.can_print:
            print_btn.state(["disabled"])
        save_btn = ttk.Button(btns, text="บันทึกเป็นภาพ (Ctrl+S)", command=self.do_save)
        save_btn.pack(side="left", padx=6, ipady=4, ipadx=10)
        ttk.Button(btns, text="ยกเลิก (Esc)", command=self.destroy).pack(
            side="left", padx=6, ipady=4, ipadx=10)

        self.bind("<Return>", lambda e: self.do_print())
        self.bind("<KP_Enter>", lambda e: self.do_print())
        self.bind("<Escape>", lambda e: self.destroy())
        # Ctrl+S ใช้ได้ทั้งแป้นอังกฤษและแป้นไทย (เช็กจากรหัสปุ่ม S)
        self.bind("<Control-KeyPress>", lambda e: self.do_save() if e.keycode == 83 else None)

        self.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        self.grab_set()
        (print_btn if self.can_print else save_btn).focus_set()
        self.wait_window(self)

    def do_print(self):
        if not self.can_print:
            self.bell()
            return
        self.result = "print"
        self.destroy()

    def do_save(self):
        self.result = "save"
        self.destroy()

# ================= หน้าจอหลัก =================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(False, False)
        self.last_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="เครื่องพิมพ์:").grid(row=0, column=0, sticky="w")
        printers = list_printers()
        self.printer_var = tk.StringVar(value=pick_printer(printers))
        cb = ttk.Combobox(f, textvariable=self.printer_var, values=printers,
                          state="readonly", width=48)
        cb.grid(row=0, column=1, sticky="we", pady=3)
        cb.bind("<<ComboboxSelected>>", self.on_printer_change)
        self.printer_info = tk.Label(f, font=("Tahoma", 9), fg="gray")
        self.printer_info.grid(row=1, column=1, sticky="w")

        ttk.Label(f, text="รูปแบบ:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.kind_var = tk.StringVar(value="barcode")
        mf = ttk.Frame(f)
        mf.grid(row=2, column=1, sticky="w", pady=(10, 0))
        for value, text in (("barcode", "บาร์โค้ด (Code 128)"), ("qr", "QR Code")):
            ttk.Radiobutton(mf, text=text, value=value, variable=self.kind_var,
                            command=self.on_kind_change).pack(side="left", padx=(0, 16))

        ttk.Label(f, text="ข้อความ / เลข:").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.code_var = tk.StringVar()
        self.entry = tk.Entry(f, textvariable=self.code_var, font=("Consolas", 22), width=24)
        self.entry.grid(row=3, column=1, sticky="we", pady=(10, 0))
        self.entry.bind("<Return>", self.on_enter)
        self.entry.bind("<KP_Enter>", self.on_enter)
        self.code_var.trace_add("write", lambda *_: self.on_text_change())

        self.check_label = tk.Label(f, font=("Tahoma", 10), anchor="w")
        self.check_label.grid(row=4, column=1, sticky="w")

        ttk.Button(f, text="ดูตัวอย่าง (Enter)", command=self.on_enter).grid(
            row=5, column=1, sticky="we", pady=8, ipady=6)

        self.status = tk.Label(f, text="ยิงหรือพิมพ์เลข แล้วกด Enter",
                               font=("Tahoma", 11, "bold"), fg="gray",
                               wraplength=620, justify="left")
        self.status.grid(row=6, column=0, columnspan=2, sticky="w")

        self.on_printer_change()
        self.entry.focus_set()

    # ---------- เครื่องพิมพ์ ----------
    def on_printer_change(self, event=None):
        p = self.printer_var.get()
        if not p:
            self.printer_info.config(text="ไม่พบเครื่องพิมพ์ (ยังบันทึกเป็นภาพได้)", fg="red")
            return
        try:
            dpi, w, h = get_printer_info(p)
            short, long_ = sorted((w, h))
            ok = abs(short - LABEL_SHORT_MM) <= 5 and abs(long_ - LABEL_LONG_MM) <= 5
            text = f"ความละเอียด {dpi} dpi • กระดาษใน driver {w:.0f} x {h:.0f} mm"
            if not ok:
                text += "  ⚠ ควรตั้งเป็น 100 x 50 mm"
            self.printer_info.config(text=text, fg="gray" if ok else "#d97706")
        except Exception as e:
            self.printer_info.config(text=f"อ่านข้อมูลเครื่องพิมพ์ไม่ได้: {e}", fg="red")

    # ---------- ตรวจข้อความระหว่างพิมพ์ ----------
    def on_kind_change(self):
        self.on_text_change()
        self.entry.focus_set()

    def on_text_change(self):
        kind = self.kind_var.get()
        code, converted = normalize(self.code_var.get(), kind)
        if not code:
            self.check_label.config(text="")
            return
        try:
            check_encodable(code, kind)
        except ValueError as e:
            self.check_label.config(text=f"✖ {e}", fg="red")
            return
        k, msg = validate(code)
        icon = {"ok": "✔", "bad": "⚠", "other": "ℹ"}[k]
        color = {"ok": "green", "bad": "#d97706", "other": "gray"}[k]
        extra = f"   (แปลงจากแป้นไทย → {code})" if converted else ""
        self.check_label.config(text=f"{icon} {msg}{extra}", fg=color)

    # ---------- Enter -> ตัวอย่าง ----------
    def on_enter(self, event=None):
        kind = self.kind_var.get()
        code, _ = normalize(self.code_var.get(), kind)
        if not code:
            return "break"
        try:
            check_encodable(code, kind)
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return self.select_entry()

        # ตรวจเครื่องพิมพ์ (ถ้าไม่พร้อม ยังดูตัวอย่างและบันทึกภาพได้)
        printer = self.printer_var.get()
        dpi, jobs, print_error = DEFAULT_DPI, 0, None
        if not printer:
            print_error = "ไม่พบเครื่องพิมพ์"
        else:
            try:
                problems, jobs = check_printer(printer)
                dpi = get_printer_info(printer)[0]
                if problems:
                    print_error = " / ".join(problems)
            except Exception as e:
                print_error = f"ติดต่อเครื่องพิมพ์ไม่ได้ ({e})"

        try:
            label = render_label(code, kind, dpi)
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return self.select_entry()

        warnings = []
        k, msg = validate(code)
        if k == "bad":
            warnings.append(msg)
        if jobs and not print_error:
            warnings.append(f"มีงานค้างในคิวเครื่องพิมพ์ {jobs} งาน — "
                            "เครื่องพิมพ์อาจไม่ได้เชื่อมต่อ หรือกำลังพิมพ์งานอื่นอยู่")

        title = f"ตัวอย่าง {KIND_NAME[kind]} — {code}"
        action = PreviewDialog(self, title, label, warnings, print_error).result
        if action == "print":
            self.print_now(printer, label, code, kind)
        elif action == "save":
            self.save_now(code, kind)
        else:
            self.status.config(text="ยกเลิกแล้ว", fg="gray")
            self.select_entry()
        return "break"

    def print_now(self, printer, label, code, kind):
        try:
            print_label(printer, label, f"{KIND_NAME[kind]} {code}")
        except Exception as e:
            self.status.config(text=f"✖ ส่งงานไม่สำเร็จ: {code}", fg="red")
            messagebox.showerror(APP_TITLE, f"ส่งงานไปเครื่องพิมพ์ไม่สำเร็จ\n\n{e}")
            return self.select_entry()
        self.status.config(
            text=f"➜ ส่งงานไปเครื่องพิมพ์แล้ว: {code} ({datetime.now():%H:%M:%S}) "
                 "— ตรวจดูว่าสติกเกอร์ออกมาหรือไม่", fg="green")
        self.code_var.set("")
        self.entry.focus_set()

    def save_now(self, code, kind):
        safe = re.sub(r'[\\/:*?"<>|\s]+', "_", code).strip("._")[:60] or "label"
        path = filedialog.asksaveasfilename(
            parent=self, title="บันทึกเป็นภาพ", initialdir=self.last_dir,
            initialfile=f"{safe}_{kind}.png", defaultextension=".png",
            filetypes=[("PNG (แนะนำ)", "*.png"), ("JPEG", "*.jpg"), ("PDF", "*.pdf")])
        if not path:
            self.status.config(text="ยกเลิกการบันทึก", fg="gray")
            return self.select_entry()
        try:
            save_label_image(path, code, kind)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"บันทึกไฟล์ไม่สำเร็จ\n\n{e}")
            return self.select_entry()
        self.last_dir = os.path.dirname(path)
        self.status.config(text=f"✔ บันทึกภาพแล้ว: {path}", fg="green")
        self.code_var.set("")
        self.entry.focus_set()

    def select_entry(self):
        self.entry.focus_set()
        self.entry.select_range(0, "end")
        return "break"


if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()