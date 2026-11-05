# -*- coding: utf-8 -*-
"""
ปริ้นบาร์โค้ด / QR Code สำรอง — ปรับขนาดตามเครื่องพิมพ์ หรือเลือกขนาดเอง
ยิง/พิมพ์เลข -> Enter -> ดูตัวอย่าง -> Enter ปริ้น / Ctrl+S บันทึกภาพ / Esc ยกเลิก
ต้องติดตั้ง: pip install pillow pywin32 qrcode
"""
import os
import re
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import qrcode
import win32con
import win32gui
import win32print
import win32ui
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageWin
from qrcode.constants import ERROR_CORRECT_M

# ================= ตั้งค่า =================
APP_TITLE = "ปริ้นบาร์โค้ดสำรอง"

SIZE_AUTO = "ตามเครื่องพิมพ์ (อัตโนมัติ)"
SIZE_CUSTOM = "กำหนดเอง..."
PRESETS = [                       # (ชื่อ, กว้าง, ยาว) แบบเดียวกับที่ driver แสดง
    ('100 x 150 mm (ใบปะหน้า 4x6")', 100, 150),
    ("100 x 100 mm", 100, 100),
    ("100 x 75 mm", 100, 75),
    ("100 x 50 mm", 100, 50),
    ("75 x 50 mm", 75, 50),
    ("60 x 40 mm", 60, 40),
    ("58 x 40 mm", 58, 40),
    ("50 x 30 mm", 50, 30),
    ("50 x 25 mm", 50, 25),
    ("40 x 30 mm", 40, 30),
    ("40 x 20 mm", 40, 20),
]
FALLBACK_SIZE = (100, 50)         # ใช้ตอนไม่มีเครื่องพิมพ์ + เลือกอัตโนมัติ
CUSTOM_RANGE_MM = (15, 250)

# ---- สัดส่วน layout (เทียบกับด้านสั้นของ label) ----
MARGIN_RATIO = 0.05               # ขอบ
MIN_MARGIN_MM = 1.5
TEXT_RATIO = 0.13                 # ขนาดตัวหนังสือ
MIN_TEXT_MM, MAX_TEXT_MM = 2.2, 7
GAP_RATIO = 0.04                  # ระยะบาร์โค้ด-ตัวหนังสือ
# ---- บาร์โค้ด ----
BARCODE_MAX_H_MM = 30
BARCODE_MIN_H_MM = 5
QUIET_MODULES = 10
MAX_MODULE_MM = 0.40
MIN_GOOD_MODULE_MM = 0.19         # เล็กกว่านี้เตือนว่าสแกนยาก
# ---- QR ----
QR_MAX_MM = 60
QR_BORDER = 4
QR_MIN_MODULE_PX = 2
# ---- อื่นๆ ----
SAVE_DPI = 600
DEFAULT_DPI = 300
PREVIEW_MAX = (600, 380)
ROTATE = Image.Transpose.ROTATE_90   # ถ้าออกมากลับหัว เปลี่ยนเป็น ROTATE_270

KIND_NAME = {"barcode": "บาร์โค้ด", "qr": "QR Code"}

def fmt_size(s):
    return f"{round(s[0], 1):g} x {round(s[1], 1):g} mm"

def same_size(a, b, tol=1.5):
    return all(abs(x - y) <= tol for x, y in zip(sorted(a), sorted(b)))

# ================= แก้แป้นภาษาไทย =================
_THAI_TO_EN = {}
for th, en in (
    ("ๅ/-ภถุึคตจ", "1234567890"),
    ('๐"ฎฑธํ๊ณฯญฤฆฏโฌ็๋ษศ()ฉฮฺ์?', "QWERTYUIOPASDFGHJKLZXCVBNM"),
    ("ๆไำพะัีรนยฟหกดเ้่าสผปแอิืท", "QWERTYUIOPASDFGHJKLZXCVBNM"),
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
    text = raw.strip()
    converted = False
    if re.search(r"[\u0E00-\u0E7F]", text):
        conv = "".join(_THAI_TO_EN.get(c, c) for c in text)
        if S10_RE.match(compact(conv)):
            return compact(conv), True
        if kind == "qr":
            return text, False
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
    qr_matrix(text) if kind == "qr" else code128_modules(text)

# ================= วาด label (ปรับตามขนาด) =================
def mm2px(mm, dpi):
    return int(round(mm * dpi / 25.4))

def px2mm(px, dpi):
    return px * 25.4 / dpi

def load_font(px, thai=False):
    names = (("tahomabd.ttf", "leelawdb.ttf", "tahoma.ttf") if thai
             else ("arialbd.ttf", "tahomabd.ttf", "arial.ttf"))
    for name in names:
        try:
            return ImageFont.truetype(name, px)
        except OSError:
            pass
    try:
        return ImageFont.load_default(size=px)   # Pillow >= 10.1
    except TypeError:
        return ImageFont.load_default()

def fit_text(draw, text, max_w, size, min_size):
    thai = any(ord(c) > 126 for c in text)
    while True:
        font = load_font(size, thai)
        if draw.textlength(text, font=font) <= max_w or size <= min_size:
            break
        size -= 1
    if draw.textlength(text, font=font) > max_w:
        while text and draw.textlength(text + "…", font=font) > max_w:
            text = text[:-1]
        text += "…"
    return text, font

def render_label(text, kind, dpi, W, H, unp_x=0, unp_y=0):
    """W >= H เสมอ (วางแนวนอน) / unp_x, unp_y = ขอบที่พิมพ์ไม่ได้ (px)
    คืนค่า (ภาพ, รายการคำเตือน)"""
    img = Image.new("L", (W, H), 255)
    draw = ImageDraw.Draw(img)
    notes = []

    base = max(mm2px(MIN_MARGIN_MM, dpi), int(H * MARGIN_RATIO))
    mx, my = max(base, unp_x), max(base, unp_y)
    aw, ah = W - 2 * mx, H - 2 * my
    if aw <= 0 or ah <= 0:
        raise ValueError("label เล็กเกินไป")

    t_px = min(max(int(H * TEXT_RATIO), mm2px(MIN_TEXT_MM, dpi)), mm2px(MAX_TEXT_MM, dpi))
    shown, font = fit_text(draw, text, aw, t_px, max(8, mm2px(MIN_TEXT_MM, dpi)))
    l, t, r, b = draw.textbbox((0, 0), shown, font=font)
    tw, th = r - l, b - t
    gap = max(2, int(H * GAP_RATIO))

    if kind == "qr":
        matrix = qr_matrix(text)
        n = len(matrix)
        side = min(ah - th - gap, aw, mm2px(QR_MAX_MM, dpi))
        m = side // n
        if m < QR_MIN_MODULE_PX:
            raise ValueError("label ขนาดนี้เล็กเกินไปสำหรับ QR Code ของข้อความนี้")
        size = n * m
        x0, y0 = (W - size) // 2, my + (ah - (size + th + gap)) // 2
        for y, row in enumerate(matrix):
            for x, dark in enumerate(row):
                if dark:
                    draw.rectangle([x0 + x * m, y0 + y * m,
                                    x0 + (x + 1) * m - 1, y0 + (y + 1) * m - 1], fill=0)
        text_y = y0 + size + gap
    else:
        modules = code128_modules(text)
        max_mod = max(1, int(MAX_MODULE_MM * dpi / 25.4))
        m = min(aw // (len(modules) + 2 * QUIET_MODULES), max_mod)
        if m < 1:
            raise ValueError("ข้อความยาวเกินไปสำหรับ label ขนาดนี้")
        if px2mm(m, dpi) < MIN_GOOD_MODULE_MM:
            notes.append(f"แท่งบาร์โค้ดเล็กมาก ({px2mm(m, dpi):.2f} mm) อาจสแกนยาก "
                         "ลองใช้ label ที่กว้างขึ้น")
        bc_h = min(ah - th - gap, mm2px(BARCODE_MAX_H_MM, dpi))
        if bc_h < mm2px(BARCODE_MIN_H_MM, dpi):
            raise ValueError("label เตี้ยเกินไป บาร์โค้ดจะสั้นจนสแกนไม่ได้")
        bc_w = len(modules) * m
        x0, y0 = (W - bc_w) // 2, my + (ah - (bc_h + gap + th)) // 2
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
    return img, notes

def to_bw(img):
    return img.point(lambda p: 255 if p >= 128 else 0, mode="1")

def save_label_image(path, text, kind, size_mm):
    long_, short = max(size_mm), min(size_mm)
    img, _ = render_label(text, kind, SAVE_DPI, mm2px(long_, SAVE_DPI), mm2px(short, SAVE_DPI))
    if size_mm[0] < size_mm[1]:      # label แนวตั้ง -> หมุนกลับให้ตรงกับสติกเกอร์จริง
        img = img.transpose(ROTATE)
    img = to_bw(img)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        img.save(path, "PDF", resolution=SAVE_DPI)
    elif ext in (".jpg", ".jpeg"):
        img.convert("L").save(path, "JPEG", quality=95, dpi=(SAVE_DPI, SAVE_DPI))
    else:
        img.save(path, "PNG", dpi=(SAVE_DPI, SAVE_DPI))

# ================= เครื่องพิมพ์ =================
HORZRES, VERTRES, LOGPIXELSX = 8, 10, 88
PHYSICALWIDTH, PHYSICALHEIGHT = 110, 111
PHYSICALOFFSETX, PHYSICALOFFSETY = 112, 113
DMPAPER_USER = 256

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
    try:
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return []

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
    h = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(h, 2)
    finally:
        win32print.ClosePrinter(h)
    problems = [t for bit, t in STATUS_PROBLEMS.items() if info["Status"] & bit]
    if info["Attributes"] & ATTR_WORK_OFFLINE:
        problems.insert(0, "เครื่องพิมพ์ไม่ได้เชื่อมต่อ (Windows ตั้งเป็นใช้งานแบบออฟไลน์)")
    return problems, info.get("cJobs", 0)

def open_dc(printer, size_mm=None):
    """size_mm = (กว้าง, ยาว) ที่จะสั่ง driver ชั่วคราว / None = ใช้ค่าใน driver"""
    if size_mm:
        try:
            h = win32print.OpenPrinter(printer)
            try:
                dm = win32print.GetPrinter(h, 2)["pDevMode"]
                if dm is not None:
                    dm.PaperSize = DMPAPER_USER
                    dm.PaperWidth = int(round(size_mm[0] * 10))   # หน่วย 0.1 mm
                    dm.PaperLength = int(round(size_mm[1] * 10))
                    dm.Fields |= (win32con.DM_PAPERSIZE | win32con.DM_PAPERWIDTH
                                  | win32con.DM_PAPERLENGTH)
                    try:
                        win32print.DocumentProperties(
                            0, h, printer, dm, dm,
                            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER)
                    except Exception:
                        pass
                    hdc = win32gui.CreateDC("WINSPOOL", printer, dm)
                    return win32ui.CreateDCFromHandle(hdc)
            finally:
                win32print.ClosePrinter(h)
        except Exception:
            pass                       # driver ไม่รับ -> ใช้ค่าเดิมใน driver
    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(printer)
    return dc

def page_info(dc):
    dpi = dc.GetDeviceCaps(LOGPIXELSX)
    pw, ph = dc.GetDeviceCaps(PHYSICALWIDTH), dc.GetDeviceCaps(PHYSICALHEIGHT)
    ox, oy = dc.GetDeviceCaps(PHYSICALOFFSETX), dc.GetDeviceCaps(PHYSICALOFFSETY)
    hr, vr = dc.GetDeviceCaps(HORZRES), dc.GetDeviceCaps(VERTRES)
    ux = max(ox, pw - ox - hr, 0)      # ขอบพิมพ์ไม่ได้ (เอาด้านที่มากกว่า)
    uy = max(oy, ph - oy - vr, 0)
    return {"dpi": dpi, "pw": pw, "ph": ph, "ox": ox, "oy": oy,
            "size_mm": (px2mm(pw, dpi), px2mm(ph, dpi)), "unp": (ux, uy)}

def read_printer_page(printer, size_mm=None):
    dc = open_dc(printer, size_mm)
    try:
        return page_info(dc)
    finally:
        dc.DeleteDC()

def layout_from_page(info):
    """แปลงเป็นแนวนอน (W >= H) สำหรับวาด"""
    pw, ph = info["pw"], info["ph"]
    ux, uy = info["unp"]
    return (pw, ph, ux, uy) if pw >= ph else (ph, pw, uy, ux)

def print_label(printer, size_mm, label, doc_name):
    dc = open_dc(printer, size_mm)
    started = False
    try:
        info = page_info(dc)
        pw, ph, ox, oy = info["pw"], info["ph"], info["ox"], info["oy"]
        if pw < ph:
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
    def __init__(self, master, title, label_img, size_mm, warnings, print_error):
        super().__init__(master)
        self.result = None
        self.can_print = print_error is None
        self.title(title)
        self.resizable(False, False)
        self.transient(master)

        scale = min(PREVIEW_MAX[0] / label_img.width, PREVIEW_MAX[1] / label_img.height)
        w, h = round(label_img.width * scale), round(label_img.height * scale)
        self.photo = ImageTk.PhotoImage(label_img.resize((w, h), Image.Resampling.LANCZOS))
        tk.Label(self, image=self.photo, bd=1, relief="solid").pack(padx=16, pady=(16, 4))
        tk.Label(self, text=f"ขนาดจริง {fmt_size(size_mm)}", fg="gray").pack()

        wrap = PREVIEW_MAX[0]
        for text in warnings:
            tk.Label(self, text="⚠ " + text, fg="#d97706", font=("Tahoma", 11, "bold"),
                     wraplength=wrap, justify="left").pack(padx=16, pady=2, anchor="w")
        if print_error:
            tk.Label(self, text=f"✖ ปริ้นไม่ได้: {print_error}\nยังบันทึกเป็นภาพได้ (Ctrl+S)",
                     fg="red", font=("Tahoma", 11, "bold"),
                     wraplength=wrap, justify="left").pack(padx=16, pady=2, anchor="w")

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
        self.bind("<Control-s>", lambda e: self.do_save())
        self.bind("<Control-S>", lambda e: self.do_save())

        self.update_idletasks()   # วางกลางจอ
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 3
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
        self.driver_size = None

        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)

        # เครื่องพิมพ์
        ttk.Label(f, text="เครื่องพิมพ์:").grid(row=0, column=0, sticky="w")
        printers = list_printers()
        self.printer_var = tk.StringVar(value=pick_printer(printers))
        cb = ttk.Combobox(f, textvariable=self.printer_var, values=printers,
                          state="readonly", width=48)
        cb.grid(row=0, column=1, sticky="we", pady=3)
        cb.bind("<<ComboboxSelected>>", self.on_printer_change)
        self.printer_info = tk.Label(f, font=("Tahoma", 9), fg="gray")
        self.printer_info.grid(row=1, column=1, sticky="w")

        # ขนาด label
        ttk.Label(f, text="ขนาด label:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.size_var = tk.StringVar(value=SIZE_AUTO)
        sc = ttk.Combobox(f, textvariable=self.size_var, state="readonly", width=48,
                          values=[SIZE_AUTO] + [p[0] for p in PRESETS] + [SIZE_CUSTOM])
        sc.grid(row=2, column=1, sticky="we", pady=(10, 0))
        sc.bind("<<ComboboxSelected>>", self.on_size_change)

        self.custom_frame = ttk.Frame(f)
        self.cw_var, self.ch_var = tk.StringVar(value="100"), tk.StringVar(value="50")
        ttk.Label(self.custom_frame, text="กว้าง").pack(side="left")
        ttk.Entry(self.custom_frame, textvariable=self.cw_var, width=7).pack(side="left", padx=4)
        ttk.Label(self.custom_frame, text="x  ยาว").pack(side="left")
        ttk.Entry(self.custom_frame, textvariable=self.ch_var, width=7).pack(side="left", padx=4)
        ttk.Label(self.custom_frame, text="mm").pack(side="left")
        self.custom_frame.grid(row=3, column=1, sticky="w", pady=(4, 0))
        self.custom_frame.grid_remove()
        for v in (self.cw_var, self.ch_var):
            v.trace_add("write", lambda *_: self.update_size_info())

        self.size_info = tk.Label(f, font=("Tahoma", 9), fg="gray",
                                  wraplength=560, justify="left")
        self.size_info.grid(row=4, column=1, sticky="w")

        # รูปแบบ
        ttk.Label(f, text="รูปแบบ:").grid(row=5, column=0, sticky="w", pady=(10, 0))
        self.kind_var = tk.StringVar(value="barcode")
        mf = ttk.Frame(f)
        mf.grid(row=5, column=1, sticky="w", pady=(10, 0))
        for value, text in (("barcode", "บาร์โค้ด (Code 128)"), ("qr", "QR Code")):
            ttk.Radiobutton(mf, text=text, value=value, variable=self.kind_var,
                            command=self.on_kind_change).pack(side="left", padx=(0, 16))

        # ข้อความ
        ttk.Label(f, text="ข้อความ / เลข:").grid(row=6, column=0, sticky="w", pady=(10, 0))
        self.code_var = tk.StringVar()
        self.entry = tk.Entry(f, textvariable=self.code_var, font=("Consolas", 22), width=24)
        self.entry.grid(row=6, column=1, sticky="we", pady=(10, 0))
        self.entry.bind("<Return>", self.on_enter)
        self.entry.bind("<KP_Enter>", self.on_enter)
        self.code_var.trace_add("write", lambda *_: self.on_text_change())

        self.check_label = tk.Label(f, font=("Tahoma", 10), anchor="w")
        self.check_label.grid(row=7, column=1, sticky="w")

        ttk.Button(f, text="ดูตัวอย่าง (Enter)", command=self.on_enter).grid(
            row=8, column=1, sticky="we", pady=8, ipady=6)

        self.status = tk.Label(f, text="ยิงหรือพิมพ์เลข แล้วกด Enter",
                               font=("Tahoma", 11, "bold"), fg="gray",
                               wraplength=620, justify="left")
        self.status.grid(row=9, column=0, columnspan=2, sticky="w")

        self.on_printer_change()
        self.entry.focus_set()

    # ---------- เครื่องพิมพ์ / ขนาด ----------
    def on_printer_change(self, event=None):
        p = self.printer_var.get()
        self.driver_size = None
        if not p:
            self.printer_info.config(text="ไม่พบเครื่องพิมพ์ (ยังบันทึกเป็นภาพได้)", fg="red")
        else:
            try:
                info = read_printer_page(p)
                self.driver_size = info["size_mm"]
                ux, uy = (px2mm(v, info["dpi"]) for v in info["unp"])
                text = f"ความละเอียด {info['dpi']} dpi • กระดาษใน driver {fmt_size(self.driver_size)}"
                if ux > 0.5 or uy > 0.5:
                    text += f" • ขอบพิมพ์ไม่ได้ {max(ux, uy):.1f} mm"
                self.printer_info.config(text=text, fg="gray")
            except Exception as e:
                self.printer_info.config(text=f"อ่านข้อมูลเครื่องพิมพ์ไม่ได้: {e}", fg="red")
        self.update_size_info()

    def on_size_change(self, event=None):
        if self.size_var.get() == SIZE_CUSTOM:
            self.custom_frame.grid()
        else:
            self.custom_frame.grid_remove()
        self.update_size_info()
        self.entry.focus_set()

    def selected_size(self):
        """None = ตามเครื่องพิมพ์ / (กว้าง, ยาว) mm"""
        name = self.size_var.get()
        if name == SIZE_AUTO:
            return None
        if name == SIZE_CUSTOM:
            try:
                w, h = float(self.cw_var.get()), float(self.ch_var.get())
            except ValueError:
                raise ValueError("ใส่ขนาดเป็นตัวเลข (mm)")
            lo, hi = CUSTOM_RANGE_MM
            if not (lo <= w <= hi and lo <= h <= hi):
                raise ValueError(f"ขนาดต้องอยู่ระหว่าง {lo}-{hi} mm")
            return (w, h)
        for p in PRESETS:
            if p[0] == name:
                return (p[1], p[2])
        return None

    def update_size_info(self):
        try:
            size = self.selected_size()
        except ValueError as e:
            self.size_info.config(text=f"✖ {e}", fg="red")
            return
        drv = self.driver_size
        if size is None:
            if drv:
                text, color = f"ใช้ขนาดตาม driver: {fmt_size(drv)}", "gray"
            else:
                text, color = f"ไม่พบเครื่องพิมพ์ — ใช้ {fmt_size(FALLBACK_SIZE)} สำหรับบันทึกภาพ", "gray"
        elif drv is None or same_size(size, drv):
            text, color = "✔ ตรงกับขนาดใน driver" if drv else "", "green"
        else:
            text = (f"⚠ ต่างจาก driver ({fmt_size(drv)}) — ตอนปริ้นโปรแกรมจะสั่งขนาดนี้ชั่วคราว "
                    "ม้วนสติกเกอร์ในเครื่องต้องเป็นขนาดนี้จริง")
            color = "#d97706"
        self.size_info.config(text=text, fg=color)

    # ---------- ตรวจข้อความ ----------
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
            req_size = self.selected_size()
            check_encodable(code, kind)
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return self.select_entry()

        printer = self.printer_var.get()
        jobs, print_error, info, warnings = 0, None, None, []
        if not printer:
            print_error = "ไม่พบเครื่องพิมพ์"
        else:
            try:
                problems, jobs = check_printer(printer)
                if problems:
                    print_error = " / ".join(problems)
                info = read_printer_page(printer, req_size)
            except Exception as e:
                print_error = f"ติดต่อเครื่องพิมพ์ไม่ได้ ({e})"

        if info:
            W, H, ux, uy = layout_from_page(info)
            dpi, size_mm = info["dpi"], info["size_mm"]
            if req_size and not same_size(req_size, size_mm):
                warnings.append(f"driver ไม่ยอมเปลี่ยนเป็น {fmt_size(req_size)} "
                                f"จึงใช้ขนาด {fmt_size(size_mm)} แทน")
        else:
            size_mm = req_size or FALLBACK_SIZE
            dpi = DEFAULT_DPI
            W, H = mm2px(max(size_mm), dpi), mm2px(min(size_mm), dpi)
            ux = uy = 0

        try:
            label, notes = render_label(code, kind, dpi, W, H, ux, uy)
        except ValueError as e:
            messagebox.showerror(APP_TITLE, str(e))
            return self.select_entry()
        warnings += notes

        k, msg = validate(code)
        if k == "bad":
            warnings.append(msg)
        if jobs and not print_error:
            warnings.append(f"มีงานค้างในคิวเครื่องพิมพ์ {jobs} งาน — "
                            "เครื่องพิมพ์อาจไม่ได้เชื่อมต่อ หรือกำลังพิมพ์งานอื่นอยู่")

        title = f"ตัวอย่าง {KIND_NAME[kind]} — {code}"
        action = PreviewDialog(self, title, label, size_mm, warnings, print_error).result
        if action == "print":
            self.print_now(printer, req_size, label, code, kind)
        elif action == "save":
            self.save_now(code, kind, size_mm)
        else:
            self.status.config(text="ยกเลิกแล้ว", fg="gray")
            self.select_entry()
        return "break"

    def print_now(self, printer, req_size, label, code, kind):
        try:
            print_label(printer, req_size, label, f"{KIND_NAME[kind]} {code}")
        except Exception as e:
            self.status.config(text=f"✖ ส่งงานไม่สำเร็จ: {code}", fg="red")
            messagebox.showerror(APP_TITLE, f"ส่งงานไปเครื่องพิมพ์ไม่สำเร็จ\n\n{e}")
            return self.select_entry()
        self.status.config(
            text=f"➜ ส่งงานไปเครื่องพิมพ์แล้ว: {code} ({datetime.now():%H:%M:%S}) "
                 "— ตรวจดูว่าสติกเกอร์ออกมาหรือไม่", fg="green")
        self.code_var.set("")
        self.entry.focus_set()

    def save_now(self, code, kind, size_mm):
        safe = re.sub(r'[\\/:*?"<>|\s]+', "_", code).strip("._")[:60] or "label"
        path = filedialog.asksaveasfilename(
            parent=self, title="บันทึกเป็นภาพ", initialdir=self.last_dir,
            initialfile=f"{safe}_{kind}.png", defaultextension=".png",
            filetypes=[("PNG (แนะนำ)", "*.png"), ("JPEG", "*.jpg"), ("PDF", "*.pdf")])
        if not path:
            self.status.config(text="ยกเลิกการบันทึก", fg="gray")
            return self.select_entry()
        try:
            save_label_image(path, code, kind, size_mm)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"บันทึกไฟล์ไม่สำเร็จ\n\n{e}")
            return self.select_entry()
        self.last_dir = os.path.dirname(path)
        self.status.config(text=f"✔ บันทึกภาพแล้ว ({fmt_size(size_mm)}): {path}", fg="green")
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