# Barcode Reprint S10 (ปริ้นบาร์โค้ด / QR Code สำรองสำหรับฉลากพัสดุและสินค้า)

โปรแกรม Desktop GUI สำหรับสร้างและพิมพ์บาร์โค้ด (Code 128) และ QR Code สำรองบนสติกเกอร์ฉลากหลายขนาด (ปรับตามเครื่องพิมพ์อัตโนมัติหรือกำหนดเอง) พร้อมระบบตรวจสอบความถูกต้องของเลขพัสดุตามมาตรฐานสากล UPU S10 และแก้ปัญหาแป้นพิมพ์ภาษาไทยอัตโนมัติ

---

## ✨ ฟีเจอร์หลัก (Key Features)

- **รองรับทั้ง Barcode (Code 128) และ QR Code**: เลือกประเภทบาร์โค้ดที่ต้องการพิมพ์ได้อย่างยืดหยุ่น
- **ตรวจสอบเลขพัสดุ UPU S10**: มีฟังก์ชันคำนวณ Check Digit ตามมาตรฐาน S10 เพื่อแจ้งเตือนหากยิงหรือพิมพ์เลขผิดพลาด
- **ระบบแก้แป้นพิมพ์ภาษาไทยอัตโนมัติ**: หากลืมเปลี่ยนภาษาเป็นภาษาอังกฤษตอนยิงบาร์โค้ด ระบบจะถอดรหัสแป้นไทยกลับเป็นอังกฤษให้อัตโนมัติ
- **รองรับเครื่องพิมพ์สติกเกอร์ (Thermal Printer)**: สั่งพิมพ์โดยตรงผ่านไดรเวอร์ Windows (`pywin32`) พร้อมหมุนภาพให้อัตโนมัติ (90 องศา)
- **พรีวิวและบันทึกไฟล์ภาพ**: ดูตัวอย่างบนจอก่อนพิมพ์ และรองรับการบันทึกภาพความละเอียดสูง (600 DPI) ด้วยช็อตคัต `Ctrl+S`
- **ปุ่มคีย์ลัดสะดวกรวดเร็ว**:
  - `Enter`: สลับดูพรีวิว / สั่งพิมพ์ทันที
  - `Ctrl + S`: บันทึกรูปภาพลงเครื่อง
  - `Esc`: รีเซ็ตหรือยกเลิก

---

## 📥 วิธีติดตั้งและรันจาก Source Code

### ข้อกำหนดระบบ (Prerequisites)
- **Windows OS** (เนื่องจากใช้ Win32 API สำหรับระบบ Print Spooler)
- **Python 3.8+**

### ติดตั้งไลบรารีที่จำเป็น
```bash
git clone https://github.com/fiwflash/barcode-reprint-s10.git
cd barcode-reprint-s10
pip install -r requirements.txt
```

### การเรียกใช้งาน
```bash
python barcode_reprint.py
```

---

## 🛠️ วิธีคอมไพล์เป็นไฟล์ Executable (.exe)

หากต้องการนำไปแจกจ่ายหรือใช้งานโดยไม่ต้องลง Python ให้บิลด์ผ่าน `pyinstaller`:

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "BarcodeReprint" barcode_reprint.py
```
ไฟล์ `.exe` จะถูกสร้างขึ้นในโฟลเดอร์ `dist/BarcodeReprint.exe`

> 💡 **สำหรับผู้ใช้งานทั่วไป:** สามารถดาวน์โหลดไฟล์ `.exe` พร้อมใช้งานได้โดยตรงที่หน้า [Releases](https://github.com/fiwflash/barcode-reprint-s10/releases)

---

## 📁 โครงสร้างโปรเจกต์ (Project Structure)

```text
├── barcode_reprint.py    # ซอร์สโค้ดหลัก (Tkinter GUI, Barcode/QR Generator, Win32 Print)
├── requirements.txt      # รายการแพ็กเกจ Python ที่จำเป็น
├── .gitignore            # การตั้งค่าละเว้นไฟล์ชั่วคราวและ build outputs
└── README.md             # คู่มือการติดตั้งและใช้งานโปรเจกต์
```

---

## ⚙️ เทคโนโลยีที่ใช้ (Tech Stack)

- **GUI Framework:** Python Tkinter
- **Image Processing & Barcode:** Pillow (PIL), qrcode
- **Printer Integration:** `pywin32` (win32print, win32ui)
- **Standards:** Universal Postal Union (UPU) S10, Code 128 Specification

---

## 📄 ใบอนุญาต (License)

โปรเจกต์นี้เปิดให้ใช้งานและพัฒนาต่อเพื่อการศึกษาและการปฏิบัติงาน (MIT License หรือ Public Domain)
