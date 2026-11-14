"""แสดงบาร์โค้ดจากคอลัมน์ waybill_no ทีละอัน
- ลูกศรลง/ขวา/Space = อันถัดไป, ลูกศรขึ้น/ซ้าย = ย้อนกลับ
- คลิกที่ช่องตัวเลขด้านล่าง พิมพ์ลำดับที่ต้องการ แล้วกด Enter เพื่อข้ามไปอันนั้น
- ด้านบนบาร์โค้ดแสดงชื่อ (waybill_no) และชื่อสาขา (drop_name ถ้ามีในไฟล์)

ติดตั้ง:  pip install python-barcode pillow openpyxl "setuptools<81"
รัน:      python waybill_barcode.py            (จะมีหน้าต่างให้เลือกไฟล์)
หรือ:     python waybill_barcode.py data.xlsx
"""
import csv
import io
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import barcode
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from PIL import Image, ImageTk


# แปลงตัวอักษรไทยบนแถวตัวเลข (คีย์บอร์ดเกษมณี) และเลขไทย ให้เป็นเลขอารบิก
THAI_TO_DIGIT = str.maketrans("ๅ/-ภถุึคตจ๑๒๓๔๕๖๗๘๙๐", "12345678901234567890")


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def load_items(path):
    """คืนรายการ [(waybill_no, drop_name), ...]"""
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            sheets = [list(csv.reader(f))]
    else:
        wb = load_workbook(path, read_only=True, data_only=True)
        sheets = [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]

    for rows in sheets:
        if not rows:
            continue
        col, name_col, has_header = 0, None, False
        for i, h in enumerate(rows[0]):
            t = norm(h).lower()
            if "waybill" in t and not has_header:
                col, has_header = i, True
            if "drop_name" in t:
                name_col = i
        data = rows[1:] if has_header else rows
        items = []
        for r in data:
            code = norm(r[col]) if col < len(r) else ""
            if not code:
                continue
            name = norm(r[name_col]) if name_col is not None and name_col < len(r) else ""
            items.append((code, name))
        if items:
            return items
    return []


def make_image(text, width=720):
    code = barcode.get("code128", text, writer=ImageWriter())
    buf = io.BytesIO()
    code.write(buf, options={
        "module_width": 0.4, "module_height": 20, "font_size": 11,
        "text_distance": 4, "quiet_zone": 6, "dpi": 200,
    })
    buf.seek(0)
    img = Image.open(buf).convert("RGB")
    ratio = width / img.width
    return img.resize((width, int(img.height * ratio)), Image.LANCZOS)


class App:
    def __init__(self, root, items):
        self.root, self.items, self.i = root, items, 0
        root.title("Waybill Barcode")
        root.configure(bg="black")
        root.geometry("820x580")
        root.lift()
        root.attributes("-topmost", True)  # ดึงหน้าต่างขึ้นมาอยู่หน้าสุดชั่วครู่
        root.after(300, lambda: root.attributes("-topmost", False))
        root.focus_force()

        # ชื่อด้านบน
        self.title = tk.Label(root, bg="black", fg="white", font=("Tahoma", 20, "bold"))
        self.title.pack(pady=(24, 0))
        self.sub = tk.Label(root, bg="black", fg="#9a9a9a", font=("Tahoma", 12))
        self.sub.pack(pady=(2, 10))

        # กรอบบาร์โค้ด
        panel = tk.Frame(root, bg="white")
        panel.pack(padx=24, pady=(0, 16), fill="x")
        self.img_label = tk.Label(panel, bg="white")
        self.img_label.pack(padx=16, pady=24)

        # ลำดับที่ (พิมพ์ได้)
        seq = tk.Frame(root, bg="white")
        seq.pack()
        self.var = tk.StringVar()
        self.entry = tk.Entry(seq, textvariable=self.var, width=6, justify="center",
                              font=("Tahoma", 18, "bold"), relief="solid", bd=1)
        self.entry.pack(side="left", padx=(20, 12), pady=8)
        self.var.trace_add("write", self.only_digits)
        tk.Label(seq, text="ลำดับที่", bg="white", fg="#111",
                 font=("Tahoma", 13)).pack(side="left")
        tk.Label(seq, text=f"จาก {len(items)}", bg="white", fg="#777",
                 font=("Tahoma", 11)).pack(side="left", padx=(12, 20))

        tk.Label(root, text="ลูกศรลง ↓ = ถัดไป   ลูกศรขึ้น ↑ = ย้อนกลับ   "
                            "หรือพิมพ์เลขแล้วกด Enter เพื่อข้ามไป",
                 bg="black", fg="#9a9a9a", font=("Tahoma", 10)).pack(pady=12)

        for k in ("<Down>", "<Right>", "<space>"):
            root.bind(k, lambda e: self.move(1))
        for k in ("<Up>", "<Left>"):
            root.bind(k, lambda e: self.move(-1))
        self.entry.bind("<Return>", self.jump)
        self.entry.bind("<FocusIn>", lambda e: self.entry.selection_range(0, "end"))
        self.entry.bind("<FocusOut>", lambda e: self.var.set(str(self.i + 1)))
        root.bind("<Button-1>", self.click_away)
        self.show()

    def only_digits(self, *_):
        cur = self.var.get()
        s = "".join(c for c in cur.translate(THAI_TO_DIGIT) if c.isdigit())
        if s != cur:
            self.var.set(s)

    def click_away(self, e):
        if e.widget is self.entry:
            self.entry.focus_force()
        else:
            self.root.focus_set()

    def move(self, d):
        if self.root.focus_get() is self.entry:  # กำลังพิมพ์ ไม่ต้องเลื่อน
            return
        self.go(self.i + d)

    def jump(self, _=None):
        try:
            n = int(self.var.get()) - 1
        except ValueError:
            n = self.i
        self.go(max(0, min(len(self.items) - 1, n)))
        self.root.focus_set()

    def go(self, n):
        if 0 <= n < len(self.items):
            self.i = n
            self.show()

    def show(self):
        code, name = self.items[self.i]
        self.var.set(str(self.i + 1))
        self.title.config(text=code)
        self.sub.config(text=name)
        try:
            self.photo = ImageTk.PhotoImage(make_image(code))
            self.img_label.config(image=self.photo, text="")
        except Exception as e:
            self.img_label.config(image="", text=f"สร้างบาร์โค้ดไม่ได้: {code}\n{e}",
                                  fg="red")


def main():
    root = tk.Tk()
    path = sys.argv[1] if len(sys.argv) > 1 else filedialog.askopenfilename(
        title="เลือกไฟล์ Excel/CSV",
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
    if not path:
        return
    try:
        items = load_items(path)
    except Exception as e:
        messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
        return
    if not items:
        messagebox.showerror("ไม่พบข้อมูล", "ไม่พบเลข waybill ในไฟล์")
        return
    App(root, items)
    root.mainloop()


if __name__ == "__main__":
    main()