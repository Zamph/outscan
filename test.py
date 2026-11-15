"""แสดงบาร์โค้ดจากไฟล์ Excel/CSV ทีละอัน มี 2 โหมด
  - waybill_no (คอลัมน์ A)
  - Drop (คอลัมน์ E)
ใช้งาน:
  - ลูกศรลง/ขวา = อันถัดไป, ลูกศรขึ้น/ซ้าย = ย้อนกลับ
  - ปุ่ม "เริ่ม" (หรือกด Space) = เลื่อนอัตโนมัติจนถึงอันสุดท้าย กดอีกครั้งเพื่อหยุด
  - ช่อง "ลำดับที่" พิมพ์เลขแล้วกด Enter เพื่อข้ามไป
  - ช่อง "ค้นหา" พิมพ์บางส่วนของเลข/ชื่อ แล้วกด Enter (กด Enter ซ้ำ = ผลถัดไป)

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

MODES = {"waybill": "waybill_no (A)", "drop": "Drop (E)"}


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def read_sheets(path):
    if path.lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [list(csv.reader(f))]
    wb = load_workbook(path, read_only=True, data_only=True)
    return [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]


def extract(rows, kind):
    """คืน [(รหัสที่ใช้ทำบาร์โค้ด, drop_name), ...] จากคอลัมน์ waybill_no หรือ Drop"""
    if not rows:
        return []
    default_col = 0 if kind == "waybill" else 4  # A หรือ E
    col, name_col, has_header = default_col, None, False
    for i, h in enumerate(rows[0]):
        t = norm(h).lower()
        if kind == "waybill" and "waybill" in t and not has_header:
            col, has_header = i, True
        if kind == "drop" and t == "drop":
            col, has_header = i, True
        if "drop_name" in t:
            name_col = i
    if not has_header and rows[0] and "waybill" in norm(rows[0][0]).lower():
        has_header = True  # มีหัวตารางแต่หาคอลัมน์ตามชื่อไม่เจอ
    items = []
    for r in (rows[1:] if has_header else rows):
        code = norm(r[col]) if col < len(r) else ""
        if not code:
            continue
        name = norm(r[name_col]) if name_col is not None and name_col < len(r) else ""
        items.append((code, name))
    return items


def load_all(path):
    sheets = read_sheets(path)
    result = {}
    for kind in MODES:
        result[kind] = next((it for it in (extract(r, kind) for r in sheets) if it), [])
    return result


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
    def __init__(self, root, data):
        self.root, self.data = root, data
        self.mode = "waybill" if data["waybill"] else "drop"
        self.items, self.i = data[self.mode], 0
        self.playing, self.job = False, None
        self.matches, self.mpos, self.last_q = [], -1, None

        root.title("Barcode Viewer")
        root.configure(bg="black")
        root.geometry("820x740")
        root.lift()
        root.attributes("-topmost", True)  # ดึงหน้าต่างขึ้นมาอยู่หน้าสุดชั่วครู่
        root.after(300, lambda: root.attributes("-topmost", False))
        root.focus_force()

        # ---- แถวบน: เลือกโหมด + ค้นหา ----
        top = tk.Frame(root, bg="black")
        top.pack(pady=(14, 0))
        self.mode_btns = {}
        for k, label in MODES.items():
            b = tk.Button(top, text=label, command=lambda k=k: self.set_mode(k),
                          takefocus=0, relief="flat", font=("Tahoma", 11, "bold"),
                          padx=12, pady=3)
            b.pack(side="left", padx=(0, 6))
            self.mode_btns[k] = b
        tk.Label(top, text="ค้นหา", bg="black", fg="#ddd",
                 font=("Tahoma", 11)).pack(side="left", padx=(18, 6))
        self.qvar = tk.StringVar()
        self.search = tk.Entry(top, textvariable=self.qvar, width=20,
                               font=("Tahoma", 12), relief="solid", bd=1)
        self.search.pack(side="left")
        tk.Button(top, text="🔍", command=self.do_search, takefocus=0, relief="flat",
                  bg="#222", fg="white", font=("Tahoma", 11),
                  padx=8).pack(side="left", padx=(6, 0))
        self.status = tk.Label(root, bg="black", fg="#9a9a9a", font=("Tahoma", 10))
        self.status.pack(pady=(4, 0))

        # ---- ชื่อด้านบน ----
        self.title = tk.Label(root, bg="black", fg="white", font=("Tahoma", 20, "bold"))
        self.title.pack(pady=(8, 0))
        self.sub = tk.Label(root, bg="black", fg="#9a9a9a", font=("Tahoma", 12))
        self.sub.pack(pady=(2, 10))

        # ---- กรอบบาร์โค้ด ----
        panel = tk.Frame(root, bg="white")
        panel.pack(padx=24, pady=(0, 16), fill="x")
        self.img_label = tk.Label(panel, bg="white")
        self.img_label.pack(padx=16, pady=24)

        # ---- ลำดับที่ (พิมพ์ได้) ----
        seq = tk.Frame(root, bg="white")
        seq.pack()
        self.var = tk.StringVar()
        self.entry = tk.Entry(seq, textvariable=self.var, width=6, justify="center",
                              font=("Tahoma", 18, "bold"), relief="solid", bd=1)
        self.entry.pack(side="left", padx=(20, 12), pady=8)
        self.var.trace_add("write", self.only_digits)
        tk.Label(seq, text="ลำดับที่", bg="white", fg="#111",
                 font=("Tahoma", 13)).pack(side="left")
        self.total = tk.Label(seq, bg="white", fg="#777", font=("Tahoma", 11))
        self.total.pack(side="left", padx=(12, 20))

        # ---- ปุ่มเริ่ม/หยุด และความเร็ว ----
        ctl = tk.Frame(root, bg="black")
        ctl.pack(pady=(14, 0))
        self.btn = tk.Button(ctl, text="▶ เริ่ม", command=self.toggle, takefocus=0,
                             font=("Tahoma", 13, "bold"), bg="#3b82f6", fg="white",
                             activebackground="#2563eb", activeforeground="white",
                             relief="flat", padx=18, pady=4)
        self.btn.pack(side="left")
        tk.Label(ctl, text="เลื่อนทุก", bg="black", fg="#ddd",
                 font=("Tahoma", 11)).pack(side="left", padx=(16, 6))
        self.speed = tk.StringVar(value="2")
        tk.Spinbox(ctl, values=("0.5", "1", "2", "3", "5", "10"), textvariable=self.speed,
                   width=4, state="readonly", font=("Tahoma", 12),
                   takefocus=0).pack(side="left")
        tk.Label(ctl, text="วินาที", bg="black", fg="#ddd",
                 font=("Tahoma", 11)).pack(side="left", padx=(6, 0))

        tk.Label(root, text="ลูกศรลง ↓ = ถัดไป   ลูกศรขึ้น ↑ = ย้อนกลับ   Space = เริ่ม/หยุด\n"
                            "พิมพ์เลขที่ช่องลำดับที่ แล้วกด Enter เพื่อข้ามไป",
                 bg="black", fg="#9a9a9a", font=("Tahoma", 10)).pack(pady=12)

        for k in ("<Down>", "<Right>"):
            root.bind(k, lambda e: self.move(1))
        for k in ("<Up>", "<Left>"):
            root.bind(k, lambda e: self.move(-1))
        root.bind("<space>", lambda e: self.toggle())
        self.entry.bind("<Return>", self.jump)
        self.entry.bind("<FocusIn>", lambda e: self.entry.selection_range(0, "end"))
        self.entry.bind("<FocusOut>", lambda e: self.var.set(str(self.i + 1)))
        self.search.bind("<Return>", self.do_search)
        self.search.bind("<Escape>", lambda e: self.root.focus_set())
        root.bind("<Button-1>", self.click_away)
        self.refresh_mode_buttons()
        self.show()

    # ---------- ช่วยเหลือ ----------
    def typing(self):
        return self.root.focus_get() in (self.entry, self.search)

    def only_digits(self, *_):
        cur = self.var.get()
        s = "".join(c for c in cur.translate(THAI_TO_DIGIT) if c.isdigit())
        if s != cur:
            self.var.set(s)

    def click_away(self, e):
        if e.widget in (self.entry, self.search):
            e.widget.focus_force()
        else:
            self.root.focus_set()

    # ---------- โหมด ----------
    def refresh_mode_buttons(self):
        for k, b in self.mode_btns.items():
            on = k == self.mode
            b.config(bg="#3b82f6" if on else "#222", fg="white",
                     activebackground="#2563eb" if on else "#333", activeforeground="white")

    def set_mode(self, kind):
        if kind == self.mode:
            return
        if not self.data[kind]:
            messagebox.showinfo("ไม่พบข้อมูล", f"ไม่พบข้อมูลของ {MODES[kind]} ในไฟล์")
            return
        self.stop()
        self.mode, self.items, self.i = kind, self.data[kind], 0
        self.matches, self.mpos, self.last_q = [], -1, None
        self.status.config(text="")
        self.refresh_mode_buttons()
        self.show()

    # ---------- ค้นหา ----------
    def do_search(self, _=None):
        q = self.qvar.get().strip().lower()
        if not q:
            self.status.config(text="")
            return
        if q != self.last_q:
            self.matches = [k for k, (c, n) in enumerate(self.items)
                            if q in c.lower() or q in n.lower()]
            self.mpos, self.last_q = -1, q
        if not self.matches:
            self.status.config(text="ไม่พบรายการที่ค้นหา", fg="#f87171")
            return
        self.stop()
        self.mpos = (self.mpos + 1) % len(self.matches)
        self.go(self.matches[self.mpos])
        self.status.config(
            text=f"พบ {len(self.matches)} รายการ (ผลที่ {self.mpos + 1}) กด Enter เพื่อดูผลถัดไป",
            fg="#9a9a9a")

    # ---------- เลื่อน ----------
    def move(self, d):
        if self.typing():  # กำลังพิมพ์ ไม่ต้องเลื่อน
            return
        self.go(self.i + d)

    def toggle(self):
        if self.typing():
            return
        if self.playing:
            self.stop()
            return
        if self.i >= len(self.items) - 1:  # อยู่ท้ายสุดแล้ว เริ่มใหม่จากต้น
            self.go(0)
        self.playing = True
        self.btn.config(text="⏸ หยุด")
        self.schedule()

    def schedule(self):
        try:
            ms = int(float(self.speed.get()) * 1000)
        except ValueError:
            ms = 2000
        self.job = self.root.after(ms, self.tick)

    def tick(self):
        if self.i < len(self.items) - 1:
            self.go(self.i + 1)
        if self.i < len(self.items) - 1:
            self.schedule()
        else:
            self.stop()  # ถึงอันสุดท้ายแล้ว

    def stop(self):
        self.playing = False
        if self.job:
            self.root.after_cancel(self.job)
            self.job = None
        self.btn.config(text="▶ เริ่ม")

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
        self.total.config(text=f"จาก {len(self.items)}")
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
        data = load_all(path)
    except Exception as e:
        messagebox.showerror("อ่านไฟล์ไม่ได้", str(e))
        return
    if not any(data.values()):
        messagebox.showerror("ไม่พบข้อมูล", "ไม่พบเลข waybill หรือ Drop ในไฟล์")
        return
    App(root, data)
    root.mainloop()


if __name__ == "__main__":
    main()