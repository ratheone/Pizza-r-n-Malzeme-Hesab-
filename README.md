#streamlit run pizza_Tkinter.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import json
import os
import openpyxl
import random

# ---------------------------------------------------------
# 1. VERİ VE AYARLAR
# ---------------------------------------------------------
DEFAULT_MALZEMELER = [
    {"Malzeme": "Peynir", "Oran": 13.3, "Fiyat": 70},
    {"Malzeme": "Mantar", "Oran": 6.7, "Fiyat": 180},
    {"Malzeme": "Zeytin", "Oran": 4.0, "Fiyat": 40},
    {"Malzeme": "Biber", "Oran": 4.0, "Fiyat": 60},
    {"Malzeme": "Kabak", "Oran": 4.0, "Fiyat": 50},
    {"Malzeme": "Domates", "Oran": 6.7, "Fiyat": 8},
    {"Malzeme": "Sucuk", "Oran": 6.7, "Fiyat": 110},
    {"Malzeme": "Salam", "Oran": 4.0, "Fiyat": 163},
    {"Malzeme": "Sosis", "Oran": 4.0, "Fiyat": 30},
    {"Malzeme": "Balık", "Oran": 6.7, "Fiyat": 25},
    {"Malzeme": "Tavuk", "Oran": 6.7, "Fiyat": 20},
    {"Malzeme": "Mısır", "Oran": 2.7, "Fiyat": 20},
    {"Malzeme": "Soğan", "Oran": 2.7, "Fiyat": 5},
    {"Malzeme": "Sarımsak", "Oran": 0.7, "Fiyat": 25},
    {"Malzeme": "Kabak", "Oran": 0.8, "Fiyat": 30}
]

HAMUR_ARALIKLARI = {
    "Küçük": {"ince": (120, 150), "kalin": (180, 200)},
    "Orta": {"ince": (180, 220), "kalin": (250, 280)},
    "Büyük": {"ince": (220, 280), "kalin": (300, 350)}
}

# --- RENK PALETİ ---
COLOR_ACCENT = "#e67e22"        # Turuncu
COLOR_ACCENT_HOVER = "#d35400"
COLOR_BG_MAIN = "#ffffff"
COLOR_TEXT = "#2c3e50"
COLOR_MENU_BTN = "#34495e"

class PizzaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lucy Pizza - Professional Calculator")
        self.root.geometry("1100x750")
        
        self.root.attributes('-alpha', 0.96)
        
        self.malzemeler = {}
        self.hamur_fiyat = 7.5
        self.kar_orani = 140.0
        self.load_data()

        self.menu_acik = True
        self.bg_image_path = "background.png"
        self.bg_image_original = None
        self.bg_photo = None
        
        self.setup_background()
        self.create_styles()
        self.create_interface()

    def load_data(self):
        json_file = "pizza_data.json"
        xlsx_file = "pizza_data.xlsx"
        
        for m in DEFAULT_MALZEMELER:
            self.malzemeler[m["Malzeme"]] = {"oran": m["Oran"], "fiyat": m["Fiyat"]}

        if os.path.exists(json_file):
            try:
                with open(json_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                raw = data.get("malzemeler", [])
                if raw: self.malzemeler = {m["Malzeme"]: {"oran": m["Oran"], "fiyat": m["Fiyat"]} for m in raw}
                if "ayarlar" in data:
                    self.hamur_fiyat = float(data["ayarlar"].get("hamur_fiyat", 7.5))
                    self.kar_orani = float(data["ayarlar"].get("kar_orani", 140.0))
            except: pass
        elif os.path.exists(xlsx_file):
            try:
                wb = openpyxl.load_workbook(xlsx_file)
                if "Malzemeler" in wb.sheetnames:
                    temp_malz = {}
                    for row in wb["Malzemeler"].iter_rows(min_row=2, values_only=True):
                        if row[0]: temp_malz[row[0]] = {"oran": row[1], "fiyat": row[2]}
                    if temp_malz: self.malzemeler = temp_malz
                if "Ayarlar" in wb.sheetnames:
                    for row in wb["Ayarlar"].iter_rows(min_row=2, values_only=True):
                        if row[0] == "Hamur 100g Fiyatı": self.hamur_fiyat = float(row[1])
                        if row[0] == "Kar Oranı (%)": self.kar_orani = float(row[1])
            except: pass

    def setup_background(self):
        self.canvas_bg = tk.Canvas(self.root, highlightthickness=0)
        self.canvas_bg.pack(fill="both", expand=True)

        if os.path.exists(self.bg_image_path):
            self.bg_image_original = Image.open(self.bg_image_path)
            self.root.bind("<Configure>", self.resize_background)
        else:
            self.canvas_bg.configure(bg="#f0f2f6")

    def resize_background(self, event):
        if self.bg_image_original:
            w, h = event.width, event.height
            if w > 0 and h > 0:
                resized = self.bg_image_original.resize((w, h), Image.Resampling.LANCZOS)
                self.bg_photo = ImageTk.PhotoImage(resized)
                self.canvas_bg.create_image(0, 0, image=self.bg_photo, anchor="nw")

    def create_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("Card.TFrame", background=COLOR_BG_MAIN, relief="flat")
        style.configure("TLabel", background=COLOR_BG_MAIN, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 18, "bold"), foreground=COLOR_ACCENT)
        style.configure("SubHeader.TLabel", font=("Segoe UI", 11, "bold"), foreground="#555")
        
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), background=COLOR_ACCENT, foreground="white", borderwidth=0)
        style.map("Accent.TButton", background=[("active", COLOR_ACCENT_HOVER), ("pressed", "#a04000")], relief=[("pressed", "sunken")])
        
        style.configure("Menu.TButton", font=("Segoe UI", 12, "bold"), background=COLOR_MENU_BTN, foreground="white", borderwidth=0)
        style.map("Menu.TButton", background=[("active", "#2c3e50")])

        style.configure("TButton", font=("Segoe UI", 9))
        style.configure("TCheckbutton", background=COLOR_BG_MAIN, font=("Segoe UI", 10))
        style.map("TCheckbutton", foreground=[("active", COLOR_ACCENT)])
        style.configure("TRadiobutton", background=COLOR_BG_MAIN, font=("Segoe UI", 10))
        style.map("TRadiobutton", foreground=[("active", COLOR_ACCENT)])

    def create_interface(self):
        self.btn_menu = ttk.Button(self.canvas_bg, text="☰ AYARLAR", style="Menu.TButton", command=self.toggle_menu)
        self.btn_menu.place(x=40, y=20, height=40, width=120)

        # --- SOL PANEL ---
        self.frame_left = ttk.Frame(self.canvas_bg, style="Card.TFrame", padding=25)
        self.win_left = self.canvas_bg.create_window(40, 70, window=self.frame_left, anchor="nw", width=420, height=650)

        btn_area = ttk.Frame(self.frame_left, style="Card.TFrame")
        btn_area.pack(side="bottom", fill="x", pady=(10, 0))
        ttk.Button(btn_area, text="Maliyet Hesapla", style="Accent.TButton", command=self.hesapla, cursor="hand2").pack(fill="x", pady=(0, 10), ipady=8)
        sub_btns = ttk.Frame(btn_area, style="Card.TFrame")
        sub_btns.pack(fill="x")
        ttk.Button(sub_btns, text="Kaydet", command=self.indir).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(sub_btns, text="Yükle", command=self.yukle).pack(side="right", fill="x", expand=True, padx=(2, 0))

        top_container = ttk.Frame(self.frame_left, style="Card.TFrame")
        top_container.pack(side="top", fill="x")

        row1 = ttk.Frame(top_container, style="Card.TFrame")
        row1.pack(fill="x", pady=5)
        f_boyut = ttk.Frame(row1, style="Card.TFrame")
        f_boyut.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Label(f_boyut, text="Boyut:", style="SubHeader.TLabel").pack(anchor="w")
        self.cmb_boyut = ttk.Combobox(f_boyut, values=list(HAMUR_ARALIKLARI.keys()), state="readonly")
        self.cmb_boyut.current(1)
        self.cmb_boyut.pack(fill="x", pady=2)
        self.cmb_boyut.bind("<<ComboboxSelected>>", self.update_gramaj)

        f_gram = ttk.Frame(row1, style="Card.TFrame")
        f_gram.pack(side="right", fill="x", expand=True)
        ttk.Label(f_gram, text="Gramaj (g):", style="SubHeader.TLabel").pack(anchor="w")
        self.ent_gram = ttk.Entry(f_gram, state="readonly", font=("Segoe UI", 10, "bold")) 
        self.ent_gram.pack(fill="x", pady=2)

        ttk.Label(top_container, text="Hamur Tipi:", style="SubHeader.TLabel").pack(anchor="w", pady=(10, 5))
        f_tip = ttk.Frame(top_container, style="Card.TFrame")
        f_tip.pack(fill="x")
        self.var_hamur = tk.StringVar(value="ince")
        ttk.Radiobutton(f_tip, text="İnce Hamur", variable=self.var_hamur, value="ince", command=self.update_gramaj).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(f_tip, text="Kalın Hamur", variable=self.var_hamur, value="kalin", command=self.update_gramaj).pack(side="left")

        ttk.Label(top_container, text="Malzemeler:", style="SubHeader.TLabel").pack(anchor="w", pady=(15, 5))
        list_container = tk.Frame(self.frame_left, bg="#f0f0f0", bd=1, relief="sunken")
        list_container.pack(side="top", fill="both", expand=True, pady=(0, 10))
        canvas_list = tk.Canvas(list_container, bg=COLOR_BG_MAIN, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas_list.yview)
        self.scrollable_frame = ttk.Frame(canvas_list, style="Card.TFrame")
        self.scrollable_frame.bind("<Configure>", lambda e: canvas_list.configure(scrollregion=canvas_list.bbox("all")))
        
        def _on_mousewheel(event):
            canvas_list.yview_scroll(int(-1*(event.delta/120)), "units")
        list_container.bind('<Enter>', lambda e: canvas_list.bind_all("<MouseWheel>", _on_mousewheel))
        list_container.bind('<Leave>', lambda e: canvas_list.unbind_all("<MouseWheel>"))

        canvas_list.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas_list.configure(yscrollcommand=scrollbar.set)
        canvas_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.malzeme_vars = {}
        self.refresh_checkboxes()

        # --- SAĞ PANEL ---
        self.frame_right = ttk.Frame(self.canvas_bg, style="Card.TFrame", padding=25)
        self.win_right = self.canvas_bg.create_window(500, 70, window=self.frame_right, anchor="nw", width=540, height=650)

        ttk.Label(self.frame_right, text="MALİYET TABLOSU", style="Header.TLabel").pack(anchor="w", pady=(0, 10))

        # --- DÜZELTME 1: Liste yüksekliğini azalttım (height=10 yapıldı) ---
        cols = ("Malzeme", "Gram", "Fiyat")
        self.tree = ttk.Treeview(self.frame_right, columns=cols, show="headings", height=10) # Burası 15'ten 10'a düşürüldü
        style = ttk.Style()
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        self.tree.heading("Malzeme", text="Malzeme Adı")
        self.tree.heading("Gram", text="Gramaj (g)")
        self.tree.heading("Fiyat", text="Tutar (TL)")
        self.tree.column("Malzeme", width=200)
        self.tree.column("Gram", width=100, anchor="center")
        self.tree.column("Fiyat", width=100, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=(0, 15))

        # --- DÜZELTME 2: Sonuç Ekranı Genişletildi ve Görünür Yapıldı ---
        self.result_frame = tk.Label(self.frame_right, 
                                     text="Lütfen malzemeleri seçip\n'Maliyet Hesapla' butonuna basınız.", 
                                     font=("Consolas", 12), 
                                     bg="#f4f6f7", fg="#555", 
                                     justify="left", 
                                     padx=20, pady=20, 
                                     relief="groove", borderwidth=2) # Çerçeve eklendi
        # pack ayarları ile genişlemesini sağladık
        self.result_frame.pack(fill="x", pady=(0, 10), ipady=10) 
        
        info_txt = f"Hamur Birim Fiyatı: {self.hamur_fiyat} TL/100g"
        self.lbl_info = ttk.Label(self.frame_right, text=info_txt, font=("Segoe UI", 9), foreground="#888")
        self.lbl_info.pack(side="bottom", anchor="e")

        self.update_gramaj()

    def toggle_menu(self):
        if self.menu_acik:
            self.canvas_bg.itemconfigure(self.win_left, state='hidden')
            self.canvas_bg.coords(self.win_right, 280, 70)
            self.btn_menu.config(text="☰ AYARLAR")
            self.menu_acik = False
        else:
            self.canvas_bg.itemconfigure(self.win_left, state='normal')
            self.canvas_bg.coords(self.win_right, 500, 70)
            self.btn_menu.config(text="✕ GİZLE")
            self.menu_acik = True

    def refresh_checkboxes(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.malzeme_vars = {}
        for malz in self.malzemeler.keys():
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(self.scrollable_frame, text=malz, variable=var, cursor="hand2")
            chk.pack(anchor="w", pady=3, padx=5)
            self.malzeme_vars[malz] = var

    def update_gramaj(self, event=None):
        boyut = self.cmb_boyut.get()
        tip = self.var_hamur.get()
        if boyut in HAMUR_ARALIKLARI:
            min_g, max_g = HAMUR_ARALIKLARI[boyut][tip]
            yeni_gram = random.randint(min_g, max_g)
            self.ent_gram.config(state="normal")
            self.ent_gram.delete(0, tk.END)
            self.ent_gram.insert(0, str(yeni_gram))
            self.ent_gram.config(state="readonly")

    def hesapla(self):
        try:
            gram_hamur = float(self.ent_gram.get())
        except:
            messagebox.showerror("Hata", "Gramaj bilgisi okunamadı.")
            return

        for i in self.tree.get_children():
            self.tree.delete(i)

        toplam_gram = gram_hamur
        hamur_maliyet = (gram_hamur / 100) * self.hamur_fiyat
        toplam_fiyat = hamur_maliyet
        self.tree.insert("", "end", values=("Hamur (Taban)", f"{gram_hamur}", f"{hamur_maliyet:.2f}"))

        for malz, var in self.malzeme_vars.items():
            if var.get():
                data = self.malzemeler[malz]
                m_gram = gram_hamur * (data["oran"] / 100)
                m_fiyat = (m_gram / 100) * data["fiyat"]
                toplam_gram += m_gram
                toplam_fiyat += m_fiyat
                self.tree.insert("", "end", values=(malz, f"{m_gram:.2f}", f"{m_fiyat:.2f}"))

        carpan = 1 + (self.kar_orani / 100)
        satis_fiyati = toplam_fiyat * carpan
        net_kar = satis_fiyati - toplam_fiyat

        # --- DÜZELTME 3: Hizalama ve İçerik Gösterimi ---
        ozet_text = (
            f"Toplam Ağırlık : {toplam_gram:.2f} g\n"
            f"Toplam Maliyet : {toplam_fiyat:.2f} TL\n"
            f"-----------------------------------\n"
            f"SATIŞ FİYATI   : {satis_fiyati:.2f} TL\n"
            f"NET KÂR        : {net_kar:.2f} TL\n"
            f"KÂR MARJI      : %{self.kar_orani:.0f}"
        )
        
        # --- RENK: AÇIK TURKUAZ ---
        turkuaz_acik = "#AFEEEE"  # PaleTurquoise
        
        self.result_frame.config(
            text=ozet_text, 
            bg=turkuaz_acik, 
            fg="black", 
            font=("Consolas", 14, "bold") # Font biraz daha büyütüldü
        )

    def indir(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json"), ("Excel", "*.xlsx")])
        if not file_path: return
        if file_path.endswith(".json"):
            data = {
                "malzemeler": [{"Malzeme": k, "Oran": v["oran"], "Fiyat": v["fiyat"]} for k,v in self.malzemeler.items()],
                "ayarlar": {"hamur_fiyat": self.hamur_fiyat, "kar_orani": self.kar_orani}
            }
            with open(file_path, "w", encoding="utf-8") as f: json.dump(data, f, indent=4, ensure_ascii=False)
        else:
            wb = openpyxl.Workbook()
            ws1 = wb.active; ws1.title = "Malzemeler"; ws1.append(["Malzeme", "Oran", "Fiyat"])
            for k, v in self.malzemeler.items(): ws1.append([k, v["oran"], v["fiyat"]])
            ws2 = wb.create_sheet("Ayarlar"); ws2.append(["Ayar", "Deger"])
            ws2.append(["Hamur 100g Fiyatı", self.hamur_fiyat]); ws2.append(["Kar Oranı (%)", self.kar_orani])
            wb.save(file_path)
        messagebox.showinfo("Başarılı", "Dosya kaydedildi.")

    def yukle(self):
        file_path = filedialog.askopenfilename(filetypes=[("Veri Dosyası", "*.json *.xlsx")])
        if not file_path: return
        ext = os.path.splitext(file_path)[1]
        target = "pizza_data" + ext
        with open(file_path, "rb") as src, open(target, "wb") as dst: dst.write(src.read())
        self.load_data(); self.refresh_checkboxes(); self.update_gramaj()
        self.lbl_info.config(text=f"Hamur Birim Fiyatı: {self.hamur_fiyat} TL/100g")
        messagebox.showinfo("Başarılı", "Veriler güncellendi.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PizzaApp(root)
    root.mainloop()
