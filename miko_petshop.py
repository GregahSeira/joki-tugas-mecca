# Miko Petshop - Smart Inventory Management System (Tkinter + Matplotlib)
# Sumber data: "data miko petshop.xlsx" (hanya dibaca, file asli tidak diubah)
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date, timedelta
import os
import openpyxl
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ---------- Konstanta tema (mengacu pada desain Glacier) ----------
BG, SIDE, CARD, FIELD, LINE = "#040e21", "#071328", "#0c1931", "#111f39", "#1c2c4a"
PRI, TXT, MUT, DARK = "#7bd1fa", "#dce5ff", "#a1abc4", "#00465d"
OK, WARN, ERR = "#34d399", "#fbbf24", "#ff716c"
FONT = "Segoe UI"
WARNA = {"Aman": PRI, "Menipis": WARN, "Kritis": ERR, "Masuk": OK, "Keluar": ERR}
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
MENU = [("dashboard", "▦  Dashboard"), ("produk", "▤  Data Produk"), ("masuk", "⬇  Produk Masuk"),
        ("keluar", "⬆  Produk Keluar"), ("kategori", "◈  Kategori Produk"),
        ("alert", "⚠  Peringatan Stok"), ("laporan", "▥  Laporan & Visualisasi")]
SUBMENU = [("masuk_tambah", "Tambah Transaksi Produk Masuk"), ("masuk_riwayat", "Riwayat Produk Masuk")]
FIELD_PRODUK = [("kode", "Kode Produk"), ("nama", "Nama Produk"), ("kategori", "Kategori"),
                ("satuan", "Satuan"), ("awal", "Stok Awal"), ("min", "Stok Minimum"),
                ("harga", "Harga per Unit (Rp)"), ("lokasi", "Lokasi Penyimpanan")]
PERIODE = {"Seluruh Data": None, "7 Hari Terakhir": 7, "30 Hari Terakhir": 30}
FILE_EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data miko petshop.xlsx")

# ---------- Pemuatan & pembersihan data ----------
def bersihkan_tanggal(v):
    """Sebagian tanggal Excel terbaca terbalik (bulan/hari), koreksi ke tanggal asli."""
    if isinstance(v, datetime):
        return date(v.year, v.day, v.month)
    return datetime.strptime("".join(str(v).split()), "%d/%m/%Y").date()

def muat_data():
    """Baca Excel -> (dict produk, list transaksi, list kategori)."""
    ws = openpyxl.load_workbook(FILE_EXCEL, read_only=True, data_only=True).active
    produk, transaksi, kategori = {}, [], []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        kode, nama, kat, satuan, awal, mn, harga, tgl, jenis, jml, ket, lokasi = r[:12]
        kat = " ".join(kat.split())  # hilangkan spasi ganda
        if kat not in kategori:
            kategori.append(kat)
        if kode not in produk:  # stok awal = nilai pertama tiap produk
            produk[kode] = dict(kode=kode, nama=nama, kategori=kat, satuan=satuan, awal=int(awal),
                                min=int(mn), harga=int("".join(c for c in str(harga) if c.isdigit())),
                                lokasi=lokasi)
        transaksi.append(dict(tgl=bersihkan_tanggal(tgl), kode=kode, jenis=jenis, jml=int(jml),
                              lok=produk[kode]["lokasi"] if jenis == "Masuk" else "-", ket=ket))
    transaksi.sort(key=lambda t: t["tgl"])
    return dict(sorted(produk.items())), transaksi, kategori

def status(stok, mn):
    """Aman: stok > minimum, Menipis: stok == minimum, Kritis: stok < minimum."""
    return "Aman" if stok > mn else "Menipis" if stok == mn else "Kritis"

# ---------- Widget bantu ----------
def label(parent, teks, size=10, fg=TXT, bg=CARD, bold=False, **kw):
    return tk.Label(parent, text=teks, fg=fg, bg=bg, font=(FONT, size, "bold" if bold else "normal"), **kw)

def tombol(parent, teks, cmd, bg=PRI, fg=DARK, **kw):
    """Tombol berbasis Label supaya warna konsisten di semua OS."""
    b = label(parent, teks, 10, fg, bg, True, padx=16, pady=7, cursor="hand2", **kw)
    b.bind("<Button-1>", lambda e: cmd())
    b.bind("<Enter>", lambda e: b.config(bg=TXT if bg == PRI else LINE))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b

def kartu(parent, **kw):
    return tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1, **kw)

def entri(parent, var, lebar=30, **kw):
    return tk.Entry(parent, textvariable=var, width=lebar, bg=FIELD, fg=TXT, insertbackground=TXT, relief="flat",
                    highlightthickness=1, highlightbackground=LINE, highlightcolor=PRI, font=(FONT, 10), **kw)

def tabel(parent, kolom):
    """Treeview bergaya gelap; kolom = [(judul, lebar, rata)]."""
    box = tk.Frame(parent, bg=CARD)
    t = ttk.Treeview(box, columns=[k[0] for k in kolom], show="headings", height=8)
    for judul, lebar, rata in kolom:
        t.heading(judul, text=judul.upper())
        t.column(judul, width=lebar, anchor=rata)
    sb = ttk.Scrollbar(box, orient="vertical", command=t.yview)
    t.configure(yscrollcommand=sb.set)
    t.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    for tag, warna in WARNA.items():
        t.tag_configure(tag, foreground=warna)
    t.box = box
    return t

def isi_tabel(t, baris, tags=None):
    t.delete(*t.get_children())
    for i, b in enumerate(baris):
        t.insert("", "end", iid=str(b[0]) if tags == "iid" else None, values=b[1:] if tags == "iid" else b,
                 tags=(tags[i],) if tags and tags != "iid" else ())

class App:
    def __init__(self, root):
        self.root = root
        root.title("Miko Petshop")
        root.geometry("1320x780")
        root.minsize(1180, 700)
        root.configure(bg=BG)
        self.produk, self.transaksi, self.kategori = muat_data()
        self.atur_style()
        self.halaman_login()

    # ---------- Perhitungan (satu sumber data untuk semua menu) ----------
    def stok(self, kode):
        s = self.produk[kode]["awal"]
        for t in self.transaksi:
            if t["kode"] == kode:
                s += t["jml"] if t["jenis"] == "Masuk" else -t["jml"]
        return s

    def status_produk(self, kode):
        return status(self.stok(kode), self.produk[kode]["min"])

    def hitung(self, jenis, data=None):
        """Mengembalikan (jumlah transaksi, total unit) untuk jenis tertentu."""
        j = [t["jml"] for t in (data or self.transaksi) if t["jenis"] == jenis]
        return len(j), sum(j)

    def atur_style(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TXT, rowheight=30,
                    borderwidth=0, bordercolor=CARD, lightcolor=CARD, darkcolor=CARD, font=(FONT, 10))
        s.configure("Treeview.Heading", background=FIELD, foreground=MUT, relief="flat", font=(FONT, 9, "bold"))
        s.map("Treeview", background=[("selected", "#1c4c60")], foreground=[("selected", TXT)])
        s.configure("TCombobox", fieldbackground=FIELD, background=FIELD, foreground=TXT, arrowcolor=PRI,
                    bordercolor=LINE, lightcolor=FIELD, darkcolor=FIELD, padding=5)
        s.map("TCombobox", fieldbackground=[("readonly", FIELD)], foreground=[("readonly", TXT)],
              selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", TXT)])
        s.configure("TScrollbar", background=FIELD, troughcolor=CARD, bordercolor=CARD, lightcolor=FIELD, darkcolor=FIELD, arrowcolor=MUT)
        for k, v in (("background", FIELD), ("foreground", TXT), ("selectBackground", "#1c4c60")):
            self.root.option_add("*TCombobox*Listbox." + k, v)

    # ---------- Login ----------
    def halaman_login(self):
        self.login = tk.Frame(self.root, bg=BG)
        self.login.pack(fill="both", expand=True)
        box = kartu(self.login, padx=44, pady=32)
        box.place(relx=0.5, rely=0.5, anchor="center")
        label(box, "🐾 MIKO PET SHOP", 20, TXT, bold=True).pack()
        label(box, "SISTEM MANAJEMEN INVENTORI", 9, PRI, bold=True).pack(pady=(2, 6))
        self.var_user, self.var_pass = tk.StringVar(), tk.StringVar()
        label(box, "USERNAME", 9, TXT, bold=True).pack(anchor="w")
        entri(box, self.var_user, 34).pack(fill="x", ipady=7, pady=(3, 10))
        label(box, "PASSWORD", 9, TXT, bold=True).pack(anchor="w")
        baris = tk.Frame(box, bg=FIELD, highlightthickness=1, highlightbackground=LINE)
        baris.pack(fill="x", pady=(3, 6))
        self.ent_pass = tk.Entry(baris, textvariable=self.var_pass, show="•", bg=FIELD, fg=TXT, relief="flat",
                                 insertbackground=TXT, font=(FONT, 10))
        self.ent_pass.pack(side="left", fill="x", expand=True, ipady=7, padx=6)
        mata = label(baris, "👁", 11, PRI, FIELD, cursor="hand2", padx=8)
        mata.pack(side="right")
        mata.bind("<Button-1>", lambda e: self.ent_pass.config(show="" if self.ent_pass.cget("show") else "•"))
        self.lbl_err = label(box, "", 9, ERR)
        self.lbl_err.pack(anchor="w", pady=(0, 8))
        tombol(box, "MASUK KE DASHBOARD  →", self.cek_login).pack(fill="x")
        label(box, "Akun uji: admin / admin123", 8, MUT).pack(pady=(14, 0))
        self.root.bind("<Return>", lambda e: self.cek_login() if self.login.winfo_exists() else None)

    def cek_login(self):
        if self.var_user.get() == "admin" and self.var_pass.get() == "admin123":
            self.login.destroy()
            self.bangun_utama()
        else:
            self.lbl_err.config(text="⚠ Username atau password salah.")
            self.var_pass.set("")

    # ---------- Kerangka utama: sidebar, header, jam ----------
    def bangun_utama(self):
        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(fill="both", expand=True)
        side = tk.Frame(self.main, bg=SIDE, width=260, highlightbackground=LINE, highlightthickness=1)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        label(side, "🐾 MIKO PET SHOP", 14, TXT, SIDE, True).pack(anchor="w", padx=18, pady=(20, 0))
        label(side, "INVENTORY SYSTEM\n", 8, PRI, SIDE, True).pack(anchor="w", padx=18)
        self.menu = tk.Frame(side, bg=SIDE)
        self.menu.pack(fill="x", padx=10)
        bawah = tk.Frame(side, bg=SIDE)
        bawah.pack(side="bottom", fill="x", padx=16, pady=16)
        label(bawah, "Admin Miko\nSistem Akun", 9, TXT, SIDE, True, justify="left").pack(side="left")
        out = label(bawah, "⏻ Keluar", 9, ERR, SIDE, True, cursor="hand2")
        out.pack(side="right")
        out.bind("<Button-1>", lambda e: self.logout())
        kanan = tk.Frame(self.main, bg=BG)
        kanan.pack(side="left", fill="both", expand=True)
        head = tk.Frame(kanan, bg=BG, height=64, highlightbackground=LINE, highlightthickness=1)
        head.pack(fill="x")
        head.pack_propagate(False)
        label(head, "MIKO PET SHOP  ·  SISTEM MANAJEMEN INVENTARIS", 10, PRI, BG, True).pack(side="left", padx=28)
        pil = tk.Frame(head, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        pil.pack(side="right", padx=28)
        self.lbl_jam = label(pil, "", 9, PRI, CARD, True, padx=14, pady=6)
        self.lbl_jam.pack()
        self.content = tk.Frame(kanan, bg=BG, padx=28, pady=18)
        self.content.pack(fill="both", expand=True)
        self.tik()
        self.buka("dashboard")

    def tik(self):
        """Jam real-time, diperbarui tiap 1 detik."""
        n = datetime.now()
        self.lbl_jam.config(text=f"📅 {n.day} {BULAN[n.month - 1].upper()} {n.year}  •  {n:%H:%M:%S}")
        self.job = self.root.after(1000, self.tik)

    def logout(self):
        self.root.after_cancel(self.job)
        self.main.destroy()
        self.halaman_login()

    def buka(self, key):
        """Navigasi: gambar ulang sidebar dan halaman agar selalu memakai data terbaru."""
        self.aktif = key
        for w in self.menu.winfo_children():
            w.destroy()
        n_alert = sum(1 for k in self.produk if self.status_produk(k) != "Aman")
        for k, teks in MENU:
            on = k == key or (k == "masuk" and key.startswith("masuk"))
            if k == "alert" and n_alert:
                teks += f"   ({n_alert})"
            m = label(self.menu, teks + ("  ▾" if k == "masuk" else ""), 10, DARK if on else TXT, PRI if on else SIDE,
                      True, anchor="w", padx=14, pady=9, cursor="hand2")
            m.pack(fill="x", pady=2)
            m.bind("<Button-1>", lambda e, k=k: self.buka("masuk_tambah" if k == "masuk" else k))
            if k == "masuk" and on:  # dropdown submenu Produk Masuk
                for sk, st in SUBMENU:
                    s = label(self.menu, "•  " + st, 9, PRI if sk == key else MUT, SIDE, sk == key, anchor="w",
                              padx=30, pady=5, cursor="hand2")
                    s.pack(fill="x")
                    s.bind("<Button-1>", lambda e, sk=sk: self.buka(sk))
        for w in self.content.winfo_children():
            w.destroy()
        getattr(self, "hal_" + key)()

    def judul(self, teks, sub):
        h = tk.Frame(self.content, bg=BG)
        h.pack(fill="x", pady=(0, 14))
        kiri = tk.Frame(h, bg=BG)
        kiri.pack(side="left")
        label(kiri, teks, 22, TXT, BG, True).pack(anchor="w")
        label(kiri, sub, 10, MUT, BG).pack(anchor="w")
        return h

    def deretan(self, daftar, f=None):
        """Baris kartu ringkasan: daftar = [(judul, nilai, sub, warna)]."""
        if f is None:
            f = tk.Frame(self.content, bg=BG)
            f.pack(fill="x", pady=(0, 14))
        for i, (judul, nilai, sub, warna) in enumerate(daftar):
            k = kartu(f, padx=16, pady=12)
            k.pack(side="left", fill="x", expand=True, padx=(0 if i == 0 else 12, 0))
            label(k, judul, 9, MUT, bold=True).pack(anchor="w")
            label(k, str(nilai), 26, TXT, bold=True).pack(anchor="w")
            label(k, sub, 9, warna, bold=True).pack(anchor="w")

    # ---------- A. Dashboard ----------
    def hal_dashboard(self):
        self.judul("Dashboard Overview", "Ringkasan persediaan produk Miko Pet Shop")
        butuh = [k for k in self.produk if self.status_produk(k) != "Aman"]
        (nm, um), (nk, uk) = self.hitung("Masuk"), self.hitung("Keluar")
        self.deretan([("JENIS PRODUK", len(self.produk), "SKU terdaftar", PRI),
                      ("TOTAL STOK", sum(self.stok(k) for k in self.produk), "unit tersedia", PRI),
                      ("STOK MENIPIS", len(butuh), "produk perlu dicek", WARN),
                      ("TRANSAKSI MASUK", nm, f"+{um} unit", OK), ("TRANSAKSI KELUAR", nk, f"-{uk} unit", ERR)])
        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=3)
        grid.columnconfigure(1, weight=2)
        grid.rowconfigure(0, weight=1)
        kiri, kanan = kartu(grid, padx=14, pady=12), kartu(grid, padx=14, pady=12)
        kiri.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        kanan.grid(row=0, column=1, sticky="nsew")
        label(kiri, "Aktivitas Inventaris Terbaru", 13, bold=True).pack(anchor="w", pady=(0, 8))
        t = tabel(kiri, [("Tanggal", 90, "w"), ("Kode", 60, "w"), ("Produk", 130, "w"), ("Jenis", 60, "center"),
                         ("Jumlah", 60, "center"), ("Keterangan", 110, "w")])
        t.box.pack(fill="both", expand=True)
        terbaru = sorted(self.transaksi, key=lambda x: x["tgl"], reverse=True)[:12]
        isi_tabel(t, [(x["tgl"].strftime("%d/%m/%Y"), x["kode"], self.produk[x["kode"]]["nama"], x["jenis"],
                       x["jml"], x["ket"]) for x in terbaru], [x["jenis"] for x in terbaru])
        label(kanan, "Produk Perlu Diperiksa", 13, bold=True).pack(anchor="w", pady=(0, 8))
        t2 = tabel(kanan, [("Kode", 60, "w"), ("Produk", 120, "w"), ("Stok", 50, "center"), ("Min", 50, "center"),
                           ("Status", 70, "center")])
        t2.box.pack(fill="both", expand=True)
        isi_tabel(t2, [(k, self.produk[k]["nama"], self.stok(k), self.produk[k]["min"], self.status_produk(k))
                       for k in butuh], [self.status_produk(k) for k in butuh])

    # ---------- B. Data Produk ----------
    def hal_produk(self):
        self.judul("Data Produk", "Kelola daftar produk, stok, dan lokasi penyimpanan.")
        bar = tk.Frame(self.content, bg=BG)
        bar.pack(fill="x", pady=(0, 10))
        self.cari = tk.StringVar()
        label(bar, "🔍 Cari:", 10, MUT, BG).pack(side="left")
        entri(bar, self.cari, 34).pack(side="left", ipady=6, padx=8)
        self.cari.trace_add("write", lambda *a: self.isi_produk())
        tombol(bar, "Hapus", lambda: self.pilih_produk(self.hapus_produk), LINE, ERR).pack(side="right", padx=(8, 0))
        tombol(bar, "Edit", lambda: self.pilih_produk(self.form_produk), LINE, PRI).pack(side="right", padx=(8, 0))
        tombol(bar, "+ Tambah", lambda: self.form_produk()).pack(side="right")
        k = kartu(self.content, padx=10, pady=10)
        k.pack(fill="both", expand=True)
        self.tb_produk = tabel(k, [("Kode", 80, "w"), ("Nama Produk", 200, "w"), ("Kategori", 150, "w"),
                                   ("Stok", 70, "center"), ("Satuan", 90, "center"), ("Lokasi", 110, "center"),
                                   ("Min", 80, "center"), ("Status", 90, "center")])
        self.tb_produk.box.pack(fill="both", expand=True)
        self.isi_produk()

    def isi_produk(self):
        q = self.cari.get().lower()
        data = [(k, k, p["nama"], p["kategori"], self.stok(k), p["satuan"], p["lokasi"], p["min"], self.status_produk(k))
                for k, p in self.produk.items() if q in (k + p["nama"] + p["kategori"]).lower()]
        isi_tabel(self.tb_produk, data, "iid")

    def pilih_produk(self, aksi):
        """Jalankan aksi untuk produk yang dipilih pada tabel."""
        sel = self.tb_produk.selection()
        if sel:
            aksi(sel[0])
        else:
            messagebox.showinfo("Pilih Produk", "Pilih salah satu produk pada tabel terlebih dahulu.")

    def form_produk(self, kode=None):
        p = self.produk.get(kode, {})
        win = tk.Toplevel(self.root)
        win.title("Edit Produk" if p else "Tambah Produk")
        win.configure(bg=CARD, padx=24, pady=18)
        win.grab_set()
        isi = {}
        for i, (k, teks) in enumerate(FIELD_PRODUK):
            label(win, teks, 9, MUT).grid(row=i, column=0, sticky="w", pady=5)
            isi[k] = tk.StringVar(value=str(p.get(k, "")))
            if k == "kategori":
                w = ttk.Combobox(win, textvariable=isi[k], values=self.kategori, state="readonly", width=28)
            else:
                w = entri(win, isi[k], 31, **({"state": "disabled", "disabledbackground": LINE} if p and k == "kode" else {}))
            w.grid(row=i, column=1, padx=(14, 0), pady=5, ipady=3)

        def simpan():
            d = {k: v.get().strip() for k, v in isi.items()}
            if not all(d.values()):
                return messagebox.showerror("Gagal", "Semua kolom wajib diisi.", parent=win)
            try:
                angka = [int(d[k]) for k in ("awal", "min", "harga")]
            except ValueError:
                return messagebox.showerror("Gagal", "Stok awal, minimum, dan harga harus berupa angka.", parent=win)
            if min(angka) < 0:
                return messagebox.showerror("Gagal", "Angka tidak boleh negatif.", parent=win)
            kd = kode or d["kode"].upper()
            if not p and kd in self.produk:
                return messagebox.showerror("Gagal", "Kode produk sudah terdaftar.", parent=win)
            self.produk[kd] = dict(d, kode=kd, awal=angka[0], min=angka[1], harga=angka[2])
            self.produk = dict(sorted(self.produk.items()))
            win.destroy()
            self.buka("produk")
        tombol(win, "Simpan", simpan).grid(row=len(FIELD_PRODUK), column=0, columnspan=2, sticky="ew", pady=(14, 0))

    def hapus_produk(self, kode):
        if messagebox.askyesno("Hapus Produk", f"Hapus {kode} beserta seluruh riwayat transaksinya?"):
            del self.produk[kode]
            self.transaksi = [t for t in self.transaksi if t["kode"] != kode]
            self.buka("produk")

    # ---------- C & D. Produk Masuk / Keluar ----------
    def hal_masuk_tambah(self):
        self.judul("Tambah Transaksi Produk Masuk", "Catat penerimaan persediaan dari supplier atau retur.")
        self.form_transaksi(self.content, "Masuk").pack(anchor="nw")

    def hal_masuk_riwayat(self):
        self.judul("Riwayat Produk Masuk", "Seluruh transaksi barang masuk, terbaru di atas.")
        self.riwayat(self.content, "Masuk").pack(fill="both", expand=True)

    def hal_keluar(self):
        self.judul("Produk Keluar", "Catat pemakaian atau pengeluaran barang. Stok berkurang otomatis.")
        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="both", expand=True)
        self.form_transaksi(grid, "Keluar").pack(side="left", anchor="n", padx=(0, 14))
        self.riwayat(grid, "Keluar").pack(side="left", fill="both", expand=True)

    def form_transaksi(self, parent, jenis):
        masuk = jenis == "Masuk"
        f = kartu(parent, padx=22, pady=18)
        v = {k: tk.StringVar() for k in ("produk", "tgl", "jml", "lok", "ket")}
        v["tgl"].set(date.today().strftime("%d/%m/%Y"))
        label(f, f"Form Produk {jenis}", 14, bold=True).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        info = label(f, "Pilih produk untuk melihat info stok.", 9, MUT, justify="left")

        def pilih(_=None):
            kd = v["produk"].get().split(" - ")[0]
            p, st = self.produk[kd], self.status_produk(kd)
            info.config(text=f"Stok saat ini: {self.stok(kd)} ({p['satuan']})   Minimum: {p['min']}\n"
                             f"Harga/unit: Rp{p['harga']:,}   Status: {st}", fg=WARNA[st])
            if masuk:
                v["lok"].set(p["lokasi"])
        kolom = [("Produk *", "produk"), ("Tanggal (dd/mm/yyyy) *", "tgl"), ("Jumlah *", "jml"),
                 ("Lokasi Penyimpanan *" if masuk else "Tujuan / Lokasi Penggunaan *", "lok"), ("Keterangan", "ket")]
        for i, (teks, k) in enumerate(kolom, 1):
            label(f, teks, 9, MUT).grid(row=i, column=0, sticky="w", pady=6)
            if k == "produk":
                w = ttk.Combobox(f, textvariable=v[k], state="readonly", width=22,
                                 values=[f"{kd} - {p['nama']}" for kd, p in self.produk.items()])
                w.bind("<<ComboboxSelected>>", pilih)
            else:
                w = entri(f, v[k], 24)
            w.grid(row=i, column=1, padx=(16, 0), ipady=3)
        info.grid(row=6, column=0, columnspan=2, sticky="w", pady=8)

        def simpan():
            kd = v["produk"].get().split(" - ")[0]
            if kd not in self.produk:
                return messagebox.showerror("Gagal", "Pilih produk terlebih dahulu.")
            try:
                tgl, jml = datetime.strptime(v["tgl"].get().strip(), "%d/%m/%Y").date(), int(v["jml"].get())
            except ValueError:
                return messagebox.showerror("Gagal", "Tanggal harus dd/mm/yyyy dan jumlah harus berupa angka.")
            lok = v["lok"].get().strip()
            if jml <= 0 or not lok:
                return messagebox.showerror("Gagal", "Jumlah harus lebih dari 0 dan lokasi/tujuan wajib diisi.")
            p = self.produk[kd]
            if not masuk and jml > self.stok(kd):  # tolak pengeluaran melebihi stok
                return messagebox.showerror("Stok Tidak Cukup", f"Pengeluaran ditolak. Stok tersedia hanya {self.stok(kd)} {p['satuan']}.")
            self.transaksi.append(dict(tgl=tgl, kode=kd, jenis=jenis, jml=jml, lok=lok, ket=v["ket"].get().strip() or "-"))
            self.transaksi.sort(key=lambda t: t["tgl"])
            if masuk:
                p["lokasi"] = lok
            st = self.status_produk(kd)
            messagebox.showinfo("Berhasil", f"Transaksi {jenis.lower()} tersimpan. Stok {p['nama']} sekarang {self.stok(kd)}.")
            if st != "Aman":
                messagebox.showwarning("Peringatan Stok", f"{p['nama']} berstatus {st.upper()} (stok {self.stok(kd)}, minimum {p['min']}).")
            self.buka(self.aktif)
        tombol(f, "Simpan Transaksi", simpan).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        return f

    def riwayat(self, parent, jenis):
        k = kartu(parent, padx=14, pady=12)
        label(k, f"Riwayat Produk {jenis}", 14, bold=True).pack(anchor="w", pady=(0, 8))
        t = tabel(k, [("No", 30, "center"), ("Tanggal", 88, "w"), ("Kode", 52, "w"), ("Produk", 100, "w"),
                      ("Jml", 40, "center"), ("Lokasi", 48, "w"), ("Keterangan", 80, "w")])
        t.box.pack(fill="both", expand=True)
        data = [x for x in sorted(self.transaksi, key=lambda x: x["tgl"], reverse=True) if x["jenis"] == jenis]
        isi_tabel(t, [(i, x["tgl"].strftime("%d/%m/%Y"), x["kode"], self.produk[x["kode"]]["nama"], x["jml"],
                       x["lok"], x["ket"]) for i, x in enumerate(data, 1)],
                  [jenis] * len(data))
        return k

    # ---------- E. Kategori Produk ----------
    def hal_kategori(self):
        self.judul("Kategori Produk", "Kategori terhubung langsung dengan Data Produk.")
        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="both", expand=True)
        f = kartu(grid, padx=20, pady=18)
        f.pack(side="left", anchor="n", padx=(0, 14))
        label(f, "Nama Kategori", 9, MUT).pack(anchor="w")
        var = tk.StringVar()
        entri(f, var, 30).pack(ipady=5, pady=(4, 12))
        k = kartu(grid, padx=10, pady=10)
        k.pack(side="left", fill="both", expand=True)
        t = tabel(k, [("Kategori", 260, "w"), ("Jumlah Produk", 140, "center")])
        t.box.pack(fill="both", expand=True)
        isi_tabel(t, [(c, c, sum(1 for p in self.produk.values() if p["kategori"] == c)) for c in self.kategori], "iid")
        t.bind("<<TreeviewSelect>>", lambda e: var.set(t.selection()[0]) if t.selection() else None)

        def tambah(lama=None):
            n = " ".join(var.get().split())
            if not n or (n in self.kategori and n != lama):
                return messagebox.showerror("Gagal", "Nama kategori kosong atau sudah ada.")
            if lama is None:
                self.kategori.append(n)
                self.buka("kategori")
            return n

        def edit():
            sel = t.selection()
            if not sel:
                return messagebox.showinfo("Pilih Kategori", "Pilih kategori pada tabel terlebih dahulu.")
            n = tambah(sel[0])
            if n:
                self.kategori[self.kategori.index(sel[0])] = n
                for p in self.produk.values():  # perubahan nama ikut ke Data Produk
                    if p["kategori"] == sel[0]:
                        p["kategori"] = n
                self.buka("kategori")

        def hapus():
            sel = t.selection()
            if not sel:
                return messagebox.showinfo("Pilih Kategori", "Pilih kategori pada tabel terlebih dahulu.")
            pakai = sum(1 for p in self.produk.values() if p["kategori"] == sel[0])
            if pakai:
                messagebox.showerror("Gagal", f"Kategori masih dipakai {pakai} produk.")
            elif messagebox.askyesno("Hapus Kategori", f"Hapus kategori {sel[0]}?"):
                self.kategori.remove(sel[0])
                self.buka("kategori")
        tombol(f, "+ Tambah", lambda: tambah()).pack(fill="x", pady=3)
        tombol(f, "Simpan Perubahan", edit, LINE, PRI).pack(fill="x", pady=3)
        tombol(f, "Hapus", hapus, LINE, ERR).pack(fill="x", pady=3)

    # ---------- F. Stok Alert ----------
    def hal_alert(self):
        self.judul("Peringatan Stok", "Status diperbarui otomatis setiap terjadi perubahan stok.")
        hit = {s: sum(1 for k in self.produk if self.status_produk(k) == s) for s in WARNA if s in ("Aman", "Menipis", "Kritis")}
        self.deretan([("AMAN", hit["Aman"], "stok di atas minimum", PRI), ("MENIPIS", hit["Menipis"], "stok sama dengan minimum", WARN),
                      ("KRITIS", hit["Kritis"], "stok di bawah minimum", ERR)])
        k = kartu(self.content, padx=10, pady=10)
        k.pack(fill="both", expand=True)
        t = tabel(k, [("Kode", 80, "w"), ("Nama Produk", 200, "w"), ("Kategori", 150, "w"), ("Stok", 80, "center"),
                      ("Minimum", 80, "center"), ("Selisih", 80, "center"), ("Status", 100, "center")])
        t.box.pack(fill="both", expand=True)
        urut = sorted(self.produk, key=lambda k: ["Kritis", "Menipis", "Aman"].index(self.status_produk(k)))
        isi_tabel(t, [(k, self.produk[k]["nama"], self.produk[k]["kategori"], self.stok(k), self.produk[k]["min"],
                       self.stok(k) - self.produk[k]["min"], self.status_produk(k).upper()) for k in urut],
                  [self.status_produk(k) for k in urut])

    # ---------- G. Laporan Visualisasi ----------
    def hal_laporan(self):
        h = self.judul("Laporan & Visualisasi", "Dihitung dari riwayat transaksi barang masuk dan keluar.")
        self.var_per = tk.StringVar(value="Seluruh Data")
        cb = ttk.Combobox(h, textvariable=self.var_per, values=list(PERIODE), state="readonly", width=20)
        cb.pack(side="right")
        label(h, "Periode:", 10, MUT, BG).pack(side="right", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self.gambar())
        self.ringkas = tk.Frame(self.content, bg=BG)
        self.ringkas.pack(fill="x", pady=(0, 10))
        self.area = kartu(self.content)
        self.area.pack(fill="both", expand=True)
        self.gambar()

    def gambar(self):
        akhir = max((t["tgl"] for t in self.transaksi), default=date.today())  # acuan: transaksi terakhir
        hari = PERIODE[self.var_per.get()]
        awal = akhir - timedelta(days=hari - 1) if hari else min((t["tgl"] for t in self.transaksi), default=akhir)
        data = [t for t in self.transaksi if awal <= t["tgl"] <= akhir]
        (nm, um), (nk, uk) = self.hitung("Masuk", data), self.hitung("Keluar", data)
        for w in self.ringkas.winfo_children() + self.area.winfo_children():
            w.destroy()
        self.deretan([("TOTAL BARANG MASUK", f"{um} unit", f"{nm} transaksi", OK),
                      ("TOTAL BARANG KELUAR", f"{uk} unit", f"{nk} transaksi", ERR),
                      ("SELISIH (MASUK - KELUAR)", f"{um - uk:+d} unit", f"{awal:%d/%m/%Y} s/d {akhir:%d/%m/%Y}", PRI)], self.ringkas)
        fig = Figure(figsize=(10, 5.4), facecolor=CARD, layout="constrained")
        gs = fig.add_gridspec(2, 2, width_ratios=[2, 1])
        a1, a2, a3 = fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[0, 1]), fig.add_subplot(gs[1, :])
        for a, judul in ((a1, "Barang Masuk vs Keluar per Produk"), (a2, "Proporsi Masuk & Keluar"),
                         (a3, "Perubahan Total Stok Selama Periode")):
            a.set_facecolor(CARD)
            a.set_title(judul, color=TXT, fontsize=10, fontweight="bold")
            a.tick_params(colors=MUT, labelsize=8)
            for sp in a.spines.values():
                sp.set_color(LINE)
        kode = list(self.produk)
        masuk = [sum(t["jml"] for t in data if t["kode"] == k and t["jenis"] == "Masuk") for k in kode]
        keluar = [sum(t["jml"] for t in data if t["kode"] == k and t["jenis"] == "Keluar") for k in kode]
        x = range(len(kode))
        a1.bar([i - 0.2 for i in x], masuk, 0.4, color=OK, label="Masuk")
        a1.bar([i + 0.2 for i in x], keluar, 0.4, color=ERR, label="Keluar")
        a1.set_xticks(list(x), kode)
        a1.legend(facecolor=CARD, edgecolor=LINE, labelcolor=TXT, fontsize=8)
        if um + uk:
            a2.pie([um, uk], labels=["Masuk", "Keluar"], colors=[OK, ERR], autopct="%1.1f%%",
                   textprops={"color": TXT, "fontsize": 9}, wedgeprops={"edgecolor": CARD})
        else:
            a2.text(0.5, 0.5, "Belum ada data", color=MUT, ha="center", transform=a2.transAxes)
        total, per_hari = sum(p["awal"] for p in self.produk.values()), {}
        for t in self.transaksi:  # total stok harian = stok awal + akumulasi mutasi
            per_hari[t["tgl"]] = per_hari.get(t["tgl"], 0) + (t["jml"] if t["jenis"] == "Masuk" else -t["jml"])
        d, xs, ys = min(per_hari, default=akhir), [], []
        while d <= akhir:
            total += per_hari.get(d, 0)
            if d >= awal:
                xs.append(d.strftime("%d/%m"))
                ys.append(total)
            d += timedelta(days=1)
        a3.plot(xs, ys, color=PRI, marker="o", markersize=4)
        a3.fill_between(xs, ys, min(ys, default=0), color=PRI, alpha=0.15)
        a3.set_xticks(range(0, len(xs), max(1, len(xs) // 12)))
        a3.set_ylabel("Total stok (unit)", color=MUT, fontsize=8)
        a3.grid(color=LINE, alpha=0.5)
        kanvas = FigureCanvasTkAgg(fig, self.area)
        kanvas.get_tk_widget().pack(fill="both", expand=True)
        kanvas.draw()

if __name__ == "__main__":
    jendela = tk.Tk()
    App(jendela)
    jendela.mainloop()
