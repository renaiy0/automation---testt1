#!/usr/bin/env python3
"""
hitung_excel.py - Import Excel (data dasar tanpa total) -> SQLite -> hitung total otomatis -> Excel hasil.

Pemakaian:
    python hitung_excel.py file.xlsx
    python hitung_excel.py file.xlsx --qty "Jumlah/QTY" --harga "Harga Barang"   # kalau auto-deteksi salah
    python hitung_excel.py folder_berisi_excel/                                  # proses semua .xlsx

Yang dilakukan:
  1. Cari baris header otomatis (tidak harus di baris 1).
  2. Buang kolom total lama (kalau ada) supaya dihitung ulang dengan benar.
  3. Simpan ke SQLite (tabel + view total + view grand total).
  4. Tulis Excel hasil: kolom TOTAL HARGA berisi RUMUS (=QTY*HARGA) + baris TOTAL KESELURUHAN.
"""
import argparse
import re
import sqlite3
import sys
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Kata kunci untuk mendeteksi kolom (huruf kecil). Tambah sendiri kalau perlu.
POLA_QTY = r"qty|jumlah|kuantitas|banyak|volume"
POLA_HARGA = r"harga|price|tarif"
POLA_TOTAL = r"total|subtotal|sub total"          # kolom yang dianggap hasil hitungan lama
POLA_NO = r"^no\.?$|^nomor$"


def bersih(nama):
    return re.sub(r"\W+", "_", str(nama).strip().lower()).strip("_")


def cari_header(ws):
    """Baris pertama yang punya >= 3 sel teks = header."""
    for row in ws.iter_rows():
        teks = [c for c in row if isinstance(c.value, str) and c.value.strip()]
        if len(teks) >= 3:
            return row[0].row, [c.column for c in teks]
    raise ValueError("Header tidak ditemukan")


def baca_excel(path, sheet=None):
    wb = openpyxl.load_workbook(path)  # rumus tidak dibaca nilainya; kolom total memang dibuang
    ws = wb[sheet] if sheet else wb.worksheets[0]
    baris_h, kolom = cari_header(ws)
    header = {c: str(ws.cell(baris_h, c).value).strip() for c in kolom}

    # Buang kolom total lama
    header = {c: h for c, h in header.items() if not re.search(POLA_TOTAL, h.lower())}

    data = []
    for r in range(baris_h + 1, ws.max_row + 1):
        nilai = {h: ws.cell(r, c).value for c, h in header.items()}
        if all(v in (None, "") for v in nilai.values()):
            continue  # baris kosong
        data.append(nilai)
    df = pd.DataFrame(data, columns=list(header.values()))
    return df, ws.title


def deteksi(df, pola, tebakan=None, tolak=None):
    if tebakan:
        if tebakan not in df.columns:
            raise ValueError(f"Kolom '{tebakan}' tidak ada. Kolom tersedia: {list(df.columns)}")
        return tebakan
    for c in df.columns:
        if re.search(pola, c.lower()) and not (tolak and re.search(tolak, c.lower())):
            return c
    return None


def proses(path, db, out_dir, qty=None, harga=None):
    path = Path(path)
    df, nama_sheet = baca_excel(path)

    kol_qty = deteksi(df, POLA_QTY, qty)
    kol_harga = deteksi(df, POLA_HARGA, harga, tolak=POLA_TOTAL)
    if not kol_qty or not kol_harga:
        raise ValueError(f"Kolom qty/harga tidak terdeteksi. Kolom ditemukan: {list(df.columns)}. "
                         f"Tambahkan --qty 'NAMA' --harga 'NAMA' (di bot: tulis di caption file, "
                         f"contoh: qty=Jumlah harga=Harga Satuan)")

    # Nomor urut otomatis diperbaiki (kalau ada kolom No.)
    for c in df.columns:
        if re.search(POLA_NO, c.lower()):
            df[c] = range(1, len(df) + 1)

    for c in (kol_qty, kol_harga):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # ---------- SQLite ----------
    tabel = bersih(path.stem)
    q, h = bersih(kol_qty), bersih(kol_harga)
    d = df.copy()
    d.columns = [bersih(c) for c in d.columns]
    with sqlite3.connect(db) as con:
        d.to_sql(tabel, con, if_exists="replace", index=False)
        con.executescript(f"""
            DROP VIEW IF EXISTS "{tabel}_hitung";
            DROP VIEW IF EXISTS "{tabel}_grand_total";
            CREATE VIEW "{tabel}_hitung" AS
                SELECT *, "{q}" * "{h}" AS total_harga FROM "{tabel}";
            CREATE VIEW "{tabel}_grand_total" AS
                SELECT SUM("{q}") AS total_qty, SUM("{q}" * "{h}") AS grand_total FROM "{tabel}";
        """)
        grand = con.execute(f"SELECT grand_total FROM \"{tabel}_grand_total\"").fetchone()[0]

    # ---------- Excel hasil ----------
    out = Path(out_dir) / f"{path.stem}_HASIL.xlsx"
    tulis_excel(df, kol_qty, kol_harga, out)
    print(f"[OK] {path.name}: {len(df)} baris | qty='{kol_qty}' harga='{kol_harga}'")
    print(f"     DB   : {db} (tabel {tabel}, view {tabel}_hitung, {tabel}_grand_total)")
    print(f"     Excel: {out}")
    print(f"     TOTAL KESELURUHAN = Rp {grand:,.0f}".replace(",", "."))
    return {"file": out, "baris": len(df), "qty": kol_qty, "harga": kol_harga,
            "total_qty": float(df[kol_qty].sum()), "grand_total": float(grand)}


def tulis_excel(df, kol_qty, kol_harga, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rekap"
    cols = list(df.columns) + ["TOTAL HARGA"]
    tipis = Side(style="thin", color="999999")
    kotak = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)

    for j, c in enumerate(cols, 1):
        cell = ws.cell(1, j, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = kotak

    iq = list(df.columns).index(kol_qty) + 1
    ih = list(df.columns).index(kol_harga) + 1
    LQ, LH = get_column_letter(iq), get_column_letter(ih)
    LT = get_column_letter(len(cols))
    uang = '"Rp" #,##0'

    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            cell = ws.cell(i, j, v)
            cell.border = kotak
            cell.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str))
        ws.cell(i, ih).number_format = uang
        t = ws.cell(i, len(cols), f"={LQ}{i}*{LH}{i}")   # rumus hidup, bukan angka mati
        t.number_format = uang
        t.border = kotak
        t.alignment = Alignment(vertical="top")

    akhir = len(df) + 1
    n = akhir + 1
    ws.cell(n, 1, "TOTAL KESELURUHAN")
    ws.cell(n, iq, f"=SUM({LQ}2:{LQ}{akhir})")
    ws.cell(n, len(cols), f"=SUM({LT}2:{LT}{akhir})").number_format = uang
    for j in range(1, len(cols) + 1):
        c = ws.cell(n, j)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFF2CC")
        c.border = kotak
    ws.cell(n, ih).value = None

    # lebar kolom otomatis (teks panjang dibatasi + wrap)
    for j, c in enumerate(cols, 1):
        panjang = max([len(str(c))] + [max(len(s) for s in str(v).split("\n")) for v in
                       (df[c] if c in df.columns else [])] or [10])
        ws.column_dimensions[get_column_letter(j)].width = min(max(panjang + 2, 12), 55)
    ws.column_dimensions[LT].width = 20
    ws.freeze_panes = "A2"
    wb.save(out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="file .xlsx atau folder berisi .xlsx")
    ap.add_argument("--db", default="data.db")
    ap.add_argument("--out", default=".", help="folder output Excel hasil")
    ap.add_argument("--qty", help="nama kolom jumlah (opsional, kalau auto-deteksi salah)")
    ap.add_argument("--harga", help="nama kolom harga satuan (opsional)")
    a = ap.parse_args()

    p = Path(a.input)
    files = sorted(p.glob("*.xlsx")) if p.is_dir() else [p]
    files = [f for f in files if not f.stem.endswith("_HASIL") and not f.name.startswith("~$")]
    Path(a.out).mkdir(parents=True, exist_ok=True)
    for f in files:
        try:
            proses(f, a.db, a.out, a.qty, a.harga)
        except ValueError as e:
            sys.exit(str(e))
