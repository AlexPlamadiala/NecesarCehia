#!/usr/bin/env python3
"""
Genereaza fisierele sablon pentru import cu ~2000 produse.
Toate cantitatile sunt numere intregi (naturale).
"""
import random
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
DATA_ALIGN = Alignment(horizontal="center")

NUM_PRODUCTS = 2000
random.seed(42)  # Rezultate reproductibile


def generate_product_codes(n):
    """Genereaza n coduri de produs realiste (format EAN-13 like)."""
    codes = set()
    while len(codes) < n:
        # Format: 2700000XXXXXX (13 cifre, similar cu cele din screenshot)
        suffix = random.randint(10000, 999999)
        code = f"2700000{suffix:06d}"
        codes.add(code)
    return sorted(codes)


def style_header(ws, row, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def create_sablon_stoc():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "StocMagazin"

    ws["A1"] = "CodProdus"
    ws["B1"] = "Stoc"
    style_header(ws, 1, 2)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 12

    codes = generate_product_codes(NUM_PRODUCTS)
    for i, code in enumerate(codes, 2):
        stoc = random.randint(0, 500)
        ws.cell(row=i, column=1, value=code).alignment = DATA_ALIGN
        ws.cell(row=i, column=2, value=stoc).alignment = DATA_ALIGN

    wb.save("sabloane/Sablon_StocMagazin.xlsx")
    print(f"  - sabloane/Sablon_StocMagazin.xlsx ({NUM_PRODUCTS} produse)")


def create_sablon_vanzari():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VanzariMagazin"

    ws["A1"] = "CodProdus"
    ws["B1"] = "Cantitate"
    style_header(ws, 1, 2)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 15

    # Folosim aceleasi coduri ca la stoc (seed identic)
    codes = generate_product_codes(NUM_PRODUCTS)
    for i, code in enumerate(codes, 2):
        cantitate = random.randint(0, 100)
        ws.cell(row=i, column=1, value=code).alignment = DATA_ALIGN
        ws.cell(row=i, column=2, value=cantitate).alignment = DATA_ALIGN

    wb.save("sabloane/Sablon_VanzariMagazin.xlsx")
    print(f"  - sabloane/Sablon_VanzariMagazin.xlsx ({NUM_PRODUCTS} produse)")


if __name__ == "__main__":
    print("Generare sabloane import...")
    create_sablon_stoc()
    create_sablon_vanzari()
    print(f"\nSabloanele cu {NUM_PRODUCTS} produse au fost generate in 'sabloane/'")
    print("Cantitatile sunt numere intregi (naturale).")
