#!/usr/bin/env python3
"""
Generează fișierele șablon pentru import.
Acestea trebuie respectate exact la structură pentru a trece validarea.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
EXAMPLE_FONT = Font(name="Calibri", italic=True, color="999999")


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
    ws.title = "StocDepozit"

    # HEADERELE OBLIGATORII - exact așa trebuie să fie
    ws["A1"] = "CodProdus"
    ws["B1"] = "Stoc"
    style_header(ws, 1, 2)

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 15

    # Exemple
    examples = [
        ("PROD001", 150),
        ("PROD002", 75.5),
        ("PROD003", 200),
    ]
    for i, (cod, stoc) in enumerate(examples, 2):
        ws.cell(row=i, column=1, value=cod).font = EXAMPLE_FONT
        ws.cell(row=i, column=2, value=stoc).font = EXAMPLE_FONT
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=i, column=2).alignment = Alignment(horizontal="center")

    wb.save("sabloane/Sablon_StocDepozit.xlsx")
    print("  ✓ sabloane/Sablon_StocDepozit.xlsx")


def create_sablon_vanzari():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VanzariMagazin"

    # HEADERELE OBLIGATORII
    ws["A1"] = "CodProdus"
    ws["B1"] = "Cantitate"
    style_header(ws, 1, 2)

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 20

    # Exemple
    examples = [
        ("PROD001", 30),
        ("PROD002", 12.5),
        ("PROD003", 45),
    ]
    for i, (cod, cant) in enumerate(examples, 2):
        ws.cell(row=i, column=1, value=cod).font = EXAMPLE_FONT
        ws.cell(row=i, column=2, value=cant).font = EXAMPLE_FONT
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=i, column=2).alignment = Alignment(horizontal="center")

    wb.save("sabloane/Sablon_VanzariMagazin.xlsx")
    print("  ✓ sabloane/Sablon_VanzariMagazin.xlsx")


if __name__ == "__main__":
    print("Generare șabloane import...")
    create_sablon_stoc()
    create_sablon_vanzari()
    print("\nȘabloanele au fost generate în folderul 'sabloane/'")
    print("Folosiți aceste fișiere ca model pentru datele de import.")
