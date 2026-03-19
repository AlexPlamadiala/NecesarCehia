#!/usr/bin/env python3
"""
Generator pentru fisierul Excel 'NecesarAprovizionare.xlsx'
Creeaza structura de baza cu sheet-uri si formatare.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# -- Constante de stil -----------------------------------------------------------
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="2F5496")
SUBTITLE_FONT = Font(name="Calibri", bold=False, size=10, color="666666")

BTN_FONT = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
BTN_FILL_GREEN = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
BTN_FILL_BLUE = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
BTN_FILL_RED = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
BTN_ALIGN = Alignment(horizontal="center", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

DATA_ALIGN = Alignment(horizontal="center", vertical="center")


def style_header_row(ws, row, col_start, col_end):
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def create_menu_sheet(wb):
    ws = wb.active
    ws.title = "Meniu"
    ws.sheet_properties.tabColor = "2F5496"

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 3

    # Titlu
    ws.merge_cells("B2:C2")
    title_cell = ws["B2"]
    title_cell.value = "NECESAR APROVIZIONARE - MAGAZIN"
    title_cell.font = Font(name="Calibri", bold=True, size=18, color="2F5496")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B3:C3")
    sub_cell = ws["B3"]
    sub_cell.value = "Sistem de calcul necesar aprovizionare"
    sub_cell.font = SUBTITLE_FONT
    sub_cell.alignment = Alignment(horizontal="center")

    # -- Sectiunea 1: Import Date --
    ws.merge_cells("B5:C5")
    ws["B5"].value = "PASUL 1: IMPORT DATE"
    ws["B5"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    # Zona buton Import Stoc Depozit
    ws.merge_cells("B7:B8")
    btn1 = ws["B7"]
    btn1.value = "IMPORT STOC DEPOZIT"
    btn1.font = BTN_FONT
    btn1.fill = BTN_FILL_GREEN
    btn1.alignment = BTN_ALIGN
    btn1.border = THIN_BORDER

    ws.merge_cells("C7:C8")
    desc1 = ws["C7"]
    desc1.value = "Importa fisierul cu stocul din depozit.\nFormat: CodProdus | Stoc"
    desc1.font = Font(name="Calibri", size=10, color="333333")
    desc1.alignment = Alignment(vertical="center", wrap_text=True)

    # Zona buton Import Vanzari Magazin
    ws.merge_cells("B10:B11")
    btn2 = ws["B10"]
    btn2.value = "IMPORT VANZARI MAGAZIN"
    btn2.font = BTN_FONT
    btn2.fill = BTN_FILL_BLUE
    btn2.alignment = BTN_ALIGN
    btn2.border = THIN_BORDER

    ws.merge_cells("C10:C11")
    desc2 = ws["C10"]
    desc2.value = "Importa fisierul cu vanzarile magazinului.\nFormat: CodProdus | Cantitate"
    desc2.font = Font(name="Calibri", size=10, color="333333")
    desc2.alignment = Alignment(vertical="center", wrap_text=True)

    # -- Sectiunea 2: Stergere Date --
    ws.merge_cells("B13:C13")
    ws["B13"].value = "ACTIUNI RAPIDE"
    ws["B13"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    ws.merge_cells("B15:B16")
    btn3 = ws["B15"]
    btn3.value = "STERGE TOATE DATELE"
    btn3.font = BTN_FONT
    btn3.fill = BTN_FILL_RED
    btn3.alignment = BTN_ALIGN
    btn3.border = THIN_BORDER

    ws.merge_cells("C15:C16")
    desc3 = ws["C15"]
    desc3.value = "Sterge toate datele importate.\nFoloseste pentru a reincepe procesul."
    desc3.font = Font(name="Calibri", size=10, color="333333")
    desc3.alignment = Alignment(vertical="center", wrap_text=True)

    # -- Status --
    ws.merge_cells("B18:C18")
    ws["B18"].value = "STATUS IMPORT"
    ws["B18"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    ws["B20"].value = "Stoc Depozit:"
    ws["B20"].font = Font(name="Calibri", bold=True, size=11)
    ws["C20"].value = "Neincarcat"
    ws["C20"].font = Font(name="Calibri", size=11, color="C62828")

    ws["B21"].value = "Vanzari Magazin:"
    ws["B21"].font = Font(name="Calibri", bold=True, size=11)
    ws["C21"].value = "Neincarcat"
    ws["C21"].font = Font(name="Calibri", size=11, color="C62828")

    ws["B22"].value = "Ultima actualizare:"
    ws["B22"].font = Font(name="Calibri", bold=True, size=11)
    ws["C22"].value = "-"
    ws["C22"].font = Font(name="Calibri", size=11, color="666666")

    # -- Instructiuni --
    ws.merge_cells("B24:C24")
    ws["B24"].value = "INSTRUCTIUNI"
    ws["B24"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    instructions = [
        "1. Pregatiti fisierele de import conform sabloanelor din folderul 'sabloane'.",
        "2. Fisierul de stoc trebuie sa aiba headerul: CodProdus | Stoc",
        "3. Fisierul de vanzari trebuie sa aiba headerul: CodProdus | Cantitate",
        "4. Apasati butonul corespunzator pentru a importa datele.",
        "5. Selectati fisierul .xlsx si asteptati confirmarea importului.",
    ]
    for i, text in enumerate(instructions):
        row = 26 + i
        ws.merge_cells(f"B{row}:C{row}")
        ws[f"B{row}"].value = text
        ws[f"B{row}"].font = Font(name="Calibri", size=10, color="444444")

    return ws


def create_stoc_depozit_sheet(wb):
    ws = wb.create_sheet("StocDepozit")
    ws.sheet_properties.tabColor = "2E7D32"

    headers = ["Cod Produs", "Stoc Depozit"]
    col_widths = [25, 20]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[chr(64 + i)].width = width

    style_header_row(ws, 1, 1, len(headers))
    ws.auto_filter.ref = "A1:B1"
    return ws


def create_vanzari_magazin_sheet(wb):
    ws = wb.create_sheet("VanzariMagazin")
    ws.sheet_properties.tabColor = "1565C0"

    headers = ["Cod Produs", "Cantitate Vanzare"]
    col_widths = [25, 22]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[chr(64 + i)].width = width

    style_header_row(ws, 1, 1, len(headers))
    ws.auto_filter.ref = "A1:B1"
    return ws


def create_compatibilitati_sheet(wb):
    ws = wb.create_sheet("Compatibilitati")
    ws.sheet_properties.tabColor = "F57C00"

    headers = ["Cod Produs", "Cod Produs Compatibil", "Observatii"]
    col_widths = [25, 25, 40]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[chr(64 + i)].width = width

    style_header_row(ws, 1, 1, len(headers))
    ws.auto_filter.ref = "A1:C1"
    return ws


def create_reguli_rotunjire_sheet(wb):
    ws = wb.create_sheet("ReguliRotunjire")
    ws.sheet_properties.tabColor = "7B1FA2"

    headers = ["Categorie / Cod Produs", "Unitate Masura", "Rotunjire La", "Tip Rotunjire"]
    col_widths = [30, 20, 18, 22]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[chr(64 + i)].width = width

    style_header_row(ws, 1, 1, len(headers))

    # Exemple
    examples = [
        ["*", "buc", 1, "Sus (Ceiling)"],
        ["EXEMPLU_COD_1", "kg", 0.5, "Sus (Ceiling)"],
        ["EXEMPLU_COD_2", "buc", 6, "Sus (Ceiling)"],
    ]
    for r, row_data in enumerate(examples, 2):
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = DATA_ALIGN
            cell.border = THIN_BORDER
            if r == 2:
                cell.font = Font(name="Calibri", italic=True, color="999999")

    ws.auto_filter.ref = "A1:D1"
    return ws


def create_log_sheet(wb):
    ws = wb.create_sheet("Log")
    ws.sheet_properties.tabColor = "757575"

    headers = ["Data/Ora", "Operatiune", "Detalii", "Status"]
    col_widths = [22, 25, 50, 15]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[chr(64 + i)].width = width

    style_header_row(ws, 1, 1, len(headers))
    return ws


def main():
    wb = openpyxl.Workbook()

    create_menu_sheet(wb)
    create_stoc_depozit_sheet(wb)
    create_vanzari_magazin_sheet(wb)
    create_compatibilitati_sheet(wb)
    create_reguli_rotunjire_sheet(wb)
    create_log_sheet(wb)

    output_path = "NecesarAprovizionare.xlsx"
    wb.save(output_path)
    print(f"Fisierul '{output_path}' a fost generat cu succes!")
    print()
    print("SETUP RAPID (3 pasi):")
    print("  1. Deschideti fisierul in Excel -> Save As -> .xlsm")
    print("  2. Alt+F11 -> Import File -> vba_modules/ModNecesar.bas")
    print("  3. Alt+F8 -> ConfigureazaAplicatia -> Run")
    print("  Gata! Salvati si folositi.")


if __name__ == "__main__":
    main()
