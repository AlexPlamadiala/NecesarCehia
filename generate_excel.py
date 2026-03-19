#!/usr/bin/env python3
"""
Generator pentru fișierul Excel 'NecesarAprovizionare.xlsm'
Creează structura de bază cu sheet-uri, formatare și protecție.
Codul VBA trebuie importat manual din fișierul vba_modules/.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

# ── Constante de stil ──────────────────────────────────────────────
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
    """Aplică stilul de header pe un rând."""
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER


def create_menu_sheet(wb):
    """Sheet-ul principal - Meniu."""
    ws = wb.active
    ws.title = "Meniu"
    ws.sheet_properties.tabColor = "2F5496"

    # Setări vizuale
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

    # Subtitlu
    ws.merge_cells("B3:C3")
    sub_cell = ws["B3"]
    sub_cell.value = "Sistem de calcul necesar aprovizionare"
    sub_cell.font = SUBTITLE_FONT
    sub_cell.alignment = Alignment(horizontal="center")

    # ── Secțiunea 1: Import Date ──
    ws.merge_cells("B5:C5")
    ws["B5"].value = "PASUL 1: IMPORT DATE"
    ws["B5"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    # Buton Import Stoc Depozit
    ws.merge_cells("B7:B8")
    btn1 = ws["B7"]
    btn1.value = "▶  IMPORT STOC DEPOZIT"
    btn1.font = BTN_FONT
    btn1.fill = BTN_FILL_GREEN
    btn1.alignment = BTN_ALIGN
    btn1.border = THIN_BORDER

    ws.merge_cells("C7:C8")
    desc1 = ws["C7"]
    desc1.value = "Importă fișierul cu stocul din depozit.\nFormat: Cod Produs | Stoc"
    desc1.font = Font(name="Calibri", size=10, color="333333")
    desc1.alignment = Alignment(vertical="center", wrap_text=True)

    # Buton Import Vânzări Magazin
    ws.merge_cells("B10:B11")
    btn2 = ws["B10"]
    btn2.value = "▶  IMPORT VÂNZĂRI MAGAZIN"
    btn2.font = BTN_FONT
    btn2.fill = BTN_FILL_BLUE
    btn2.alignment = BTN_ALIGN
    btn2.border = THIN_BORDER

    ws.merge_cells("C10:C11")
    desc2 = ws["C10"]
    desc2.value = "Importă fișierul cu vânzările magazinului.\nFormat: Cod Produs | Cantitate Vânzare"
    desc2.font = Font(name="Calibri", size=10, color="333333")
    desc2.alignment = Alignment(vertical="center", wrap_text=True)

    # ── Secțiunea 2: Ștergere Date ──
    ws.merge_cells("B13:C13")
    ws["B13"].value = "ACȚIUNI RAPIDE"
    ws["B13"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    ws.merge_cells("B15:B16")
    btn3 = ws["B15"]
    btn3.value = "✕  ȘTERGE TOATE DATELE"
    btn3.font = BTN_FONT
    btn3.fill = BTN_FILL_RED
    btn3.alignment = BTN_ALIGN
    btn3.border = THIN_BORDER

    ws.merge_cells("C15:C16")
    desc3 = ws["C15"]
    desc3.value = "Șterge toate datele importate.\nFolosește pentru a reîncepe procesul."
    desc3.font = Font(name="Calibri", size=10, color="333333")
    desc3.alignment = Alignment(vertical="center", wrap_text=True)

    # ── Status ──
    ws.merge_cells("B18:C18")
    ws["B18"].value = "STATUS IMPORT"
    ws["B18"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    ws["B20"].value = "Stoc Depozit:"
    ws["B20"].font = Font(name="Calibri", bold=True, size=11)
    ws["C20"].value = "⊘ Neîncărcat"
    ws["C20"].font = Font(name="Calibri", size=11, color="C62828")

    ws["B21"].value = "Vânzări Magazin:"
    ws["B21"].font = Font(name="Calibri", bold=True, size=11)
    ws["C21"].value = "⊘ Neîncărcat"
    ws["C21"].font = Font(name="Calibri", size=11, color="C62828")

    ws["B22"].value = "Ultima actualizare:"
    ws["B22"].font = Font(name="Calibri", bold=True, size=11)
    ws["C22"].value = "-"
    ws["C22"].font = Font(name="Calibri", size=11, color="666666")

    # ── Instrucțiuni ──
    ws.merge_cells("B24:C24")
    ws["B24"].value = "INSTRUCȚIUNI"
    ws["B24"].font = Font(name="Calibri", bold=True, size=13, color="2F5496")

    instructions = [
        "1. Pregătiți fișierele de import conform șabloanelor din folderul 'Sabloane'.",
        "2. Fișierul de stoc trebuie să aibă headerul: CodProdus | Stoc",
        "3. Fișierul de vânzări trebuie să aibă headerul: CodProdus | Cantitate",
        '4. Apăsați butonul corespunzător sau folosiți meniul "Necesar" din bara de sus.',
        "5. Selectați fișierul .xlsx și așteptați confirmarea importului.",
    ]
    for i, text in enumerate(instructions):
        row = 26 + i
        ws.merge_cells(f"B{row}:C{row}")
        ws[f"B{row}"].value = text
        ws[f"B{row}"].font = Font(name="Calibri", size=10, color="444444")

    return ws


def create_stoc_depozit_sheet(wb):
    """Sheet pentru datele de stoc din depozit."""
    ws = wb.create_sheet("StocDepozit")
    ws.sheet_properties.tabColor = "2E7D32"

    headers = ["Cod Produs", "Stoc Depozit"]
    col_widths = [25, 20]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[get_column_letter(i)].width = width

    style_header_row(ws, 1, 1, len(headers))

    # Adaugă auto-filter
    ws.auto_filter.ref = "A1:B1"

    return ws


def create_vanzari_magazin_sheet(wb):
    """Sheet pentru datele de vânzări din magazin."""
    ws = wb.create_sheet("VanzariMagazin")
    ws.sheet_properties.tabColor = "1565C0"

    headers = ["Cod Produs", "Cantitate Vânzare"]
    col_widths = [25, 22]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[get_column_letter(i)].width = width

    style_header_row(ws, 1, 1, len(headers))

    ws.auto_filter.ref = "A1:B1"

    return ws


def create_compatibilitati_sheet(wb):
    """Sheet pentru compatibilități per produs."""
    ws = wb.create_sheet("Compatibilitati")
    ws.sheet_properties.tabColor = "F57C00"

    headers = ["Cod Produs", "Cod Produs Compatibil", "Observații"]
    col_widths = [25, 25, 40]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[get_column_letter(i)].width = width

    style_header_row(ws, 1, 1, len(headers))

    ws.auto_filter.ref = "A1:C1"

    return ws


def create_reguli_rotunjire_sheet(wb):
    """Sheet pentru regulile de rotunjire."""
    ws = wb.create_sheet("ReguliRotunjire")
    ws.sheet_properties.tabColor = "7B1FA2"

    headers = ["Categorie / Cod Produs", "Unitate Măsură", "Rotunjire La", "Tip Rotunjire"]
    col_widths = [30, 20, 18, 22]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[get_column_letter(i)].width = width

    style_header_row(ws, 1, 1, len(headers))

    # Adaugă exemple
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
    """Sheet pentru logul operațiunilor."""
    ws = wb.create_sheet("Log")
    ws.sheet_properties.tabColor = "757575"

    headers = ["Data/Ora", "Operațiune", "Detalii", "Status"]
    col_widths = [22, 25, 50, 15]

    for i, (header, width) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=1, column=i, value=header)
        ws.column_dimensions[get_column_letter(i)].width = width

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

    # Salvăm ca .xlsx (VBA se adaugă manual - vezi README)
    output_path = "NecesarAprovizionare.xlsx"
    wb.save(output_path)
    print(f"✓ Fișierul '{output_path}' a fost generat cu succes!")
    print()
    print("Sheet-uri create:")
    print("  • Meniu          - Pagina principală cu butoane")
    print("  • StocDepozit    - Date stoc depozit (import)")
    print("  • VanzariMagazin - Date vânzări magazin (import)")
    print("  • Compatibilitati - Compatibilități per produs")
    print("  • ReguliRotunjire - Reguli de rotunjire")
    print("  • Log            - Jurnal operațiuni")
    print()
    print("IMPORTANT: Pentru a activa butoanele VBA:")
    print("  1. Deschideți fișierul în Excel")
    print("  2. Salvați ca .xlsm (Excel Macro-Enabled Workbook)")
    print("  3. Alt+F11 → Import File → selectați fișierele din vba_modules/")
    print("  4. Salvați din nou")


if __name__ == "__main__":
    main()
