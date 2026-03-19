# Necesar Aprovizionare - Calculator Excel VBA

Sistem de calcul necesar aprovizionare pentru magazin, cu import validat de date.

## Structura Proiectului

```
NecesarAprovizionare.xlsx    ← Fișierul Excel principal
sabloane/
  Sablon_StocDepozit.xlsx    ← Șablon pentru fișierul de stoc
  Sablon_VanzariMagazin.xlsx ← Șablon pentru fișierul de vânzări
vba_modules/
  ModImport.bas              ← Modulul principal (import + validare)
  ModButoane.bas              ← Creare butoane interactive
  ThisWorkbook.cls            ← Evenimente la deschidere
generate_excel.py             ← Script regenerare Excel
generate_sabloane.py          ← Script regenerare șabloane
```

## Configurare Inițială (o singură dată)

1. Deschideți `NecesarAprovizionare.xlsx` în Excel
2. **Salvați ca** → `NecesarAprovizionare.xlsm` (Excel Macro-Enabled Workbook)
3. Apăsați `Alt+F11` pentru a deschide editorul VBA
4. Click dreapta pe proiect → **Import File** → selectați pe rând:
   - `vba_modules/ModImport.bas`
   - `vba_modules/ModButoane.bas`
5. Faceți dublu-click pe **ThisWorkbook** în panel și copiați conținutul din `vba_modules/ThisWorkbook.cls`
6. Apăsați `Alt+F8` → selectați `CreazaButoane` → **Run** (o singură dată, creează butoanele)
7. Salvați fișierul

## Cum se Folosește

### Import Stoc Depozit
- Click pe butonul **IMPORT STOC DEPOZIT** din pagina Meniu
- Selectați fișierul `.xlsx` cu stocul din depozit
- Fișierul **trebuie** să aibă exact headerele: `CodProdus` | `Stoc`

### Import Vânzări Magazin
- Click pe butonul **IMPORT VÂNZĂRI MAGAZIN** din pagina Meniu
- Selectați fișierul `.xlsx` cu vânzările
- Fișierul **trebuie** să aibă exact headerele: `CodProdus` | `Cantitate`

### Ștergere Date
- Click pe butonul **ȘTERGE TOATE DATELE** (cu dublă confirmare)

## Validare Fișiere Import

Sistemul verifică automat:
- Headerele din rândul 1 trebuie să fie exact ca în șablon
- Fișierul trebuie să conțină cel puțin 1 rând de date
- Se acceptă doar fișiere `.xlsx` sau `.xls`

## Sheet-uri

| Sheet | Scop |
|-------|------|
| Meniu | Pagina principală cu butoane și status |
| StocDepozit | Datele importate de stoc din depozit |
| VanzariMagazin | Datele importate de vânzări |
| Compatibilitati | Tabel de compatibilități între produse |
| ReguliRotunjire | Reguli de rotunjire per categorie/produs |
| Log | Jurnal cu toate operațiunile efectuate |
