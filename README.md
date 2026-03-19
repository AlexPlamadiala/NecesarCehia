# Necesar Aprovizionare - Calculator Excel VBA

Sistem de calcul necesar aprovizionare pentru magazin, cu import validat de date.

## Setup Rapid (3 pasi)

1. Deschideti `NecesarAprovizionare.xlsx` in Excel
2. **Save As** -> selectati tipul **Excel Macro-Enabled Workbook (.xlsm)**
3. `Alt+F11` -> click dreapta pe proiect -> **Import File** -> selectati `vba_modules/ModNecesar.bas`
4. `Alt+F8` -> selectati **ConfigureazaAplicatia** -> **Run**
5. Salvati fisierul. Gata!

## Cum se Foloseste

### Import Stoc Depozit
- Click pe butonul **IMPORT STOC DEPOZIT** din pagina Meniu
- Selectati fisierul `.xlsx` cu stocul din depozit
- Fisierul **trebuie** sa aiba exact headerele: `CodProdus` | `Stoc`

### Import Vanzari Magazin
- Click pe butonul **IMPORT VANZARI MAGAZIN** din pagina Meniu
- Selectati fisierul `.xlsx` cu vanzarile
- Fisierul **trebuie** sa aiba exact headerele: `CodProdus` | `Cantitate`

### Stergere Date
- Click pe butonul **STERGE TOATE DATELE** (cu dubla confirmare)

## Validare Fisiere Import

Sistemul verifica automat:
- Headerele din randul 1 trebuie sa fie exact ca in sablon
- Fisierul trebuie sa contina cel putin 1 rand de date
- Se accepta doar fisiere `.xlsx` sau `.xls`

## Structura Proiectului

```
NecesarAprovizionare.xlsx       <- Fisierul Excel principal
sabloane/
  Sablon_StocDepozit.xlsx       <- Sablon import stoc (CodProdus | Stoc)
  Sablon_VanzariMagazin.xlsx    <- Sablon import vanzari (CodProdus | Cantitate)
vba_modules/
  ModNecesar.bas                <- Modulul VBA complet (un singur fisier)
generate_excel.py               <- Script regenerare Excel
generate_sabloane.py            <- Script regenerare sabloane
```

## Sheet-uri

| Sheet | Scop |
|-------|------|
| Meniu | Pagina principala cu butoane si status |
| StocDepozit | Datele importate de stoc din depozit |
| VanzariMagazin | Datele importate de vanzari |
| Compatibilitati | Tabel de compatibilitati intre produse |
| ReguliRotunjire | Reguli de rotunjire per categorie/produs |
| Log | Jurnal cu toate operatiunile efectuate |
