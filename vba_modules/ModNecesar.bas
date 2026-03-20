Attribute VB_Name = "ModNecesar"
'==============================================================================
' MODUL: ModNecesar
' DESCRIERE: Modul complet pentru calculul necesarului de aprovizionare.
'            Toate operatiile lucreaza in memorie cu arrays/matrici.
'            Scriere in Excel bulk (nu rand cu rand).
'
' FORMULA NECESAR:
'   MedieZilnica = Vanzare / ZileVanzare
'   NecesarBrut  = (MedieZilnica * ZileNecesar) - StocMagazin
'   Necesar      = MAX(0, CEILING(NecesarBrut)) -> apoi rotunjire categorie
'
' SETARI (pe sheet-ul Meniu):
'   C33 = Zile vanzare (pe cate zile e vanzarea importata)
'   C34 = Zile necesar (pe cate zile calculam necesarul)
'
' SETUP: Dupa import, rulati macro-ul "ConfigureazaAplicatia" (Alt+F8)
'==============================================================================
Option Explicit

' -- Constante -------------------------------------------------------------------
Private Const HEADER_STOC_COL1 As String = "CodProdus"
Private Const HEADER_STOC_COL2 As String = "Stoc"
Private Const HEADER_VANZARI_COL1 As String = "CodProdus"
Private Const HEADER_VANZARI_COL2 As String = "Cantitate"

Private Const SHEET_STOC As String = "StocMagazin"
Private Const SHEET_VANZARI As String = "VanzariMagazin"
Private Const SHEET_MENIU As String = "Meniu"
Private Const SHEET_LOG As String = "Log"
Private Const SHEET_COMPAT As String = "Compatibilitati"
Private Const SHEET_ROTUNJIRE As String = "ReguliRotunjire"
Private Const SHEET_NECESAR As String = "Necesar"

' Randul de unde incepe lista de produse in sheet-ul ReguliRotunjire
Private Const ROTUNJIRE_PRODUSE_START As Long = 8

' Celulele de setari pe sheet-ul Meniu
Private Const CELL_ZILE_VANZARE As String = "C33"
Private Const CELL_ZILE_NECESAR As String = "C34"

'==============================================================================
' CONFIGURARE APLICATIE - RULATI O SINGURA DATA DUPA IMPORT
'==============================================================================
Public Sub ConfigureazaAplicatia()
    On Error GoTo ErrHandler

    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)
    wsMeniu.Activate

    Dim shp As Shape
    For Each shp In wsMeniu.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp

    CreateButton wsMeniu, "btnImportStoc", "ImportStocDepozit", _
                 "IMPORT STOC MAGAZIN", "B7", "B8"
    CreateButton wsMeniu, "btnImportVanzari", "ImportVanzariMagazin", _
                 "IMPORT VANZARI MAGAZIN", "B10", "B11"
    CreateButton wsMeniu, "btnGenNecesar", "GenerareNecesar", _
                 "GENERARE NECESAR", "B14", "B15"
    CreateButton wsMeniu, "btnSterge", "StergeToateDatele", _
                 "STERGE TOATE DATELE", "B18", "B19"

    wsMeniu.Range("B2").Select
    MsgBox "Configurare completa! Salvati fisierul.", vbInformation, "OK"
    Exit Sub
ErrHandler:
    MsgBox "Eroare la configurare: " & Err.Description, vbCritical
End Sub

Private Sub CreateButton(ws As Worksheet, btnName As String, macroName As String, _
                          caption As String, cellTopLeft As String, cellBottomRight As String)
    Dim btn As Shape
    Dim rTop As Range, rBot As Range
    Set rTop = ws.Range(cellTopLeft)
    Set rBot = ws.Range(cellBottomRight)
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
        rTop.Left, rTop.Top, _
        rBot.Left + rBot.Width - rTop.Left, _
        rBot.Top + rBot.Height - rTop.Top)
    With btn
        .Name = btnName
        .OnAction = macroName
        .TextFrame.Characters.Text = caption
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
    End With
End Sub

'==============================================================================
' IMPORT STOC MAGAZIN (array-based)
'==============================================================================
Public Sub ImportStocDepozit()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectati fisierul cu STOC MAGAZIN")
    If filePath = "" Then Exit Sub

    If Not ValidateFileStructure(filePath, HEADER_STOC_COL1, HEADER_STOC_COL2) Then
        MsgBox "Fisierul nu are structura corecta!" & vbNewLine & _
               "Headerele asteptate: CodProdus | Stoc", vbCritical, "Eroare"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ClearSheetData SHEET_STOC
    Dim importedRows As Long
    importedRows = ImportDataBulk(filePath, SHEET_STOC, 2)
    UpdateImportStatus SHEET_STOC, importedRows

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Stoc Mag", "Importate " & importedRows & " produse", "OK"
    ThisWorkbook.Sheets(SHEET_MENIU).Activate
    MsgBox "Stoc magazin importat: " & importedRows & " produse.", vbInformation, "Import OK"
    Exit Sub
ErrHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Eroare la import: " & Err.Description, vbCritical, "Eroare"
End Sub

'==============================================================================
' IMPORT VANZARI MAGAZIN (array-based)
'==============================================================================
Public Sub ImportVanzariMagazin()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectati fisierul cu VANZARI MAGAZIN")
    If filePath = "" Then Exit Sub

    If Not ValidateFileStructure(filePath, HEADER_VANZARI_COL1, HEADER_VANZARI_COL2) Then
        MsgBox "Fisierul nu are structura corecta!" & vbNewLine & _
               "Headerele asteptate: CodProdus | Cantitate", vbCritical, "Eroare"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ClearSheetData SHEET_VANZARI
    Dim importedRows As Long
    importedRows = ImportDataBulk(filePath, SHEET_VANZARI, 2)
    UpdateImportStatus SHEET_VANZARI, importedRows

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Vanzari", "Importate " & importedRows & " produse", "OK"
    ThisWorkbook.Sheets(SHEET_MENIU).Activate
    MsgBox "Vanzari magazin importate: " & importedRows & " produse.", vbInformation, "Import OK"
    Exit Sub
ErrHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Eroare la import: " & Err.Description, vbCritical, "Eroare"
End Sub

'==============================================================================
' STERGE TOATE DATELE
'==============================================================================
Public Sub StergeToateDatele()
    If MsgBox("Stergeti toate datele importate si necesarul generat?", _
              vbQuestion + vbYesNo, "Confirmare") = vbNo Then Exit Sub

    ClearSheetData SHEET_STOC
    ClearSheetData SHEET_VANZARI
    ClearSheetData SHEET_NECESAR

    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)
    wsMeniu.Range("C20").Value = "Neincarcat"
    wsMeniu.Range("C20").Font.Color = RGB(198, 40, 40)
    wsMeniu.Range("C21").Value = "Neincarcat"
    wsMeniu.Range("C21").Font.Color = RGB(198, 40, 40)
    wsMeniu.Range("C22").Value = "-"
    wsMeniu.Range("C23").Value = "Negenerat"
    wsMeniu.Range("C23").Font.Color = RGB(198, 40, 40)

    WriteLog "Stergere", "Toate datele au fost sterse.", "OK"
    wsMeniu.Activate
    MsgBox "Datele au fost sterse.", vbInformation, "OK"
End Sub

'==============================================================================
' GENERARE NECESAR APROVIZIONARE (totul in memorie, scriere bulk)
'
' Formula:
'   MedieZilnica = VanzareTotal / ZileVanzare
'   NecesarBrut  = (MedieZilnica * ZileNecesar) - StocMagazin
'   Necesar      = MAX(0, rotunjire_sus_la_intreg(NecesarBrut))
'   apoi se aplica rotunjirea pe categorii daca exista
'==============================================================================
Public Sub GenerareNecesar()
    On Error GoTo ErrHandler

    Dim wsVanzari As Worksheet
    Set wsVanzari = ThisWorkbook.Sheets(SHEET_VANZARI)
    If wsVanzari.Cells(2, 1).Value = "" Then
        MsgBox "Nu exista date de vanzari importate!", vbExclamation, "Atentie"
        Exit Sub
    End If

    ' -- Citim setarile de pe Meniu --
    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)

    Dim zileVanzare As Long
    Dim zileNecesar As Long

    If IsNumeric(wsMeniu.Range(CELL_ZILE_VANZARE).Value) Then
        zileVanzare = CLng(wsMeniu.Range(CELL_ZILE_VANZARE).Value)
    Else
        zileVanzare = 0
    End If

    If IsNumeric(wsMeniu.Range(CELL_ZILE_NECESAR).Value) Then
        zileNecesar = CLng(wsMeniu.Range(CELL_ZILE_NECESAR).Value)
    Else
        zileNecesar = 0
    End If

    If zileVanzare <= 0 Or zileNecesar <= 0 Then
        MsgBox "Setarile de zile nu sunt configurate corect!" & vbNewLine & vbNewLine & _
               "Pe sheet-ul Meniu verificati:" & vbNewLine & _
               "  - Zile vanzare (C33) = " & wsMeniu.Range(CELL_ZILE_VANZARE).Value & vbNewLine & _
               "  - Zile necesar (C34) = " & wsMeniu.Range(CELL_ZILE_NECESAR).Value & vbNewLine & vbNewLine & _
               "Ambele trebuie sa fie numere > 0.", _
               vbCritical, "Eroare Setari"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' -- 1. Incarcare date in memorie (arrays + dictionare) --

    ' Compatibilitati: CodOriginal (col A=1) -> CodCompatibil (col C=3)
    Dim dictCompat As Object
    Set dictCompat = LoadDictFromSheet(SHEET_COMPAT, 2, 1, 3)

    ' Reguli rotunjire: CodProdus (col A=1) -> Categorie (col C=3)
    Dim dictRotunjire As Object
    Set dictRotunjire = LoadDictFromSheet(SHEET_ROTUNJIRE, ROTUNJIRE_PRODUSE_START, 1, 3)

    ' Stoc magazin: citire bulk in matrice -> dictionar
    Dim dictStoc As Object
    Set dictStoc = CreateObject("Scripting.Dictionary")
    dictStoc.CompareMode = vbTextCompare

    Dim wsStoc As Worksheet
    Set wsStoc = ThisWorkbook.Sheets(SHEET_STOC)
    Dim lastRowS As Long
    lastRowS = wsStoc.Cells(wsStoc.Rows.Count, 1).End(xlUp).Row
    If lastRowS >= 2 Then
        Dim arrStoc As Variant
        arrStoc = wsStoc.Range("A2:B" & lastRowS).Value
        Dim s As Long
        For s = 1 To UBound(arrStoc, 1)
            Dim codS As String
            codS = Trim(CStr(arrStoc(s, 1)))
            If codS <> "" Then
                If Not dictStoc.Exists(codS) Then
                    dictStoc.Add codS, CLng(arrStoc(s, 2))
                End If
            End If
        Next s
        Erase arrStoc
    End If

    ' -- 2. Citire vanzari in memorie (array bulk) --
    Dim lastRowV As Long
    lastRowV = wsVanzari.Cells(wsVanzari.Rows.Count, 1).End(xlUp).Row

    Dim arrVanzari As Variant
    arrVanzari = wsVanzari.Range("A2:B" & lastRowV).Value

    ' Dictionare pentru agregare
    Dim dictVanzare As Object   ' CodFinal -> vanzare totala
    Set dictVanzare = CreateObject("Scripting.Dictionary")
    dictVanzare.CompareMode = vbTextCompare

    Dim dictOrigCodes As Object ' CodFinal -> "CodOrig1(cant), CodOrig2(cant)"
    Set dictOrigCodes = CreateObject("Scripting.Dictionary")
    dictOrigCodes.CompareMode = vbTextCompare

    Dim r As Long
    Dim codOriginal As String, codFinal As String
    Dim cantitate As Long

    For r = 1 To UBound(arrVanzari, 1)
        codOriginal = Trim(CStr(arrVanzari(r, 1)))
        If codOriginal = "" Then GoTo NextVanzare

        cantitate = 0
        If IsNumeric(arrVanzari(r, 2)) Then cantitate = CLng(arrVanzari(r, 2))
        If cantitate <= 0 Then GoTo NextVanzare

        ' Verificam compatibilitate
        If dictCompat.Exists(codOriginal) Then
            codFinal = CStr(dictCompat(codOriginal))
        Else
            codFinal = codOriginal
        End If

        ' Agregam cantitatile per cod final
        If dictVanzare.Exists(codFinal) Then
            dictVanzare(codFinal) = dictVanzare(codFinal) + cantitate
        Else
            dictVanzare.Add codFinal, cantitate
        End If

        ' Evidenta coduri originale (doar la inlocuire)
        If codFinal <> codOriginal Then
            If Not dictOrigCodes.Exists(codFinal) Then
                dictOrigCodes.Add codFinal, codOriginal & "(" & cantitate & ")"
            Else
                dictOrigCodes(codFinal) = dictOrigCodes(codFinal) & ", " & codOriginal & "(" & cantitate & ")"
            End If
        End If
NextVanzare:
    Next r
    Erase arrVanzari

    ' -- 3. Construim matricea de output in memorie --
    Dim nrProduse As Long
    nrProduse = dictVanzare.Count

    If nrProduse = 0 Then
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Nu s-au gasit produse cu vanzari > 0.", vbExclamation, "Atentie"
        Exit Sub
    End If

    ' Matrice output: 6 coloane (Cod, Vanzare, MedieZi, StocMag, Necesar, Observatii)
    Dim arrOutput() As Variant
    ReDim arrOutput(1 To nrProduse, 1 To 6)

    Dim keys As Variant
    keys = dictVanzare.keys

    Dim i As Long
    Dim vanzareTotal As Long, necesarFinal As Long
    Dim catRotunjire As Long, stocMag As Long
    Dim medieZilnica As Double, necesarBrut As Double
    Dim necesarRotunjit As Long
    Dim obs As String

    For i = 0 To nrProduse - 1
        codFinal = CStr(keys(i))
        vanzareTotal = CLng(dictVanzare(codFinal))

        ' Calculam media zilnica si necesarul brut
        medieZilnica = CDbl(vanzareTotal) / CDbl(zileVanzare)

        ' Stoc magazin
        stocMag = 0
        If dictStoc.Exists(codFinal) Then stocMag = CLng(dictStoc(codFinal))

        ' Necesar brut = (medie * zile_necesar) - stoc
        necesarBrut = (medieZilnica * CDbl(zileNecesar)) - CDbl(stocMag)

        ' Rotunjim in sus la numar intreg, minim 0
        If necesarBrut <= 0 Then
            necesarRotunjit = 0
        Else
            necesarRotunjit = -Int(-necesarBrut)  ' Ceiling in VBA
        End If

        ' Categorie rotunjire
        catRotunjire = 0
        If dictRotunjire.Exists(codFinal) Then
            If IsNumeric(dictRotunjire(codFinal)) Then
                catRotunjire = CLng(dictRotunjire(codFinal))
            End If
        End If

        ' Aplicam rotunjirea pe categorii (daca exista si daca necesarul > 0)
        If necesarRotunjit > 0 And catRotunjire >= 1 And catRotunjire <= 4 Then
            necesarFinal = ApplyRounding(necesarRotunjit, catRotunjire, stocMag)
        Else
            necesarFinal = necesarRotunjit
        End If

        ' Construim observatii
        obs = ""
        If dictOrigCodes.Exists(codFinal) Then
            obs = "Compatibil: " & CStr(dictOrigCodes(codFinal))
        End If
        If catRotunjire > 0 And necesarFinal <> necesarRotunjit And necesarRotunjit > 0 Then
            If obs <> "" Then obs = obs & " | "
            obs = obs & "Rot.cat." & catRotunjire & " (" & necesarRotunjit & "->" & necesarFinal & ")"
        ElseIf catRotunjire > 0 And necesarRotunjit > 0 Then
            If obs <> "" Then obs = obs & " | "
            obs = obs & "Rot.cat." & catRotunjire
        End If

        ' Scriem in matricea de output (NU in celule!)
        arrOutput(i + 1, 1) = codFinal
        arrOutput(i + 1, 2) = vanzareTotal
        arrOutput(i + 1, 3) = Round(medieZilnica, 2)
        arrOutput(i + 1, 4) = stocMag
        arrOutput(i + 1, 5) = necesarFinal
        arrOutput(i + 1, 6) = obs
    Next i

    ' -- 4. Scriere BULK in sheet (o singura operatie) --
    Dim wsNecesar As Worksheet
    Set wsNecesar = ThisWorkbook.Sheets(SHEET_NECESAR)
    ClearSheetData SHEET_NECESAR

    Dim lastOut As Long
    lastOut = nrProduse + 1

    ' Setam coloana A ca Text INAINTE de scriere
    wsNecesar.Range("A2:A" & lastOut).NumberFormat = "@"

    ' Scriere matrice intreaga dintr-o data
    wsNecesar.Range("A2").Resize(nrProduse, 6).Value = arrOutput
    Erase arrOutput

    ' Formatare bulk
    wsNecesar.Range("A2:A" & lastOut).HorizontalAlignment = xlCenter
    With wsNecesar.Range("B2:E" & lastOut)
        .HorizontalAlignment = xlCenter
    End With
    wsNecesar.Range("B2:B" & lastOut).NumberFormat = "#,##0"
    wsNecesar.Range("C2:C" & lastOut).NumberFormat = "#,##0.00"
    wsNecesar.Range("D2:E" & lastOut).NumberFormat = "#,##0"
    wsNecesar.Range("F2:F" & lastOut).HorizontalAlignment = xlLeft

    ' Borders pe toata zona de date inclusiv header
    ApplyBorders wsNecesar.Range("A1:F" & lastOut)

    ' Auto-filter
    If wsNecesar.AutoFilterMode Then wsNecesar.AutoFilterMode = False
    wsNecesar.Range("A1:F" & lastOut).AutoFilter

    ' Actualizam statusul
    wsMeniu.Range("C23").Value = "Generat (" & nrProduse & " produse, " & _
        zileVanzare & "z vanz / " & zileNecesar & "z nec)"
    wsMeniu.Range("C23").Font.Color = RGB(46, 125, 50)
    wsMeniu.Range("C22").Value = Format(Now, "dd.mm.yyyy hh:nn:ss")

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Generare Necesar", "Generat " & nrProduse & " produse (" & _
        zileVanzare & "z/" & zileNecesar & "z)", "OK"
    ThisWorkbook.Sheets(SHEET_MENIU).Activate
    MsgBox "Necesar generat: " & nrProduse & " produse." & vbNewLine & _
           "(" & zileVanzare & " zile vanzare / " & zileNecesar & " zile necesar)", _
           vbInformation, "Necesar OK"
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WriteLog "Generare Necesar", "Eroare: " & Err.Description, "EROARE"
    MsgBox "Eroare la generare: " & Err.Description, vbCritical, "Eroare"
End Sub

'==============================================================================
' INCARCARE DICTIONAR DIN SHEET (bulk array read)
'==============================================================================
Private Function LoadDictFromSheet(sheetName As String, startRow As Long, _
                                    keyCol As Long, valCol As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < startRow Then
        Set LoadDictFromSheet = dict
        Exit Function
    End If

    Dim minCol As Long, maxCol As Long
    minCol = keyCol: maxCol = valCol
    If keyCol > valCol Then minCol = valCol: maxCol = keyCol

    ' Citire bulk in matrice
    Dim arr As Variant
    arr = ws.Range(ws.Cells(startRow, minCol), ws.Cells(lastRow, maxCol)).Value

    Dim keyIdx As Long, valIdx As Long
    keyIdx = keyCol - minCol + 1
    valIdx = valCol - minCol + 1

    Dim r As Long
    Dim k As String, v As String
    For r = 1 To UBound(arr, 1)
        If Not IsEmpty(arr(r, keyIdx)) Then
            k = Trim(CStr(arr(r, keyIdx)))
            If k <> "" And Not IsEmpty(arr(r, valIdx)) Then
                v = Trim(CStr(arr(r, valIdx)))
                If v <> "" Then
                    If Not dict.Exists(k) Then dict.Add k, v
                End If
            End If
        End If
    Next r
    Erase arr

    Set LoadDictFromSheet = dict
End Function

'==============================================================================
' IMPORT DATE BULK (citeste sursa in matrice, scrie bulk)
'==============================================================================
Private Function ImportDataBulk(ByVal filePath As String, _
                                 ByVal destSheetName As String, _
                                 ByVal nrCols As Long) As Long
    Dim wbSource As Workbook

    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Dim wsSource As Worksheet
    Set wsSource = wbSource.Sheets(1)

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        wbSource.Close SaveChanges:=False
        ImportDataBulk = 0
        Exit Function
    End If

    Dim nrRows As Long
    nrRows = lastRow - 1

    Dim colLetter As String
    colLetter = Chr(64 + nrCols)
    Dim arrData As Variant
    arrData = wsSource.Range("A2:" & colLetter & lastRow).Value

    wbSource.Close SaveChanges:=False

    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Sheets(destSheetName)

    ' Setam coloana A ca Text INAINTE de scriere
    wsDest.Range("A2:A" & lastRow).NumberFormat = "@"

    ' Scriere bulk
    wsDest.Range("A2").Resize(nrRows, nrCols).Value = arrData
    Erase arrData

    ' Formatare bulk
    wsDest.Range("A2:A" & lastRow).HorizontalAlignment = xlCenter
    With wsDest.Range("B2:" & colLetter & lastRow)
        .HorizontalAlignment = xlCenter
        .NumberFormat = "#,##0"
    End With

    ' Borders
    ApplyBorders wsDest.Range("A1:" & colLetter & lastRow)

    ' Auto-filter
    If wsDest.AutoFilterMode Then wsDest.AutoFilterMode = False
    wsDest.Range("A1:" & colLetter & lastRow).AutoFilter

    ImportDataBulk = nrRows
End Function

'==============================================================================
' REGULI DE ROTUNJIRE
'==============================================================================
Private Function ApplyRounding(ByVal val As Long, ByVal category As Long, _
                                ByVal depotStock As Long) As Long
    Dim result As Long

    If val = 0 Then
        ApplyRounding = 0
        Exit Function
    End If

    Select Case category
        Case 1  ' Multiplu de 10, minim 10
            If val < 5 Then
                result = 10
            ElseIf (val Mod 10) < 5 Then
                result = (val \ 10) * 10
            ElseIf (val Mod 10) = 0 Then
                result = val
            Else
                result = (val \ 10) * 10 + 10
            End If
            If result < 10 Then result = 10

        Case 2  ' Numar par (multiplu de 2)
            If (val Mod 2) = 0 Then
                result = val
            Else
                result = val + 1
            End If

        Case 3  ' Multiplu de 5, valori 1-2 devin 0
            If (val Mod 5) < 3 Then
                result = (val \ 5) * 5
            ElseIf (val Mod 5) = 0 Then
                result = val
            Else
                result = (val \ 5) * 5 + 5
            End If

        Case 4  ' Multiplu de 5, minim 5 daca stoc magazin < 3
            If (val Mod 5) < 3 Then
                result = (val \ 5) * 5
            ElseIf (val Mod 5) = 0 Then
                result = val
            Else
                result = (val \ 5) * 5 + 5
            End If
            If result = 0 And depotStock < 3 Then result = 5

        Case Else
            result = val
    End Select

    ApplyRounding = result
End Function

'==============================================================================
' APLICA BORDERS PE UN RANGE
'==============================================================================
Private Sub ApplyBorders(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Sub

'==============================================================================
' FUNCTII AUXILIARE
'==============================================================================

Private Function SelectImportFile(ByVal dialogTitle As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Fisiere Excel", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = -1 Then
            SelectImportFile = .SelectedItems(1)
        Else
            SelectImportFile = ""
        End If
    End With
End Function

Private Function ValidateFileStructure(ByVal filePath As String, _
                                        ByVal expectedCol1 As String, _
                                        ByVal expectedCol2 As String) As Boolean
    ValidateFileStructure = False
    Dim wbSource As Workbook

    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Dim h1 As String, h2 As String
    h1 = CleanHeader(CStr(wbSource.Sheets(1).Cells(1, 1).Value))
    h2 = CleanHeader(CStr(wbSource.Sheets(1).Cells(1, 2).Value))

    If LCase(h1) = LCase(expectedCol1) And LCase(h2) = LCase(expectedCol2) Then
        If wbSource.Sheets(1).Cells(2, 1).Value <> "" Then
            ValidateFileStructure = True
        End If
    End If

    wbSource.Close SaveChanges:=False
End Function

Private Function CleanHeader(ByVal s As String) As String
    s = Trim(s)
    If Len(s) > 0 Then
        If AscW(Left(s, 1)) = 65279 Then s = Mid(s, 2)
    End If
    s = Replace(s, Chr(160), "")
    CleanHeader = Trim(s)
End Function

Private Sub ClearSheetData(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Rows("2:" & lastRow).Clear
    End If
End Sub

Private Sub UpdateImportStatus(ByVal importType As String, ByVal rowCount As Long)
    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)

    Dim statusCell As Range
    If importType = SHEET_STOC Then
        Set statusCell = wsMeniu.Range("C20")
    ElseIf importType = SHEET_VANZARI Then
        Set statusCell = wsMeniu.Range("C21")
    Else
        Exit Sub
    End If

    statusCell.Value = "Incarcat (" & rowCount & " produse)"
    statusCell.Font.Color = RGB(46, 125, 50)
    wsMeniu.Range("C22").Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    wsMeniu.Range("C22").Font.Color = RGB(51, 51, 51)
End Sub

Private Sub WriteLog(ByVal operatiune As String, ByVal detalii As String, ByVal status As String)
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOG)
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    Dim arrLog(1 To 1, 1 To 4) As Variant
    arrLog(1, 1) = Format(Now, "dd.mm.yyyy hh:nn:ss")
    arrLog(1, 2) = operatiune
    arrLog(1, 3) = detalii
    arrLog(1, 4) = status
    wsLog.Range(wsLog.Cells(nextRow, 1), wsLog.Cells(nextRow, 4)).Value = arrLog

    If status = "OK" Then
        wsLog.Cells(nextRow, 4).Font.Color = RGB(46, 125, 50)
    ElseIf status = "EROARE" Then
        wsLog.Cells(nextRow, 4).Font.Color = RGB(198, 40, 40)
    End If
    wsLog.Cells(nextRow, 1).HorizontalAlignment = xlCenter
    wsLog.Cells(nextRow, 4).HorizontalAlignment = xlCenter

    ApplyBorders wsLog.Range(wsLog.Cells(nextRow, 1), wsLog.Cells(nextRow, 4))
    On Error GoTo 0
End Sub
