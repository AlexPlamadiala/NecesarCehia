Attribute VB_Name = "ModNecesar"
'==============================================================================
' MODUL: ModNecesar
' DESCRIERE: Modul complet pentru calculul necesarului de aprovizionare.
'            Import stoc depozit si vanzari magazin cu validare stricta.
'
' SETUP: Dupa import, rulati macro-ul "ConfigureazaAplicatia" (Alt+F8)
'        Acesta creeaza butoanele si configureaza totul automat.
'==============================================================================
Option Explicit

' -- Constante pentru validare structura -----------------------------------------------
Private Const HEADER_STOC_COL1 As String = "CodProdus"
Private Const HEADER_STOC_COL2 As String = "Stoc"
Private Const HEADER_VANZARI_COL1 As String = "CodProdus"
Private Const HEADER_VANZARI_COL2 As String = "Cantitate"

Private Const SHEET_STOC As String = "StocDepozit"
Private Const SHEET_VANZARI As String = "VanzariMagazin"
Private Const SHEET_MENIU As String = "Meniu"
Private Const SHEET_LOG As String = "Log"

'==============================================================================
' CONFIGURARE APLICATIE - RULATI O SINGURA DATA DUPA IMPORT
'==============================================================================
Public Sub ConfigureazaAplicatia()
    On Error GoTo ErrHandler

    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)
    wsMeniu.Activate

    ' Stergem butoane existente
    Dim shp As Shape
    For Each shp In wsMeniu.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp

    ' -- Buton 1: Import Stoc Depozit --
    Dim btn1 As Shape
    Set btn1 = wsMeniu.Shapes.AddFormControl( _
        xlButtonControl, _
        wsMeniu.Range("B7").Left, _
        wsMeniu.Range("B7").Top, _
        wsMeniu.Range("B8").Left + wsMeniu.Range("B8").Width - wsMeniu.Range("B7").Left, _
        wsMeniu.Range("B8").Top + wsMeniu.Range("B8").Height - wsMeniu.Range("B7").Top)
    With btn1
        .Name = "btnImportStoc"
        .OnAction = "ImportStocDepozit"
        .TextFrame.Characters.Text = "IMPORT STOC DEPOZIT"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
    End With

    ' -- Buton 2: Import Vanzari Magazin --
    Dim btn2 As Shape
    Set btn2 = wsMeniu.Shapes.AddFormControl( _
        xlButtonControl, _
        wsMeniu.Range("B10").Left, _
        wsMeniu.Range("B10").Top, _
        wsMeniu.Range("B11").Left + wsMeniu.Range("B11").Width - wsMeniu.Range("B10").Left, _
        wsMeniu.Range("B11").Top + wsMeniu.Range("B11").Height - wsMeniu.Range("B10").Top)
    With btn2
        .Name = "btnImportVanzari"
        .OnAction = "ImportVanzariMagazin"
        .TextFrame.Characters.Text = "IMPORT VANZARI MAGAZIN"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
    End With

    ' -- Buton 3: Sterge Toate Datele --
    Dim btn3 As Shape
    Set btn3 = wsMeniu.Shapes.AddFormControl( _
        xlButtonControl, _
        wsMeniu.Range("B15").Left, _
        wsMeniu.Range("B15").Top, _
        wsMeniu.Range("B16").Left + wsMeniu.Range("B16").Width - wsMeniu.Range("B15").Left, _
        wsMeniu.Range("B16").Top + wsMeniu.Range("B16").Height - wsMeniu.Range("B15").Top)
    With btn3
        .Name = "btnSterge"
        .OnAction = "StergeToateDatele"
        .TextFrame.Characters.Text = "STERGE TOATE DATELE"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Bold = True
    End With

    ' Configureaza evenimentul Workbook_Open prin auto-navigare
    wsMeniu.Range("B2").Select

    MsgBox "Configurare completa!" & vbNewLine & vbNewLine & _
           "Butoane create:" & vbNewLine & _
           "  - IMPORT STOC DEPOZIT" & vbNewLine & _
           "  - IMPORT VANZARI MAGAZIN" & vbNewLine & _
           "  - STERGE TOATE DATELE" & vbNewLine & vbNewLine & _
           "Salvati fisierul si puteti incepe lucrul.", _
           vbInformation, "Configurare Aplicatie"

    Exit Sub

ErrHandler:
    MsgBox "Eroare la configurare: " & Err.Description, vbCritical
End Sub

'==============================================================================
' IMPORT STOC DEPOZIT
'==============================================================================
Public Sub ImportStocDepozit()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectati fisierul cu STOC DEPOZIT")
    If filePath = "" Then Exit Sub

    ' Validare structura
    If Not ValidateFileStructure(filePath, HEADER_STOC_COL1, HEADER_STOC_COL2) Then
        MsgBox "Fisierul selectat NU are structura corecta!" & vbNewLine & vbNewLine & _
               "Structura asteptata (primele 2 coloane, randul 1):" & vbNewLine & _
               "  Coloana A: " & HEADER_STOC_COL1 & vbNewLine & _
               "  Coloana B: " & HEADER_STOC_COL2 & vbNewLine & vbNewLine & _
               "Verificati fisierul si incercati din nou." & vbNewLine & _
               "Puteti folosi sablonul din folderul 'Sabloane'.", _
               vbCritical, "Eroare Structura Fisier"
        WriteLog "Import Stoc Depozit", "Structura fisierului invalida: " & filePath, "EROARE"
        Exit Sub
    End If

    ' Confirmare inainte de import
    Dim rowCount As Long
    rowCount = CountDataRows(filePath)

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Se vor importa " & rowCount & " produse in sheet-ul Stoc Depozit." & vbNewLine & vbNewLine & _
                    "Datele existente vor fi inlocuite." & vbNewLine & vbNewLine & _
                    "Continuati?", _
                    vbQuestion + vbYesNo, "Confirmare Import Stoc")

    If answer = vbNo Then Exit Sub

    ' Executa importul
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ClearSheetData SHEET_STOC
    Dim importedRows As Long
    importedRows = ImportDataFromFile(filePath, SHEET_STOC)

    ' Actualizeaza statusul pe Meniu
    UpdateImportStatus SHEET_STOC, importedRows

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Stoc Depozit", "Importate " & importedRows & " randuri din: " & filePath, "OK"

    MsgBox "Import realizat cu succes!" & vbNewLine & vbNewLine & _
           "Produse importate: " & importedRows, _
           vbInformation, "Import Stoc Depozit"

    ' Navigheaza la sheet-ul cu date
    ThisWorkbook.Sheets(SHEET_STOC).Activate

    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WriteLog "Import Stoc Depozit", "Eroare: " & Err.Description, "EROARE"
    MsgBox "A aparut o eroare la import:" & vbNewLine & Err.Description, _
           vbCritical, "Eroare Import"
End Sub

'==============================================================================
' IMPORT VANZARI MAGAZIN
'==============================================================================
Public Sub ImportVanzariMagazin()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectati fisierul cu VANZARI MAGAZIN")
    If filePath = "" Then Exit Sub

    ' Validare structura
    If Not ValidateFileStructure(filePath, HEADER_VANZARI_COL1, HEADER_VANZARI_COL2) Then
        MsgBox "Fisierul selectat NU are structura corecta!" & vbNewLine & vbNewLine & _
               "Structura asteptata (primele 2 coloane, randul 1):" & vbNewLine & _
               "  Coloana A: " & HEADER_VANZARI_COL1 & vbNewLine & _
               "  Coloana B: " & HEADER_VANZARI_COL2 & vbNewLine & vbNewLine & _
               "Verificati fisierul si incercati din nou." & vbNewLine & _
               "Puteti folosi sablonul din folderul 'Sabloane'.", _
               vbCritical, "Eroare Structura Fisier"
        WriteLog "Import Vanzari Magazin", "Structura fisierului invalida: " & filePath, "EROARE"
        Exit Sub
    End If

    ' Confirmare
    Dim rowCount As Long
    rowCount = CountDataRows(filePath)

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Se vor importa " & rowCount & " produse in sheet-ul Vanzari Magazin." & vbNewLine & vbNewLine & _
                    "Datele existente vor fi inlocuite." & vbNewLine & vbNewLine & _
                    "Continuati?", _
                    vbQuestion + vbYesNo, "Confirmare Import Vanzari")

    If answer = vbNo Then Exit Sub

    ' Executa importul
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ClearSheetData SHEET_VANZARI
    Dim importedRows As Long
    importedRows = ImportDataFromFile(filePath, SHEET_VANZARI)

    UpdateImportStatus SHEET_VANZARI, importedRows

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Vanzari Magazin", "Importate " & importedRows & " randuri din: " & filePath, "OK"

    MsgBox "Import realizat cu succes!" & vbNewLine & vbNewLine & _
           "Produse importate: " & importedRows, _
           vbInformation, "Import Vanzari Magazin"

    ThisWorkbook.Sheets(SHEET_VANZARI).Activate

    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WriteLog "Import Vanzari Magazin", "Eroare: " & Err.Description, "EROARE"
    MsgBox "A aparut o eroare la import:" & vbNewLine & Err.Description, _
           vbCritical, "Eroare Import"
End Sub

'==============================================================================
' STERGE TOATE DATELE IMPORTATE
'==============================================================================
Public Sub StergeToateDatele()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("ATENTIE! Se vor sterge TOATE datele importate:" & vbNewLine & vbNewLine & _
                    "  - Stoc Depozit" & vbNewLine & _
                    "  - Vanzari Magazin" & vbNewLine & vbNewLine & _
                    "Aceasta actiune este ireversibila!" & vbNewLine & vbNewLine & _
                    "Sigur doriti sa continuati?", _
                    vbExclamation + vbYesNo, "Confirmare Stergere")

    If answer = vbNo Then Exit Sub

    ' A doua confirmare pentru siguranta
    answer = MsgBox("Ultima confirmare: Sigur doriti sa STERGETI toate datele?", _
                    vbCritical + vbYesNo, "Confirmare Finala")

    If answer = vbNo Then Exit Sub

    ClearSheetData SHEET_STOC
    ClearSheetData SHEET_VANZARI

    ' Resetare status
    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)

    wsMeniu.Range("C20").Value = "Neincarcat"
    wsMeniu.Range("C20").Font.Color = RGB(198, 40, 40)

    wsMeniu.Range("C21").Value = "Neincarcat"
    wsMeniu.Range("C21").Font.Color = RGB(198, 40, 40)

    wsMeniu.Range("C22").Value = "-"

    WriteLog "Stergere Date", "Toate datele importate au fost sterse.", "OK"

    wsMeniu.Activate

    MsgBox "Toate datele au fost sterse cu succes!", _
           vbInformation, "Stergere Completa"
End Sub

'==============================================================================
' FUNCTII AUXILIARE (PRIVATE)
'==============================================================================

'------------------------------------------------------------------------------
' Deschide dialogul de selectare fisier
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Valideaza structura fisierului importat
' Verifica daca primele 2 coloane din randul 1 au headerele asteptate
'------------------------------------------------------------------------------
Private Function ValidateFileStructure(ByVal filePath As String, _
                                        ByVal expectedCol1 As String, _
                                        ByVal expectedCol2 As String) As Boolean
    ValidateFileStructure = False

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    ' Deschidem fisierul in mod read-only si fara update links
    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Set wsSource = wbSource.Sheets(1)

    ' Citim headerele si le curatam de spatii
    Dim h1 As String, h2 As String
    h1 = CleanHeader(CStr(wsSource.Cells(1, 1).Value))
    h2 = CleanHeader(CStr(wsSource.Cells(1, 2).Value))

    ' Verificam potrivirea (case-insensitive)
    If LCase(h1) = LCase(expectedCol1) And LCase(h2) = LCase(expectedCol2) Then
        ValidateFileStructure = True
    End If

    ' Verificare suplimentara: trebuie sa aiba cel putin 1 rand de date
    If ValidateFileStructure Then
        If wsSource.Cells(2, 1).Value = "" And wsSource.Cells(2, 2).Value = "" Then
            ValidateFileStructure = False
            MsgBox "Fisierul are structura corecta dar nu contine date!", _
                   vbExclamation, "Fisier Gol"
        End If
    End If

    wbSource.Close SaveChanges:=False
End Function

'------------------------------------------------------------------------------
' Curata un header de spatii, BOM, caractere invizibile
'------------------------------------------------------------------------------
Private Function CleanHeader(ByVal s As String) As String
    s = Trim(s)
    ' Elimina BOM (Byte Order Mark) daca exista
    If Len(s) > 0 Then
        If AscW(Left(s, 1)) = 65279 Then s = Mid(s, 2)
    End If
    ' Elimina spatii non-breaking
    s = Replace(s, Chr(160), "")
    s = Trim(s)
    CleanHeader = s
End Function

'------------------------------------------------------------------------------
' Numara randurile de date din fisierul sursa (fara header)
'------------------------------------------------------------------------------
Private Function CountDataRows(ByVal filePath As String) As Long
    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Set wsSource = wbSource.Sheets(1)

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow <= 1 Then
        CountDataRows = 0
    Else
        CountDataRows = lastRow - 1
    End If

    wbSource.Close SaveChanges:=False
End Function

'------------------------------------------------------------------------------
' Importa datele din fisierul sursa in sheet-ul destinatie
' Returneaza numarul de randuri importate
'------------------------------------------------------------------------------
Private Function ImportDataFromFile(ByVal filePath As String, _
                                     ByVal destSheetName As String) As Long
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet

    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Set wsSource = wbSource.Sheets(1)
    Set wsDest = ThisWorkbook.Sheets(destSheetName)

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        wbSource.Close SaveChanges:=False
        ImportDataFromFile = 0
        Exit Function
    End If

    ' Copiem datele (fara header) - doar primele 2 coloane
    Dim sourceRange As Range
    Set sourceRange = wsSource.Range("A2:B" & lastRow)

    Dim destRange As Range
    Set destRange = wsDest.Range("A2:B" & (lastRow))

    sourceRange.Copy
    destRange.PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' Formatam datele importate
    Dim r As Long
    For r = 2 To lastRow
        wsDest.Cells(r, 1).HorizontalAlignment = xlCenter
        wsDest.Cells(r, 2).HorizontalAlignment = xlCenter
        wsDest.Cells(r, 2).NumberFormat = "#,##0.00"
    Next r

    ' Actualizam auto-filter
    If wsDest.AutoFilterMode Then wsDest.AutoFilterMode = False
    wsDest.Range("A1:B" & lastRow).AutoFilter

    wbSource.Close SaveChanges:=False

    ImportDataFromFile = lastRow - 1
End Function

'------------------------------------------------------------------------------
' Sterge datele dintr-un sheet (pastreaza headerul)
'------------------------------------------------------------------------------
Private Sub ClearSheetData(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 1 Then
        ws.Range("A2:B" & lastRow).Clear
    End If
End Sub

'------------------------------------------------------------------------------
' Actualizeaza statusul de import pe sheet-ul Meniu
'------------------------------------------------------------------------------
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
    statusCell.Font.Color = RGB(46, 125, 50) ' Verde

    ' Actualizeaza data ultimei modificari
    wsMeniu.Range("C22").Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    wsMeniu.Range("C22").Font.Color = RGB(51, 51, 51)
End Sub

'------------------------------------------------------------------------------
' Scrie o intrare in sheet-ul de Log
'------------------------------------------------------------------------------
Private Sub WriteLog(ByVal operatiune As String, ByVal detalii As String, ByVal status As String)
    On Error Resume Next

    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Sheets(SHEET_LOG)

    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    wsLog.Cells(nextRow, 2).Value = operatiune
    wsLog.Cells(nextRow, 3).Value = detalii
    wsLog.Cells(nextRow, 4).Value = status

    ' Colorare status
    If status = "OK" Then
        wsLog.Cells(nextRow, 4).Font.Color = RGB(46, 125, 50)
    ElseIf status = "EROARE" Then
        wsLog.Cells(nextRow, 4).Font.Color = RGB(198, 40, 40)
    End If

    wsLog.Cells(nextRow, 1).HorizontalAlignment = xlCenter
    wsLog.Cells(nextRow, 4).HorizontalAlignment = xlCenter

    On Error GoTo 0
End Sub
