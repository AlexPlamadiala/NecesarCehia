Attribute VB_Name = "ModImport"
'==============================================================================
' MODUL: ModImport
' DESCRIERE: Funcționalități de import pentru stoc depozit și vânzări magazin
'            cu validare strictă a structurii fișierelor importate.
'==============================================================================
Option Explicit

' ── Constante pentru validare structură ──────────────────────────────────────
Private Const HEADER_STOC_COL1 As String = "CodProdus"
Private Const HEADER_STOC_COL2 As String = "Stoc"
Private Const HEADER_VANZARI_COL1 As String = "CodProdus"
Private Const HEADER_VANZARI_COL2 As String = "Cantitate"

Private Const SHEET_STOC As String = "StocDepozit"
Private Const SHEET_VANZARI As String = "VanzariMagazin"
Private Const SHEET_MENIU As String = "Meniu"
Private Const SHEET_LOG As String = "Log"

'==============================================================================
' IMPORT STOC DEPOZIT
'==============================================================================
Public Sub ImportStocDepozit()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectați fișierul cu STOC DEPOZIT")
    If filePath = "" Then Exit Sub

    ' Validare structură
    If Not ValidateFileStructure(filePath, HEADER_STOC_COL1, HEADER_STOC_COL2) Then
        MsgBox "Fișierul selectat NU are structura corectă!" & vbNewLine & vbNewLine & _
               "Structura așteptată (primele 2 coloane, rândul 1):" & vbNewLine & _
               "  Coloana A: " & HEADER_STOC_COL1 & vbNewLine & _
               "  Coloana B: " & HEADER_STOC_COL2 & vbNewLine & vbNewLine & _
               "Verificați fișierul și încercați din nou." & vbNewLine & _
               "Puteți folosi șablonul din folderul 'Sabloane'.", _
               vbCritical, "Eroare Structură Fișier"
        WriteLog "Import Stoc Depozit", "Structura fișierului invalidă: " & filePath, "EROARE"
        Exit Sub
    End If

    ' Confirmare înainte de import
    Dim rowCount As Long
    rowCount = CountDataRows(filePath)

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Se vor importa " & rowCount & " produse în sheet-ul Stoc Depozit." & vbNewLine & vbNewLine & _
                    "Datele existente vor fi înlocuite." & vbNewLine & vbNewLine & _
                    "Continuați?", _
                    vbQuestion + vbYesNo, "Confirmare Import Stoc")

    If answer = vbNo Then Exit Sub

    ' Execută importul
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ClearSheetData SHEET_STOC
    Dim importedRows As Long
    importedRows = ImportDataFromFile(filePath, SHEET_STOC)

    ' Actualizează statusul pe Meniu
    UpdateImportStatus SHEET_STOC, importedRows

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Stoc Depozit", "Importate " & importedRows & " rânduri din: " & filePath, "OK"

    MsgBox "Import realizat cu succes!" & vbNewLine & vbNewLine & _
           "Produse importate: " & importedRows, _
           vbInformation, "Import Stoc Depozit"

    ' Navighează la sheet-ul cu date
    ThisWorkbook.Sheets(SHEET_STOC).Activate

    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WriteLog "Import Stoc Depozit", "Eroare: " & Err.Description, "EROARE"
    MsgBox "A apărut o eroare la import:" & vbNewLine & Err.Description, _
           vbCritical, "Eroare Import"
End Sub

'==============================================================================
' IMPORT VÂNZĂRI MAGAZIN
'==============================================================================
Public Sub ImportVanzariMagazin()
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = SelectImportFile("Selectați fișierul cu VÂNZĂRI MAGAZIN")
    If filePath = "" Then Exit Sub

    ' Validare structură
    If Not ValidateFileStructure(filePath, HEADER_VANZARI_COL1, HEADER_VANZARI_COL2) Then
        MsgBox "Fișierul selectat NU are structura corectă!" & vbNewLine & vbNewLine & _
               "Structura așteptată (primele 2 coloane, rândul 1):" & vbNewLine & _
               "  Coloana A: " & HEADER_VANZARI_COL1 & vbNewLine & _
               "  Coloana B: " & HEADER_VANZARI_COL2 & vbNewLine & vbNewLine & _
               "Verificați fișierul și încercați din nou." & vbNewLine & _
               "Puteți folosi șablonul din folderul 'Sabloane'.", _
               vbCritical, "Eroare Structură Fișier"
        WriteLog "Import Vânzări Magazin", "Structura fișierului invalidă: " & filePath, "EROARE"
        Exit Sub
    End If

    ' Confirmare
    Dim rowCount As Long
    rowCount = CountDataRows(filePath)

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Se vor importa " & rowCount & " produse în sheet-ul Vânzări Magazin." & vbNewLine & vbNewLine & _
                    "Datele existente vor fi înlocuite." & vbNewLine & vbNewLine & _
                    "Continuați?", _
                    vbQuestion + vbYesNo, "Confirmare Import Vânzări")

    If answer = vbNo Then Exit Sub

    ' Execută importul
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ClearSheetData SHEET_VANZARI
    Dim importedRows As Long
    importedRows = ImportDataFromFile(filePath, SHEET_VANZARI)

    UpdateImportStatus SHEET_VANZARI, importedRows

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    WriteLog "Import Vânzări Magazin", "Importate " & importedRows & " rânduri din: " & filePath, "OK"

    MsgBox "Import realizat cu succes!" & vbNewLine & vbNewLine & _
           "Produse importate: " & importedRows, _
           vbInformation, "Import Vânzări Magazin"

    ThisWorkbook.Sheets(SHEET_VANZARI).Activate

    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    WriteLog "Import Vânzări Magazin", "Eroare: " & Err.Description, "EROARE"
    MsgBox "A apărut o eroare la import:" & vbNewLine & Err.Description, _
           vbCritical, "Eroare Import"
End Sub

'==============================================================================
' ȘTERGE TOATE DATELE IMPORTATE
'==============================================================================
Public Sub StergeToateDatele()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("ATENȚIE! Se vor șterge TOATE datele importate:" & vbNewLine & vbNewLine & _
                    "  • Stoc Depozit" & vbNewLine & _
                    "  • Vânzări Magazin" & vbNewLine & vbNewLine & _
                    "Această acțiune este ireversibilă!" & vbNewLine & vbNewLine & _
                    "Sigur doriți să continuați?", _
                    vbExclamation + vbYesNo, "Confirmare Ștergere")

    If answer = vbNo Then Exit Sub

    ' A doua confirmare pentru siguranță
    answer = MsgBox("Ultima confirmare: Sigur doriți să ȘTERGEȚI toate datele?", _
                    vbCritical + vbYesNo, "Confirmare Finală")

    If answer = vbNo Then Exit Sub

    ClearSheetData SHEET_STOC
    ClearSheetData SHEET_VANZARI

    ' Resetare status
    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets(SHEET_MENIU)

    wsMeniu.Range("C20").Value = Chr(8856) & " Neîncărcat"
    wsMeniu.Range("C20").Font.Color = RGB(198, 40, 40)

    wsMeniu.Range("C21").Value = Chr(8856) & " Neîncărcat"
    wsMeniu.Range("C21").Font.Color = RGB(198, 40, 40)

    wsMeniu.Range("C22").Value = "-"

    WriteLog "Ștergere Date", "Toate datele importate au fost șterse.", "OK"

    wsMeniu.Activate

    MsgBox "Toate datele au fost șterse cu succes!", _
           vbInformation, "Ștergere Completă"
End Sub

'==============================================================================
' FUNCȚII AUXILIARE (PRIVATE)
'==============================================================================

'──────────────────────────────────────────────────────────────────────────────
' Deschide dialogul de selectare fișier
'──────────────────────────────────────────────────────────────────────────────
Private Function SelectImportFile(ByVal dialogTitle As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = dialogTitle
        .Filters.Clear
        .Filters.Add "Fișiere Excel", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"

        If .Show = -1 Then
            SelectImportFile = .SelectedItems(1)
        Else
            SelectImportFile = ""
        End If
    End With
End Function

'──────────────────────────────────────────────────────────────────────────────
' Validează structura fișierului importat
' Verifică dacă primele 2 coloane din rândul 1 au headerele așteptate
'──────────────────────────────────────────────────────────────────────────────
Private Function ValidateFileStructure(ByVal filePath As String, _
                                        ByVal expectedCol1 As String, _
                                        ByVal expectedCol2 As String) As Boolean
    ValidateFileStructure = False

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    ' Deschidem fișierul în mod read-only și fără update links
    Application.DisplayAlerts = False
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    Set wsSource = wbSource.Sheets(1)

    ' Citim headerele și le curățăm de spații
    Dim h1 As String, h2 As String
    h1 = CleanHeader(CStr(wsSource.Cells(1, 1).Value))
    h2 = CleanHeader(CStr(wsSource.Cells(1, 2).Value))

    ' Verificăm potrivirea (case-insensitive)
    If LCase(h1) = LCase(expectedCol1) And LCase(h2) = LCase(expectedCol2) Then
        ValidateFileStructure = True
    End If

    ' Verificare suplimentară: trebuie să aibă cel puțin 1 rând de date
    If ValidateFileStructure Then
        If wsSource.Cells(2, 1).Value = "" And wsSource.Cells(2, 2).Value = "" Then
            ValidateFileStructure = False
            MsgBox "Fișierul are structura corectă dar nu conține date!", _
                   vbExclamation, "Fișier Gol"
        End If
    End If

    wbSource.Close SaveChanges:=False
End Function

'──────────────────────────────────────────────────────────────────────────────
' Curăță un header de spații, BOM, caractere invizibile
'──────────────────────────────────────────────────────────────────────────────
Private Function CleanHeader(ByVal s As String) As String
    s = Trim(s)
    ' Elimină BOM (Byte Order Mark) dacă există
    If Len(s) > 0 Then
        If AscW(Left(s, 1)) = 65279 Then s = Mid(s, 2)
    End If
    ' Elimină spații non-breaking
    s = Replace(s, Chr(160), "")
    s = Trim(s)
    CleanHeader = s
End Function

'──────────────────────────────────────────────────────────────────────────────
' Numără rândurile de date din fișierul sursă (fără header)
'──────────────────────────────────────────────────────────────────────────────
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

'──────────────────────────────────────────────────────────────────────────────
' Importă datele din fișierul sursă în sheet-ul destinație
' Returnează numărul de rânduri importate
'──────────────────────────────────────────────────────────────────────────────
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

    ' Copiem datele (fără header) - doar primele 2 coloane
    Dim sourceRange As Range
    Set sourceRange = wsSource.Range("A2:B" & lastRow)

    Dim destRange As Range
    Set destRange = wsDest.Range("A2:B" & (lastRow))

    sourceRange.Copy
    destRange.PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' Formatăm datele importate
    Dim r As Long
    For r = 2 To lastRow
        wsDest.Cells(r, 1).HorizontalAlignment = xlCenter
        wsDest.Cells(r, 2).HorizontalAlignment = xlCenter
        wsDest.Cells(r, 2).NumberFormat = "#,##0.00"
    Next r

    ' Actualizăm auto-filter
    If wsDest.AutoFilterMode Then wsDest.AutoFilterMode = False
    wsDest.Range("A1:B" & lastRow).AutoFilter

    wbSource.Close SaveChanges:=False

    ImportDataFromFile = lastRow - 1
End Function

'──────────────────────────────────────────────────────────────────────────────
' Șterge datele dintr-un sheet (păstrează headerul)
'──────────────────────────────────────────────────────────────────────────────
Private Sub ClearSheetData(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow > 1 Then
        ws.Range("A2:B" & lastRow).Clear
    End If
End Sub

'──────────────────────────────────────────────────────────────────────────────
' Actualizează statusul de import pe sheet-ul Meniu
'──────────────────────────────────────────────────────────────────────────────
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

    statusCell.Value = Chr(10004) & " Încărcat (" & rowCount & " produse)"
    statusCell.Font.Color = RGB(46, 125, 50) ' Verde

    ' Actualizează data ultimei modificări
    wsMeniu.Range("C22").Value = Format(Now, "dd.mm.yyyy hh:nn:ss")
    wsMeniu.Range("C22").Font.Color = RGB(51, 51, 51)
End Sub

'──────────────────────────────────────────────────────────────────────────────
' Scrie o intrare în sheet-ul de Log
'──────────────────────────────────────────────────────────────────────────────
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
