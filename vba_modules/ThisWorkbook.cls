VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' EVENIMENT: Workbook_Open
' La deschiderea fișierului, navighează automat la sheet-ul Meniu
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next

    ' Activează sheet-ul Meniu la deschidere
    ThisWorkbook.Sheets("Meniu").Activate

    ' Poziționează selecția în zona neutră
    ThisWorkbook.Sheets("Meniu").Range("B2").Select

    Application.StatusBar = "Necesar Aprovizionare - Gata de lucru"

    On Error GoTo 0
End Sub
