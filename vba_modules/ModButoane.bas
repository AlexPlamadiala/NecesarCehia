Attribute VB_Name = "ModButoane"
'==============================================================================
' MODUL: ModButoane
' DESCRIERE: Rutine de inițializare - creează butoane pe sheet-ul Meniu
'            și configurează bara de meniu personalizată.
' UTILIZARE: Rulați o singură dată SubCreazaButoane() pentru a genera butoanele.
'==============================================================================
Option Explicit

'==============================================================================
' Creează butoanele pe sheet-ul Meniu ca Form Controls
' Rulați această rutină O SINGURĂ DATĂ după importul modulelor VBA
'==============================================================================
Public Sub CreazaButoane()
    On Error GoTo ErrHandler

    Dim wsMeniu As Worksheet
    Set wsMeniu = ThisWorkbook.Sheets("Meniu")

    wsMeniu.Activate

    ' Ștergem butoane existente
    Dim shp As Shape
    For Each shp In wsMeniu.Shapes
        If shp.Type = msoFormControl Then shp.Delete
    Next shp

    ' ── Buton 1: Import Stoc Depozit ──
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

    ' ── Buton 2: Import Vânzări Magazin ──
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

    ' ── Buton 3: Șterge Toate Datele ──
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

    MsgBox "Butoanele au fost create cu succes!" & vbNewLine & vbNewLine & _
           "Acum puteți folosi:" & vbNewLine & _
           "  • IMPORT STOC DEPOZIT" & vbNewLine & _
           "  • IMPORT VÂNZĂRI MAGAZIN" & vbNewLine & _
           "  • ȘTERGE TOATE DATELE", _
           vbInformation, "Configurare Completă"

    Exit Sub

ErrHandler:
    MsgBox "Eroare la crearea butoanelor: " & Err.Description, vbCritical
End Sub
