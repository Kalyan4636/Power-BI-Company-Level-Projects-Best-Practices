' Paste into the ThisWorkbook code window (cannot be imported as a file)
Option Explicit

Private Sub Workbook_Open()
    With ThisWorkbook.Worksheets("Dashboard")
        .Unprotect
        ' leave charts clickable, keep cells protected (filter cells stay editable)
        .Protect DrawingObjects:=False, Contents:=True, UserInterfaceOnly:=True
    End With
    InitCrossFilter
End Sub
