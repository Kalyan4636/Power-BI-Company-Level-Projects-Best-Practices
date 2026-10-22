Attribute VB_Name = "modCrossFilter"
Option Explicit

Public Handlers As Collection

Public Sub InitCrossFilter()
    Dim ws As Worksheet, co As ChartObject, h As clsChart
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set Handlers = New Collection
    For Each co In ws.ChartObjects
        Set h = New clsChart
        Set h.cht = co.Chart
        h.ChartKey = co.TopLeftCell.Address(False, False)
        Handlers.Add h
    Next co
End Sub

Private Sub Toggle(ByVal cel As Range, ByVal item As String)
    If cel.Value = item Then cel.Value = "All" Else cel.Value = item
End Sub

Public Sub ApplyChartClick(ByVal key As String, ByVal item As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If Len(item) = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Select Case key
        Case "B17"   ' monthly chart -> set From and To month
            If ws.Range("sel_FromMonth").Value = item And ws.Range("sel_ToMonth").Value = item Then
                ws.Range("sel_FromMonth").Value = "Jan": ws.Range("sel_ToMonth").Value = "Dec"
            Else
                ws.Range("sel_FromMonth").Value = item: ws.Range("sel_ToMonth").Value = item
            End If
        Case "L17": Toggle ws.Range("sel_Category"), item   ' category chart
        Case "L36": Toggle ws.Range("sel_Product"), item    ' top 5 products
        Case "B55": Toggle ws.Range("sel_Country"), item    ' top 10 countries
        Case "L56": Toggle ws.Range("sel_Channel"), item    ' sales channel
        Case "P56": Toggle ws.Range("sel_Customer"), item   ' customer type
    End Select
    Application.ScreenUpdating = True
End Sub

Public Sub ResetFilters()
    Dim nm As Variant
    For Each nm In Array("sel_Region", "sel_Country", "sel_Channel", "sel_Customer", "sel_Category", "sel_Product")
        ThisWorkbook.Worksheets("Dashboard").Range(nm).Value = "All"
    Next nm
    ThisWorkbook.Worksheets("Dashboard").Range("sel_FromMonth").Value = "Jan"
    ThisWorkbook.Worksheets("Dashboard").Range("sel_ToMonth").Value = "Dec"
End Sub
