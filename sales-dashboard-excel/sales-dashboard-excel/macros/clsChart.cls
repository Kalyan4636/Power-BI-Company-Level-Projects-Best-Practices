VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
CODE 1  -  Class Module named  clsChart
--------------------------------------------------------------------
Option Explicit

Public WithEvents cht As Chart
Public ChartKey As String

Private Sub cht_MouseDown(ByVal Button As Long, ByVal Shift As Long, _
                          ByVal x As Long, ByVal y As Long)
    Dim elemID As Long, arg1 As Long, arg2 As Long
    If Button <> xlPrimaryButton Then Exit Sub
    cht.GetChartElement x, y, elemID, arg1, arg2
    If elemID = xlSeries And arg2 > 0 Then
        ApplyChartClick ChartKey, CStr(cht.SeriesCollection(arg1).XValues(arg2))
    End If
End Sub
