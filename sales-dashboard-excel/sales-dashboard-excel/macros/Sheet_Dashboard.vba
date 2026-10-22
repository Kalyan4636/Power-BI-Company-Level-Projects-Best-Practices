' Paste into the 'Dashboard' sheet code window (cannot be imported as a file)
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim nameCell As Range, nm As String

    ' Reset button: ORDERS SHOWN box
    If Not Intersect(Target, Me.Range("R5:S6")) Is Nothing Then
        ResetFilters
        Me.Range("A1").Select
        Exit Sub
    End If

    ' Map tiles: B37:K46 (name row on top, value row below)
    If Not Intersect(Target, Me.Range("B37:K46")) Is Nothing Then
        If (Target.Row - 37) Mod 2 = 0 Then
            Set nameCell = Me.Cells(Target.Row, Target.Column)
        Else
            Set nameCell = Me.Cells(Target.Row - 1, Target.Column)
        End If
        nm = CStr(nameCell.Value)
        If Len(nm) > 0 Then
            If Me.Range("sel_Country").Value = nm Then
                Me.Range("sel_Country").Value = "All"
            Else
                Me.Range("sel_Country").Value = nm
            End If
        End If
        Me.Range("A1").Select   ' lets the same tile be clicked again
    End If
End Sub
