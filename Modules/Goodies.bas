Attribute VB_Name = "Goodies"

Public Sub ShowEmptyColumns(control As IRibbonControl)
ShowAllColumns
End Sub

Public Sub HideTheEmptyColumns(control As IRibbonControl)
Dim i As Integer
Dim myCell As Range

If Selection.Rows.count = 1 And Selection.Columns.count = ActiveSheet.Columns.count Then
    For Each myCell In Selection
        i = myCell.row
        Exit For
    Next myCell
    If i > 0 Then
        HideEmptyColumns i
    Else
        HideEmptyColumns
    End If
End If

End Sub
