Attribute VB_Name = "Utilities"

Public Function unEnter(val As String, Optional validator As Variant, Optional validatorStr As Variant, Optional allowDuplicates As Variant) As Variant
'return one of following:
'a) original val string if it doesn't contain ENTER sign
'b) string of unique value if there is only 1 unique value in original string
'c) array of unique strings if there are more than 1
'd) array of all strings if allowDuplicates is True
'if validator is assigned, it will check if in validator on the same level there's validatorStr. If there's not, it won't add  item to array of values
Dim v() As String
Dim dup As Boolean
Dim x() As String
Dim values() As String
Dim i As Integer
Dim n As Integer
Dim found As Boolean

On Error GoTo err_trap

If IsMissing(allowDuplicates) Then
    dup = False
Else
    dup = allowDuplicates
End If

If InStr(1, val, vbLf) > 0 Then
    v = Split(val, vbLf)
    If Not IsMissing(validator) Then
        x = Split(validator, vbLf)
        For i = LBound(v) To UBound(v)
            If InStr(1, x(i), validatorStr, vbTextCompare) <> 0 Then
                found = False
                If isArrayEmpty(values) Then
                    ReDim values(0) As String
                    values(0) = Replace(v(i), vbCr, "")
                Else
                    If dup = False Then
                        For n = LBound(values) To UBound(values)
                            If IsNumeric(values(n)) And IsNumeric(v(i)) Then
                                If CDbl(values(n)) = CDbl(v(i)) Then
                                    found = True
                                    Exit For
                                End If
                            Else
                                If values(n) = v(i) Then
                                    found = True
                                    Exit For
                                End If
                            End If
                        Next n
                    End If
                    If found = False Then
                        ReDim Preserve values(UBound(values) + 1) As String
                        values(UBound(values)) = Replace(v(i), vbCr, "")
                    End If
                End If
            End If
        Next i
    Else
        For i = LBound(v) To UBound(v)
            found = False
            If isArrayEmpty(values) Then
                ReDim values(0) As String
                values(0) = Replace(v(i), vbCr, "")
            Else
                If dup = False Then
                    For n = LBound(values) To UBound(values)
                        If IsNumeric(values(n)) And IsNumeric(v(i)) Then
                            If CDbl(values(n)) = CDbl(v(i)) Then
                                found = True
                                Exit For
                            End If
                        Else
                            If values(n) = v(i) Then
                                found = True
                                Exit For
                            End If
                        End If
                    Next n
                End If
                If found = False Then
                    ReDim Preserve values(UBound(values) + 1) As String
                    values(UBound(values)) = Replace(v(i), vbCr, "")
                End If
            End If
        Next i
    End If
    If UBound(values) = 0 Then
        unEnter = values(0)
    Else
        unEnter = values
    End If
Else
    unEnter = val
End If

exit_here:
Exit Function

err_trap:
MsgBox "Error in ""unEnter"". Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Function

Public Function inCollection(ind As String, col As Collection) As Boolean
Dim v As Variant
Dim isError As Boolean

isError = False

On Error GoTo err_trap

If VarType(col(ind)) = vbObject Then
    Set v = col(ind)
Else
    v = col(ind)
End If


exit_here:
If isError Then
    inCollection = False
Else
    inCollection = True
End If
Exit Function

err_trap:
isError = True
Resume exit_here

End Function

Public Function cell2letter(c As Integer) As String
Dim arr() As String
With ActiveWorkbook.Sheets("Pracownicy kalendarz")
    arr = Split(.Cells(1, c).Address(True, False), "$")
    cell2letter = arr(0) & ":" & arr(0)
End With
End Function

Public Function col2Letter(c As Integer) As String
Dim arr() As String
With ActiveWorkbook.Sheets(1)
    arr = Split(.Cells(1, c).Address(True, False), "$")
    col2Letter = arr(0)
End With
End Function

Public Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Public Function isArrayEmpty(parArray As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
  If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False

End Function

Public Function Week2Date(weekNo As Long, Optional ByVal Yr As Long = 0, Optional ByVal DOW As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbUseSystemDayOfWeek, Optional ByVal FWOY As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbUseSystem) As Date
 ' Returns First Day of week
 Dim Jan1 As Date
 Dim Sub1 As Boolean
 Dim ret As Date

 If Yr = 0 Then
   Jan1 = VBA.DateSerial(VBA.year(VBA.Date()), 1, 1)
 Else
   Jan1 = VBA.DateSerial(Yr, 1, 1)
 End If
 Sub1 = (VBA.Format(Jan1, "ww", DOW, FWOY) = 1)
 ret = VBA.DateAdd("ww", weekNo + Sub1, Jan1)
 ret = ret - VBA.Weekday(ret, DOW) + 1
 Week2Date = ret
End Function


Public Function IsoWeekNumber(InDate As Date) As Long
    IsoWeekNumber = DatePart("ww", InDate, vbMonday, vbFirstFourDays)
End Function

Public Sub HideEmptyColumns(Optional row As Variant)

On Error GoTo err_trap

Dim i As Integer
Dim lngLastColumn As Long
' Get last cell
lngLastColumn = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

Application.ScreenUpdating = False
For i = 1 To lngLastColumn
    If Application.WorksheetFunction.CountA(Columns(i)) = 0 And IsMissing(row) Then
        Columns(i).Hidden = True
    ElseIf Not IsMissing(row) And row > 0 Then
        If ActiveSheet.Cells(row, i) = "" Or ActiveSheet.Cells(row, i) = 0 Then
            Columns(i).Hidden = True
        End If
    End If
Next i

exit_here:
Application.ScreenUpdating = True
Exit Sub

err_trap:
MsgBox "Błąd. Szczegóły: " & Err.Description, vbCritical + vbOKOnly, "Błąd"
Resume exit_here

End Sub

Public Sub ShowAllColumns()
On Error GoTo err_trap

Dim i As Integer
Dim lngLastColumn As Long
' Get last cell
lngLastColumn = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

Application.ScreenUpdating = False
For i = 1 To lngLastColumn
    If Columns(i).Hidden Then
        Columns(i).Hidden = False
    End If
Next i

exit_here:
Application.ScreenUpdating = True
Exit Sub

err_trap:
MsgBox "Błąd. Szczegóły: " & Err.Description, vbCritical + vbOKOnly, "Błąd"
Resume exit_here

End Sub
