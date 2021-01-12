VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} chooseWeekForm 
   Caption         =   "Wybierz okres.."
   ClientHeight    =   1965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2655
   OleObjectBlob   =   "chooseWeekForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "chooseWeekForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()
dividerWeek = CInt(Me.txtWeek)
dividerYear = CInt(Me.txtYear)
Me.Hide
End Sub

Private Sub UserForm_Initialize()
On Error GoTo err_trap

Dim wb As Workbook
Dim name As String
Dim x() As String
Dim sYear As String
Dim sWeek As String

dividerWeek = 0
dividerYear = 0

Set wb = ActiveWorkbook
name = wb.name

x = Split(name, "_", , vbTextCompare)
If UBound(x) >= 1 Then
    If Len(x(1)) >= 4 Then
        sYear = Mid(x(1), 1, 4)
        If IsNumeric(sYear) Then
            Me.txtYear = sYear
        End If
    End If
End If
If Len(ActiveSheet.name) > 2 Then
    sWeek = Mid(ActiveSheet.name, Len(ActiveSheet.name) - 1, 2)
    If IsNumeric(sWeek) Then
        Me.txtWeek = CInt(sWeek)
    End If
End If

exit_here:
Exit Sub

err_trap:
Resume exit_here

End Sub
