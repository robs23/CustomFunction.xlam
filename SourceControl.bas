Attribute VB_Name = "SourceControl"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode(control As IRibbonControl)
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    Dim go As Boolean
    Dim wb As Workbook
    
    go = False
    
    If ActiveWorkbook Is Nothing Then
        res = MsgBox("Czy chcia³byœ eksportowaæ kod z tej wtyczki", vbQuestion + vbYesNo, "PotwierdŸ eksport wtyczki")
        If res = vbYes Then
            go = True
            Set wb = ThisWorkbook
        End If
    Else
        go = True
        Set wb = ActiveWorkbook
    End If
    
    If go Then
        directory = wb.path & "\VisualBasic"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In wb.VBProject.VBComponents
        If Not VBComponent.name = "Secrets" Then
        
            Select Case VBComponent.Type
                Case ClassModule, Document
                    extension = ".cls"
                Case Form
                    extension = ".frm"
                Case Module
                    extension = ".bas"
                Case Else
                    extension = ".txt"
            End Select
                
                    
            On Error Resume Next
            Err.Clear
            
            path = directory & "\" & VBComponent.name & extension
            Call VBComponent.Export(path)
            
            If Err.Number <> 0 Then
                Call MsgBox("Failed to export " & VBComponent.name & " to " & path, vbCritical)
            Else
                count = count + 1
                Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & path
            End If
    
            On Error GoTo 0
        End If
    Next
    
    MsgBox "Successfully exported " & CStr(count) & " VBA files to " & directory
    End If
    
End Sub

Sub ImportVisualBasicCode()
 
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim directory As String
     
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
    Set oFolder = oFSO.GetFolder(ActiveWorkbook.path & "\VisualBasic")
     
    For Each oFile In oFolder.Files
     
        directory = ActiveWorkbook.path & "\VisualBasic\" & oFile.name
        ActiveWorkbook.VBProject.VBComponents.Import directory
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to import " & oFile.name, vbCritical)
        End If
     
    Next oFile
 
End Sub
