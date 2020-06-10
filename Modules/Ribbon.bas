Attribute VB_Name = "Ribbon"
Public rib As IRibbonUI
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)

Public Sub StoreObjRef(obj As Object, propertyName As String)
' Store an object reference
 Dim longObj As Long
 longObj = ObjPtr(obj)
 updateCustomProperty propertyName, longObj
End Sub
 
Function RetrieveObjRef(propertyName As String) As Object
' Retrieve the object reference
 Dim longObj As Long, obj As Object
 longObj = ThisWorkbook.CustomDocumentProperties(propertyName)
 CopyMemory obj, longObj, 4
 Set RetrieveObjRef = obj
End Function

Public Sub bbRib()
If rib Is Nothing Then
    Set rib = RetrieveObjRef("ribbon_ref")
End If
End Sub


Public Sub OnRibbonLoad(objRibbon As IRibbonUI)
    Set rib = objRibbon
    StoreObjRef rib, "ribbon_ref"
End Sub

Public Function propertyExists(name As String) As Boolean
Dim prop As DocumentProperty
propertyExists = False
For Each prop In ThisWorkbook.CustomDocumentProperties
    If prop.name = name Then
        propertyExists = True
        Exit For
    End If
Next prop
End Function

Public Sub updateCustomProperty(propName As String, propValue As Variant)

With ThisWorkbook.CustomDocumentProperties
    If propertyExists(propName) Then
        .item(propName).Value = propValue
    Else
        createCustomProperty propName, propValue
    End If
End With

End Sub

Public Sub createCustomProperty(theName As String, theValue As Variant)
Dim theType As Variant

Select Case VarType(theValue)
    Case 0 To 1
    theType = Null
    Case 2 To 3
    theType = msoPropertyTypeNumber
    Case 4 Or 5 Or 14
    theType = msoPropertyTypeFloat
    Case 7
    theType = msoPropertyTypeDate
    Case 8
    theType = msoPropertyTypeString
    Case 11
    theType = msoPropertyTypeBoolean
    Case Else
    theType = Null
End Select

If theType = Null Then
    MsgBox "Type of variable ""theValue"" passed to ""createCustomProperty"" could not be determined or is unsuported. No custom property has been created", vbOKOnly + vbExclamation
Else
    ThisWorkbook.CustomDocumentProperties.Add name:=theName, LinkToContent:=False, Type:=theType, Value:=theValue
'    MsgBox "Property " & theName & " has been created successfully and set to " & theValue, vbOKOnly + vbInformation


End If

End Sub



