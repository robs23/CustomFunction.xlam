VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pColumn As Integer
Private pPos As New Collection
Public shipmentId As Integer
Public ShipmentDate As Date
Public DeliveryNotes As String
Public Items As Collection

Property Let Column(val As Integer)
pColumn = val
End Property

Property Get Column() As Integer
Column = pColumn
End Property

Public Sub append(po As Variant)
pPos.Add po, po.LPlant
End Sub

Public Function getPos(Optional LPlant As Variant) As Variant
Dim po As clsPO
Dim bool As Boolean

If IsMissing(LPlant) Then
    Set getPos = pPos
Else
    For Each po In pPos
        If po.LPlant = CStr(LPlant) Then
            bool = True
            Set getPos = po
            Exit For
        End If
    Next po
    If Not bool Then
        'Po for given Lplant number is not present yet, return Nothing
        Set getPos = Nothing
    End If
End If
End Function

Public Function Exists(shipmentId As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim sSql As String

updateConnection

sSql = "SELECT * FROM tbPlannedShipments WHERE PlannedShipmentId=" & shipmentId & ";"
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If rs.EOF Then
    Exists = False
Else
    Exists = True
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
closeConnection
Exit Function

End Function

Public Sub SaveToDb()
Dim rs As ADODB.Recordset
Dim po As clsPO

updateConnection
On Error GoTo err_trap

If shipmentId = 0 Then
    'new shipment
    Set rs = AdoConn.Execute("INSERT INTO tbPlannedShipments (PlannedDate, DeliveryNotes, AddedOn) VALUES ('" & ShipmentDate & "','" & DeliveryNotes & "','" & Now & "');SELECT SCOPE_IDENTITY()")
    shipmentId = rs.Fields(0).Value
    ActiveWorkbook.Sheets("Plan").Cells(5, Column) = shipmentId
    rs.Close
    Set rs = Nothing
Else
    AdoConn.Execute "UPDATE tbPlannedShipments SET PlannedDate = '" & ShipmentDate & "', DeliveryNotes = '" & DeliveryNotes & "' WHERE PlannedShipmentId = " & shipmentId
End If

For Each po In pPos
        po.Save shipmentId
Next po
    
Exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Błąd w clsShipment.SaveToDb. " & Err.Description
Resume Exit_here

End Sub

