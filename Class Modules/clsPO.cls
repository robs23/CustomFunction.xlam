VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public poId As Integer
Private pPoNumber As Long
Private pPoDate As Date
Private pPurchasingOrg As String
Private pPurchasingGroup As String
Private pCompanyCode As String
Private pLplant As String
Private pUnit As String
Private pItems As New Collection
Private pShipmentId As Integer

Public Property Let Number(val As Long)
pPoNumber = val
End Property

Public Property Get Number() As Long
Number = pPoNumber
End Property

Public Property Let PoDate(val As Date)
If Weekday(val, vbMonday) = 5 Then
    pPoDate = DateAdd("d", 3, val)
Else
    pPoDate = DateAdd("d", 2, val)
End If

End Property

Public Property Get PoDate() As Date
PoDate = pPoDate
End Property

Public Property Let PurchasingOrg(val As String)
pPurchasingOrg = val
End Property

Public Property Get PurchasingOrg() As String
PurchasingOrg = pPurchasingOrg
End Property

Public Property Let PurchasingGroup(val As String)
pPurchasingGroup = val
End Property

Public Property Get PurchasingGroup() As String
PurchasingGroup = pPurchasingGroup
End Property

Public Property Let companyCode(val As String)
pCompanyCode = val
End Property

Public Property Get companyCode() As String
companyCode = pCompanyCode
End Property

Public Property Let LPlant(val As String)
pLplant = val
End Property

Public Property Get LPlant() As String
LPlant = pLplant
End Property

Public Property Let Unit(val As String)
pUnit = val
End Property

Public Property Get Unit() As String
Unit = pUnit
End Property

Public Property Let shipmentId(val As Integer)
pShipmentId = val
End Property

Public Property Get shipmentId() As Integer
shipmentId = pShipmentId
End Property

Public Sub append(item As clsPoItem)
pItems.Add item
End Sub

Public Function toString() As String
Dim str As String
Dim item As clsPoItem

str = "Data: " & pPoDate & vbNewLine
str = str & "Purchasing Organisation: " & pPurchasingOrg & vbNewLine
str = str & "Purchasing Group: " & pPurchasingGroup & vbNewLine
str = str & "CompanyCode: " & pCompanyCode & vbNewLine
str = str & "L-plant: " & pLplant & vbNewLine
str = str & "Unit: " & pUnit & vbNewLine
For Each item In pItems
    str = str & item.index & " - " & item.amount & vbNewLine
Next item
toString = str
End Function

Public Function getItems() As Collection
Set getItems = pItems
End Function


Public Sub Save(shId As Integer)
Dim rs As ADODB.Recordset
Dim poi As clsPoItem

Set rs = AdoConn.Execute("INSERT INTO tbPo (Lplant, shipmentId) VALUES ('" & LPlant & "'," & shId & ");SELECT SCOPE_IDENTITY()")
Set rs = rs.NextRecordset
poId = rs.Fields(0).Value
rs.Close
Set rs = Nothing

For Each poi In pItems
    poi.Save poId
Next poi
End Sub

