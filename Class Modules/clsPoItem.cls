VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public poItemId As Long
Private pIndex As Long
Private pAmount As Double

Public Sub initialize(index As Long, amount As Double)
pIndex = index
pAmount = amount
End Sub

Public Property Get index() As Long
index = pIndex
End Property

Public Property Get amount() As Double
amount = pAmount
End Property

Public Sub Save(poId As Integer)
Dim rs As ADODB.Recordset

Set rs = AdoConn.Execute("INSERT INTO tbPoItem (ProductId, Amount, PoId) VALUES (" & zfins(CStr(index)).zfinId & "," & amount & "," & poId & ");SELECT SCOPE_IDENTITY()")
Set rs = rs.NextRecordset
poItemId = rs.Fields(0).Value
rs.Close
Set rs = Nothing
End Sub
