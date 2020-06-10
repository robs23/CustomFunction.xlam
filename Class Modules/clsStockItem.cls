VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStockItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public PlannedDate As Date
Public ProductId As Integer
Public LPlant As String
Public amount As Double

Public Sub Save()
AdoConn.Execute "INSERT INTO tbPlannedStock (PlannedDate, ProductId, Lplant, Amount, CreatedOn) VALUES ('" & PlannedDate & "', " & ProductId & ", '" & LPlant & "', " & amount & ",'" & Now & "')"
End Sub

