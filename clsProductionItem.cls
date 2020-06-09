VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProductionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public week As Integer
Public year As Integer
Public ProductId As Integer
Public LPlant As String
Public amount As Double

Public Sub Save()
AdoConn.Execute "INSERT INTO tbPlannedProduction (Week, Year, ProductId, Lplant, Amount, CreatedOn) VALUES (" & week & ", " & year & ", " & ProductId & ", '" & LPlant & "', " & amount & ",'" & Now & "')"
End Sub
