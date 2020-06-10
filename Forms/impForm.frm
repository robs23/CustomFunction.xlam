VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} impForm 
   Caption         =   "Wybierz typ raportu"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   OleObjectBlob   =   "impForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "impForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
impForm.Hide
End Sub

Private Sub btnOK_Click()
If Me.cmbReport.Value = "Grafik produkcyjny" Then
    Me.Hide
    saveProductionPlan
ElseIf Me.cmbReport.Value = "Lista przypisań atrybutów do artykułów" Then
    Me.Hide
    updateProperty
ElseIf Me.cmbReport.Value = "Zestawienie obrotów wg artykułów" Then
    Me.Hide
    exportPW_WZ
ElseIf Me.cmbReport.Value = "Powiązania operacji" Then
    Me.Hide
    importConnections
ElseIf Me.cmbReport.Value = "COOIS" Then
    Me.Hide
    importBatch2order
ElseIf Me.cmbReport.Value = "Zestawienie ilości wyprodukowanej w zleceniu" Then
    Me.Hide
    importMesQuantities
ElseIf Me.cmbReport.Value = "Requirements View" Then
    Me.Hide
    importReqs
ElseIf Me.cmbReport.Value = "BOM overview" Then
    Me.Hide
    importBOM
ElseIf Me.cmbReport.Value = "Component scrap" Then
    Me.Hide
    importComponentScrap
ElseIf Me.cmbReport.Value = "Production order variance" Then
    importComponentUsage
ElseIf Me.cmbReport.Value = "Costing data (ZCOMM_HU)" Then
    importCosting
ElseIf Me.cmbReport.Value = "INET's OrderList" Then
    'importInetsOrderList
ElseIf Me.cmbReport.Value = "Zlecenia z partią dosypaną" Then
    importRework
ElseIf Me.cmbReport.Value = "Ruchy powrotne z magazynu (MB51)" Then
    importMb51
Else
    MsgBox "Najpierw wybierz jedną z pozycji na liście", vbOKOnly + vbInformation, "Błąd"
End If
End Sub


Private Sub cmbReport_Change()
If IsNull(Me.cmbReport) Then
    Me.btnOK.Enabled = False
Else
    Me.btnOK.Enabled = True
End If
End Sub

Private Sub UserForm_Initialize()
Dim i As Integer
With Me.cmbReport
    For i = .ListCount - 1 To 0 Step -1
        .RemoveItem i
    Next i
    .AddItem "Grafik produkcyjny"
    .AddItem "Lista przypisań atrybutów do artykułów"
    .AddItem "Zestawienie obrotów wg artykułów"
    .AddItem "Powiązania operacji"
    .AddItem "COOIS"
    .AddItem "Zestawienie ilości wyprodukowanej w zleceniu"
    .AddItem "Requirements View"
    .AddItem "BOM overview"
    .AddItem "Component scrap"
    .AddItem "Production order variance"
    .AddItem "Costing data (ZCOMM_HU)"
    .AddItem "INET's OrderList"
    .AddItem "Zlecenia z partią dosypaną"
    .AddItem "Ruchy powrotne z magazynu (MB51)"
    Me.btnOK.Enabled = False
End With

End Sub
