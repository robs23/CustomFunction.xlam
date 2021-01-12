Attribute VB_Name = "this"
'Public AdoConn As New ADODB.Connection
Public ScadaConn As New ADODB.Connection
'Public zfins As New Collection

Public Sub Snapshot()
Dim res As Variant
'check if saveable

If Not ActiveWorkbook.ReadOnly Then
    res = MsgBox("Czy chcesz zapisac dane do bazy?", vbQuestion + vbYesNo, "Zapisujemy?")
    If res = vbYes Then
        Dim shipments As New clsShipments
        If shipments.Scan Then
            shipments.Save
        Else
            Cancel = True
            MsgBox "Nie udalo sie zapisac danych z powodu bledów. Popraw dane i spróbuj jeszcze raz", vbOKOnly, "Lipa"
        End If
    End If
End If




End Sub


Public Function DropLSU(orginalBatch As String) As Double
Dim b As String
b = Replace(orginalBatch, "LSU", "", 1, , vbTextCompare)
DropLSU = CDbl(b)

End Function

Public Function conversionPossible(ind As Variant) As Variant
Dim lastRow As Long
Dim rng As Range
Dim sht As Worksheet
Dim bool As Variant
Dim c As Range
Dim lr As Range

bool = False

Set sht = ActiveWorkbook.Sheets("Baza danych")

Set lr = sht.Range("A2:A100").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues)

If Not lr Is Nothing Then
    Set rng = sht.Range("A2:A" & lr.Row)
    
    Set c = rng.Find(ind, searchorder:=xlByRows, SearchDirection:=xlNext, LookIn:=xlValues)
    If Not c Is Nothing Then
        If c.Offset(0, 3) > 0 Then
            bool = c.Offset(0, 2)
        End If
    End If
End If

conversionPossible = bool
End Function


Public Sub addOrdersToSAP()
Dim po As clsPO
Dim sht As Worksheet
Dim i As Integer
Dim total As Double
Dim isError As Boolean
Dim added As Integer
Dim sTime As Date
Dim eTime As Date
Dim poNumber As String
Dim y As Integer
Dim index As Long
Dim amount As Double
Dim poItem As clsPoItem
Dim pos As New Collection
Dim shipments As New Collection
Dim ship As clsShipment
Dim poStr As String
Dim Lplant As String
Dim newPo As Variant
Dim missingStr As String
Dim missingCount As Integer
Dim msgStr As String

On Error GoTo err_trap

Set sht = ActiveWorkbook.Sheets("Plan")
added = 0

sTime = Now

For i = 1 To 100
    If sht.Cells(6, 11 + i) = "Shipment" Then
        total = sht.Cells(11, 11 + i)
        If total > 0 Then
            poNumber = sht.Cells(9, 11 + i)
            If Len(poNumber) = 0 Then
                'we encountered first po that hasnt't been added yet
                Set ship = New clsShipment
                ship.Column = 11 + i
                shipments.Add ship, CStr(11 + i)
                For y = 15 To 114
                    index = sht.Cells(y, 1)
                    If index = 0 Then
                        Exit For
                    Else
                        amount = sht.Cells(y, 11 + i)
                        If amount > 0 Then
                            Lplant = sht.Range("G" & y)
                            'check if we have all data of location
                            If Not IsNull(purOrg(Lplant)) And Not IsNull(purGroup(Lplant)) And Not IsNull(companyCode(Lplant)) Then
                                Set newPo = ship.getPos(Lplant)
                                If newPo Is Nothing Then
                                    'create new one
                                    Set newPo = New clsPO
                                    With newPo
                                        .PurchasingOrg = purOrg(Lplant)
                                        .PurchasingGroup = purGroup(Lplant)
                                        .companyCode = companyCode(Lplant)
                                        .Lplant = Lplant
                                        .PoDate = sht.Cells(8, 11 + i)
                                        .Unit = sht.Range("E11")
                                    End With
                                    ship.append newPo
                                End If
                                Set poItem = New clsPoItem
                                poItem.initialize index, amount
                                newPo.append poItem
                            Else
                                missingCount = missingCount + 1
                                If Len(missingStr) = 0 Then
                                    missingStr = " zleceń z powodu niewypełnionych danych lokacji (arkusz ""Lokacje""). Brakujące lokacje: " & Lplant & ", "
                                Else
                                    missingStr = missingStr & Lplant & ", "
                                End If
                            End If
                        End If
                    End If
                Next y
            End If
        End If
    End If
Next i

For Each ship In shipments
    Set pos = ship.getPos
    For Each po In pos
        added = added + 1
        poStr = orderToPO(po)
        If Len(sht.Cells(9, ship.Column)) > 0 Then
            sht.Cells(9, ship.Column) = sht.Cells(9, ship.Column) & ", " & poStr
        Else
            sht.Cells(9, ship.Column) = poStr
        End If
    Next po
Next ship

eTime = Now

exit_here:
If Not isError Then
    msgStr = "Dodano " & added & " zleceń w czasie " & DateDiff("s", sTime, eTime) & " sek."
    If Len(missingStr) > 0 Then
        msgStr = msgStr & vbNewLine
        missingStr = Left(missingStr, Len(missingStr) - 2)
        msgStr = msgStr & "Pominięto " & missingCount & missingStr
    End If
    MsgBox msgStr
End If
Exit Sub

err_trap:
isError = True
MsgBox "Error in addOrdersToSAP. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


Public Function orderToPO(po As clsPO) As String

'setting the connection with sap:
Dim App, Connection, session As Object
Dim y As Integer
Dim strDate As String
Dim item As clsPoItem
Set SapGuiAuto = GetObject("SAPGUI")
Set App = SapGuiAuto.GetScriptingEngine
Dim poStr As String
Set Connection = App.Children(0)
Set session = Connection.Children(0)
Dim v() As String

'launch a transaction
session.findById("wnd[0]").Maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ME21n"
session.findById("wnd[0]").sendVKey 0

'If Not IsObject(Application) Then
'   Set SapGuiAuto = GetObject("SAPGUI")
'   Set Application = SapGuiAuto.GetScriptingEngine
'End If
'If Not IsObject(Connection) Then
'   Set Connection = Application.Children(0)
'End If
'If Not IsObject(session) Then
'   Set session = Connection.Children(0)
'End If
'If IsObject(WScript) Then
'   WScript.ConnectObject session, "on"
'   WScript.ConnectObject Application, "on"
'End If
If Day(po.PoDate) >= 10 Then
    strDate = Day(po.PoDate)
Else
    strDate = "0" & Day(po.PoDate)
End If
If Month(po.PoDate) >= 10 Then
    strDate = strDate & "." & Month(po.PoDate)
Else
    strDate = strDate & ".0" & Month(po.PoDate)
End If
strDate = strDate & "." & year(po.PoDate)

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").SetFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").key = "ZSTO"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = po.PurchasingOrg
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = po.PurchasingGroup
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "m024"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").SetFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").caretPosition = 4
session.findById("wnd[0]").sendVKey 0

For Each item In po.getItems
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & y & "]").text = item.index
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6," & y & "]").text = item.amount
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[7," & y & "]").text = po.Unit
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EEIND[8," & y & "]").text = strDate
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11," & y & "]").text = po.Lplant
    y = y + 1
Next item

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11,0]").SetFocus
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[11,0]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
poStr = session.findById("wnd[0]/sbar").text
session.findById("wnd[0]/tbar[0]/btn[3]").press
v = Split(poStr, " ")
If UBound(v) > 0 Then
    orderToPO = v(UBound(v))
End If

End Function

Public Function purOrg(Lplant As String) As Variant
Dim i As Integer
Dim plant As String
Dim sht As Worksheet
Dim bool As Boolean

Set sht = ActiveWorkbook.Sheets("Lokacje")

For i = 2 To 100
    plant = sht.Cells(1, i)
    If Len(plant) = 0 Then
        Exit For
    Else
        If plant = Lplant Then
            purOrg = sht.Cells(2, i)
            bool = True
            Exit For
        End If
    End If
Next i

If bool = False Then
    purOrg = Null
End If

End Function

Public Function purGroup(Lplant As String) As Variant
Dim i As Integer
Dim plant As String
Dim sht As Worksheet
Dim bool As Boolean

Set sht = ActiveWorkbook.Sheets("Lokacje")

For i = 2 To 100
    plant = sht.Cells(1, i)
    If Len(plant) = 0 Then
        Exit For
    Else
        If plant = Lplant Then
            purGroup = sht.Cells(3, i)
            bool = True
            Exit For
        End If
    End If
Next i

If bool = False Then
    purGroup = Null
End If

End Function

Public Function companyCode(Lplant As String) As Variant
Dim i As Integer
Dim plant As String
Dim sht As Worksheet
Dim bool As Boolean

Set sht = ActiveWorkbook.Sheets("Lokacje")

For i = 2 To 100
    plant = sht.Cells(1, i)
    If Len(plant) = 0 Then
        Exit For
    Else
        If plant = Lplant Then
            companyCode = sht.Cells(4, i)
            bool = True
            Exit For
        End If
    End If
Next i

If bool = False Then
    companyCode = Null
End If

End Function




Public Property Get ShipmentIds() As String
    Dim res As String
    Dim ship As clsShipment
    
    res = ""
    
    For Each ship In Items
        res = res & ship.shipmentId & ","
    Next ship
    
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
    
    ShipmentIds = res
End Property

Public Sub downloadZfins(typeStr As String)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newZfin As clsZfin

On Error GoTo err_trap

updateConnection

n = zfins.count
Do While zfins.count > 0
    zfins.Remove n
    n = n - 1
Loop

sSql = "SELECT zfinId, zfinIndex FROM tbZfin WHERE zfinType IN (" & typeStr & ");"
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newZfin = New clsZfin
        With newZfin
            .zfinId = rs.Fields("zfinId").Value
            .zfinIndex = rs.Fields("zfinIndex").Value
            zfins.Add newZfin, CStr(rs.Fields("zfinIndex").Value)
        End With
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
closeConnection
Exit Sub

err_trap:
MsgBox "Error in downloadZfins. Error number: " & Err.Number & ", " & Err.Description
Resume exit_here

End Sub


