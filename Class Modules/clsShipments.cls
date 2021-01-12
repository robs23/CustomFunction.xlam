VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShipments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Items As Collection
Public Stocks As Collection
Public Productions As Collection
Public Removals As Collection
Public ToRecreate As Collection
Public initialDate As Date
Public shipments4removal As String
Public sht As Worksheet

Public Sub Save()
Dim dSql As String
Dim ship As clsShipment
Dim shipIds As String
Dim theWeek As Integer
Dim theYear As Integer
Dim theDate As Date
Dim plants As String
Dim prod As clsProductionItem
Dim stock As clsStockItem

On Error GoTo err_trap

updateConnection

plants = GetLplants

shipIds = ShipmentIds

GetMissingIds 'check if there are any shipments where id is stored in the file, but it's not present in DB

If Len(shipIds) > 0 Then
    'delete PO items
    dSql = "DELETE poi " _
        & "FROM tbPo p LEFT JOIN tbPoItem poi ON p.PoId=poi.PoId " _
        & "WHERE p.shipmentId IN (" & shipIds & ")"
    
    AdoConn.Execute dSql
    
    'delete POs
    dSql = "DELETE p " _
        & "FROM tbPo p " _
        & "WHERE p.shipmentId IN (" & shipIds & ")"
    
    AdoConn.Execute dSql

End If
If Removals.count > 0 Then
    Dim rStr As String
    rStr = RemovalIds
    'delete PO items
    dSql = "DELETE poi " _
        & "FROM tbPo p LEFT JOIN tbPoItem poi ON p.PoId=poi.PoId " _
        & "WHERE p.shipmentId IN (" & rStr & ")"
    
    AdoConn.Execute dSql
    
    'delete POs
    dSql = "DELETE p " _
        & "FROM tbPo p " _
        & "WHERE p.shipmentId IN (" & rStr & ")"
    
    AdoConn.Execute dSql
    
    'delete shipments
    dSql = "DELETE sh " _
        & "FROM tbPlannedShipments sh " _
        & "WHERE sh.PlannedShipmentId IN (" & rStr & ")"
    
    AdoConn.Execute dSql
    
    For Each ship In Removals
        'remove any trace in the file
        sht.Cells(5, ship.Column) = ""
        sht.Cells(7, ship.Column) = ""
        sht.Cells(8, ship.Column) = ""
    Next ship
End If


If Len(plants) > 0 Then
    For x = 0 To 2
        'delete productions and stocks where week, year and plant matches
        theDate = DateAdd("d", x * 7, initialDate)
        theWeek = IsoWeekNumber(theDate)
        theYear = year(theDate)
        
        If Len(plants) > 0 And Productions.count > 0 Then
            'delete productions
            dSql = "DELETE p " _
                    & "FROM tbPlannedProduction p " _
                    & "WHERE p.Week = " & theWeek & " AND p.Year = " & theYear & " AND p.Lplant IN (" & plants & ")"
        
            AdoConn.Execute dSql
        End If
        
        If Len(plants) > 0 And Stocks.count > 0 Then
            'delete stocks
            dSql = "DELETE s " _
                    & "FROM tbPlannedStock s " _
                    & "WHERE s.PlannedDate = '" & theDate & "' AND s.Lplant IN (" & plants & ")"
        
            AdoConn.Execute dSql
        End If
        
    Next x

End If

For Each ship In Items
    If ship.shipmentId > 0 Then
        'check if such a shipment Id extists
        If ToRecreate.count > 0 Then
            If inCollection(CStr(ship.shipmentId), ToRecreate) Then
                ship.shipmentId = 0
            End If
        End If
    End If
    ship.SaveToDb
Next ship

updateConnection

For Each prod In Productions
    prod.Save
Next prod

For Each stock In Stocks
    stock.Save
Next stock

exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in clsShipments.Save. " & Err.Description
Resume exit_here

End Sub

Public Sub GetMissingIds()
'check if all shipments form ShipmentIds exist
'return collection of missing ids or null

On Error GoTo err_trap

Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim ids As String
Dim i As Integer

ids = ShipmentIds

updateConnection

Set ToRecreate = New Collection

v = Split(ids, ",", , vbTextCompare)

For i = LBound(v) To UBound(v)
    ToRecreate.Add Int(v(i)), v(i)
Next i


sSql = "SELECT PlannedShipmentId FROM tbPlannedShipments WHERE PlannedShipmentId IN (" & ids & ");"
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        If inCollection(Int(rs.Fields("PlannedShipmentId")), ToRecreate) Then
            ToRecreate.Remove CStr(rs.Fields("PlannedShipmentId"))
        End If
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in GetMissingIds of clsShipments. Description: " & Err.Description
Resume exit_here

End Sub

Public Function GetLplants() As String
Dim res As String
Dim plant As String
Dim y As Integer

On Error GoTo err_trap

For y = 15 To 93
    plant = ActiveWorkbook.Sheets("Plan").Cells(y, 7)
    If plant <> "0" And plant <> "" Then
        If InStr(1, res, plant, vbTextCompare) = 0 Then
            res = res & "'" & plant & "',"
        End If
    End If
Next y

If Len(res) > 0 Then res = Left(res, Len(res) - 1)

GetLplants = res

exit_here:
Exit Function

err_trap:
MsgBox "Error in GetLplants. Description: " & Err.Description
Resume exit_here

End Function

Public Function Scan() As Boolean
Dim i As Integer
Dim shId As Integer
Dim ship As clsShipment
Dim vDate As Variant
Dim res As Boolean
Dim index As Long
Dim amount As Double
Dim poItem As clsPoItem
Dim isError As Boolean

isError = False

On Error GoTo err_trap

Set Items = New Collection
Set Removals = New Collection

downloadZfins "'zfin'"
initialDate = GetInitialDate

res = True

Set sht = ActiveWorkbook.Sheets("Plan")
added = 0

sTime = Now

For i = 1 To 100
    If isError Then GoTo exit_here
    If sht.Cells(6, 11 + i) = "Shipment" Then
        If Not sht.Cells(12, 11 + i + 1) = "0" Or Not sht.Cells(13, 11 + i + 1) = "0" Then
            If Not sht.Cells(8, 11 + i) = "" Then
                shId = sht.Cells(5, 11 + i)
                Set ship = New clsShipment
                ship.Column = 11 + i
                ship.shipmentId = shId
                ship.DeliveryNotes = GetDN(sht.Cells(14, 11 + i))
                vDate = ConvertDate(sht.Cells(8, 11 + i))
                If IsNull(vDate) Then
                    'finish
                    res = False
                    Exit For
                Else
                    ship.ShipmentDate = vDate
                    Items.Add ship, CStr(11 + i) '<--
                    For y = 15 To 114
                        index = sht.Cells(y, 1)
                        If index = 0 Then
                            Exit For
                        Else
                            amount = sht.Cells(y, 11 + i)
                            If amount > 0 Then
                                Lplant = sht.Range("G" & y)
                                If Lplant = "0" Then
                                    MsgBox "Brak lokacji dla produktu " & index & ". Uzupełnij dane. Nic nie zostało zaimportowane!", vbCritical + vbOKOnly, "Brak danych"
                                    isError = True 'set the flag
                                    Exit For
                                End If
                                'check if we have all data of location
                                Set newPo = ship.getPos(Lplant)
                                If newPo Is Nothing Then
                                    'create new one
                                    Set newPo = New clsPO
                                    With newPo
                                        .Lplant = Lplant
                                    End With
                                    ship.append newPo
                                End If
                                Set poItem = New clsPoItem
                                poItem.initialize index, amount
                                newPo.append poItem
                            End If
                        End If
                    Next y
                End If
            End If
        Else
            'here's an empty one. If it's got shipmentId on it, mark it for removal
            shId = sht.Cells(5, 11 + i)
            If shId > 0 Then
                Set ship = New clsShipment
                ship.Column = 11 + i
                ship.shipmentId = shId
                Removals.Add ship
            End If
        End If
    End If
Next i

If res Then
    res = GetStocks
End If
If res Then
    res = GetProductions
End If

eTime = Now

exit_here:
If isError Then res = False
If res = False Then
    MsgBox "Błąd w clsShipments.Scan. Kolumna " & col2Letter(ship.Column) & ", wierz " & y & ". " & Err.Description
End If
Scan = res
Exit Function

err_trap:
MsgBox "Błąd w clsShipments.Scan. Kolumna " & col2Letter(ship.Column) & ", wierz " & y & ". " & Err.Description
Resume exit_here
End Function

Private Function GetDN(dn As String) As String
Dim start As Integer
Dim res As String

res = dn

If Not Len(dn) = 0 Then
    start = InStr(1, dn, "DN", vbTextCompare)
    If start > 0 Then
        res = Trim(Mid(dn, start + 2, Len(dn) - start))
    End If
End If

GetDN = res

End Function

Public Function GetInitialDate() As Date
Dim wb As Workbook
Dim name As String
Dim v() As String
Dim x() As String
Dim w As Long
Dim y As Long
Dim str As String

Set wb = ActiveWorkbook
name = wb.name

v = Split(name, ".", , vbTextCompare)
If UBound(v) > 0 Then
    str = Right(v(UBound(v) - 1), 7)
    x = Split(str, "_", , vbTextCompare)
    If UBound(x) > 0 Then
        w = x(0)
        y = x(1)
        If w > 0 And y > 0 Then
            GetInitialDate = Week2Date(w, y)
        End If
    End If
    
End If

End Function

Public Function GetStocks() As Boolean
Dim vDate As Variant
Dim sht As Worksheet
Dim res As Boolean
Dim y As Integer
Dim x As Integer
Dim amount As Double
Dim index As Long
Dim stock As clsStockItem
Dim theDate As Date

On Error GoTo err_trap

res = True
Set Stocks = New Collection

Set sht = ActiveWorkbook.Sheets("Plan")

If IsDate(initialDate) Then
    For x = 0 To 2
        theDate = DateAdd("d", x * 7, initialDate)
        For y = 15 To 114
            index = sht.Cells(y, 1)
            If index <> 0 Then
                amount = sht.Cells(y, 8 + (x * 46))
                If amount <> 0 Then
                    Set stock = New clsStockItem
                    stock.PlannedDate = theDate
                    stock.ProductId = zfins(CStr(index)).zfinId
                    stock.amount = amount
                    stock.Lplant = sht.Cells(y, 7)
                    Stocks.Add stock
                End If
            End If
        Next y
    Next x
Else
    MsgBox "Nie udało się ustawić daty początkowej na podstawie nazwy pliku. Upewnij się, że nazwa pliku kończy się formatem ww_yyyy np. 52_2019.xlsx", vbCritical + vbOKOnly, "Zła nazwa pliku"
    res = False
End If

exit_here:
GetStocks = res
Exit Function

err_trap:
res = False
MsgBox "Błąd w clsShipments.GetStocks. Kolumna " & col2Letter(8 + (x * 46)) & ", wiersz " & y & Err.Description
Resume exit_here

End Function

Public Function GetProductions() As Boolean
Dim vDate As Variant
Dim sht As Worksheet
Dim res As Boolean
Dim y As Integer
Dim x As Integer
Dim amount As Double
Dim index As Long
Dim prod As clsProductionItem
Dim theDate As Date
Dim theWeek As Integer
Dim theYear As Integer

On Error GoTo err_trap

res = True
Set Productions = New Collection

Set sht = ActiveWorkbook.Sheets("Plan")

If IsDate(initialDate) Then
    For x = 0 To 2
        theDate = DateAdd("d", x * 7, initialDate)
        theWeek = IsoWeekNumber(theDate)
        theYear = year(theDate)
        For y = 15 To 114
            index = sht.Cells(y, 1)
            If index <> 0 Then
                amount = sht.Cells(y, 11 + (x * 46))
                If amount <> 0 Then
                    Set prod = New clsProductionItem
                    prod.week = theWeek
                    prod.year = theYear
                    prod.amount = amount
                    prod.Lplant = sht.Cells(y, 7)
                    prod.ProductId = zfins(CStr(index)).zfinId
                    Productions.Add prod
                End If
            End If
        Next y
    Next x
Else
    MsgBox "Nie udało się ustawić daty początkowej na podstawie nazwy pliku. Upewnij się, że nazwa pliku kończy się formatem ww_yyyy np. 52_2019.xlsx", vbCritical + vbOKOnly, "Zła nazwa pliku"
    res = False
End If

exit_here:
GetProductions = res
Exit Function

err_trap:
res = False
MsgBox "Błąd w clsShipments.GetProductions. Kolumna " & col2Letter(11 + (x * 46)) & ", wiersz " & y & Err.Description
Resume exit_here

End Function


Public Function ConvertDate(theDate As String) As Variant
Dim d() As String

On Error GoTo err_trap

If IsDate(theDate) Then
    ConvertDate = theDate
Else
    If Len(theDate) = 0 Then
        ConvertDate = Null
    Else
        d = Split(theDate, ".", , vbTextCompare)
        If UBound(d) <> 2 Then
            ConvertDate = Null
        Else
            ConvertDate = DateSerial(CInt(d(2)), CInt(d(1)), CInt(d(0)))
        End If
    End If
End If


If IsNull(ConvertDate) Then
    MsgBox "Nie udało się przekonwertować daty " & theDate & ". Upewnij się że data jest w formacie dd.mm.yyyy np. 31.12.2020"
End If

exit_here:
Exit Function

err_trap:
ConvertDate = Null
MsgBox "Nie udało się przekonwertować daty " & theDate & ". Upewnij się że data jest w formacie dd.mm.yyyy np. 31.12.2020"
Resume exit_here

End Function

Public Property Get ShipmentIds() As String
    Dim res As String
    Dim ship As clsShipment
    
    res = ""
    
    For Each ship In Items
        If ship.shipmentId > 0 Then res = res & ship.shipmentId & ","
    Next ship
    
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
    
    ShipmentIds = res
End Property

Public Property Get RemovalIds() As String
    Dim res As String
    Dim ship As clsShipment
    
    res = ""
    
    For Each ship In Removals
        res = res & ship.shipmentId & ","
    Next ship
    
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
    
    RemovalIds = res
End Property


