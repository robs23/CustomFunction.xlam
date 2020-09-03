Attribute VB_Name = "Import"
Option Explicit
Public AdoConn As ADODB.Connection
Public machs As New Collection
Public zfins As New Collection
Public batches As New Collection
Public Qdocs As New Collection
Public operations As New Collection
Public orders As New Collection
Public locations As New Collection

Public Sub importButton(control As IRibbonControl)
importer
End Sub

Public Sub importer()
Dim x As Integer
Dim y As Integer
Dim str As String
Dim sht As Worksheet
Dim mStr As String
Dim found As Boolean
Dim res As VbMsgBoxResult

Set sht = ActiveWorkbook.ActiveSheet
found = False

For y = 1 To 30
    For x = 1 To 30
        str = sht.Cells(y, x)
        Select Case str
            Case Is = "GRID_PROD_GRAPH"
                res = MsgBox("Próbujesz zaimportować plan produkcyjny do bazy. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then saveProductionPlan
                found = True
                Exit For
            Case Is = "Lista przypisań atrybutów do artykułów"
                res = MsgBox("Próbujesz zaimportować atrybuty ZFINów/ZFORów do bazy. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then updateProperty
                found = True
                Exit For
            Case Is = "Zestawienie obrotów wg artykułów "
                res = MsgBox("Próbujesz zaimportować operacje PW/WZ do bazy. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then exportPW_WZ
                found = True
                Exit For
            Case Is = "Powiązania operacji"
                res = MsgBox("Próbujesz zaimportować powiązania operacji do bazy. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then importConnections
                found = True
                Exit For
            Case Is = "Zestawienie ilości wyprodukowanej w zleceniu"
                res = MsgBox("Próbujesz zaimportować ilości wyprodukowane w zleceniu do bazy. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then importMesQuantities
                found = True
                Exit For
            Case Is = "Zlecenia z partią dosypaną"
                res = MsgBox("Próbujesz zaimportować zlecenia z partią dosypaną. Kontynuujemy?", vbQuestion + vbYesNo, "Potwierdź raport")
                If res = vbYes Then importRework
                found = True
                Exit For
        End Select
    Next x
    If found Then Exit For
Next y

If Not found Then
    mStr = "Nie udało się ustalić typu raportu do zaimportowania, prawdopodobnie nie jest on obsługiwany."
    mStr = mStr & vbNewLine & "Obsługiwane raporty:" & vbNewLine
    mStr = mStr & "-""Grafik produkcyjny"" z MES," & vbNewLine & "-""Lista przypisań atrybutów do artykułów"" z MES," & vbNewLine & "-""Zestawienie obrotów wg artykułów"" z Qguar," & vbNewLine & "-""COOIS"" z SAP R/3 (ręcznie)." & vbNewLine & "-""Zestawienie ilości wyprodukowanej w zleceniu"" z MES" & vbNewLine
    mStr = mStr & "Żadne dane nie zostały dodane do bazy." & vbNewLine & vbNewLine
    mStr = mStr & "Czy chcesz ręcznie wybrać właściwy typ raportu?"
    res = MsgBox(mStr, vbYesNo + vbCritical, "Błędny typ raportu")
    If res = vbYes Then
        impForm.Show
    End If
End If
End Sub

Public Sub importMesQuantities()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim i As Integer
Dim d As Integer
Dim n As Integer
Dim text As String
Dim oCol As Integer
Dim zfinCol As Integer
Dim zfinNameCol As Integer
Dim amountCol As Integer
Dim typeCol As Integer
Dim zfinStr() As String
Dim operStr() As String
Dim counter As Integer
Dim boundery As Long
Dim cSql As String
Dim iSql As String
Dim uSql As String
Dim sSql As String
Dim theType As String
Dim realType As String
Dim s As Long
Dim sapStr As String
Dim error As Boolean
Dim lastRow As Long
Dim o As clsOrder

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If oCol > 0 And zfinCol > 0 And zfinNameCol > 0 And amountCol > 0 And typeCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Nr zlecenia"
                    oCol = n
                    Case Is = "Nr produktu"
                    zfinCol = n
                    Case Is = "Nazwa produktu"
                    zfinNameCol = n
                    Case Is = "Ilość"
                    amountCol = n
                    Case Is = "Nr operacji"
                    typeCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Zestawienie ilości wyprodukowanej w zleceniu"" z MES. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        updateConnection
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, zfinNameCol), Cells(lastRow, zfinNameCol)).Replace "'", "", xlPart, , False
        'continue
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfins & zfors indexes -------------------------------------
        counter = 0
        For s = d To 50000
            If .Cells(s, zfinCol).Value <> "" Then
                theType = ""
                realType = Left(.Cells(s, typeCol).Value, 3)
                If realType = "PRA" Then
                    theType = "zfor"
                ElseIf realType = "MIE" Then
                    theType = "zfor"
                ElseIf realType = "PAK" Then
                    theType = "zfin"
                End If
                If realType = "PRA" Or realType = "MIE" Or realType = "PAK" Then
                    If counter = 0 Then
                        boundery = 0
                        ReDim zfinStr(0) As String
                    ElseIf counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve zfinStr(boundery) As String
                    End If
                    zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'" & sht.Cells(s, zfinNameCol).Value & "','" & theType & "','pr'),"
                    counter = counter + 1
                End If
            ElseIf .Cells(s, zfinCol).Value = "" Then
                For counter = LBound(zfinStr) To UBound(zfinStr)
                    zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255), zfinType nchar(4),prodStatus nchar(2))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor'"
        
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfors'/zfins' order data -------------------------------------
        
        n = orders.count
        Do While orders.count > 0
            orders.Remove n
            n = n - 1
        Loop
        counter = 0
        For s = d To 50000
            If .Cells(s, zfinCol).Value <> "" Then
                realType = Left(.Cells(s, typeCol).Value, 3)
                If realType = "PRA" Then
                    theType = "r"
                ElseIf realType = "MIE" Then
                    theType = "r"
                ElseIf realType = "PAK" Then
                    theType = "p"
                End If
                Set o = newOrder(.Cells(s, oCol).Value, realType, .Cells(s, amountCol).Value)
                With o
                    .zfinId = zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId
                End With
            Else
                Exit For
            End If
        Next s
        For Each o In orders
            If counter = 0 Then
                boundery = 0
                ReDim operStr(0) As String
            ElseIf counter Mod 1000 = 0 Then
                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                boundery = counter / 1000
                ReDim Preserve operStr(boundery) As String
            End If
            theType = ""
            operStr(boundery) = operStr(boundery) & "(" & o.sapId & ",'" & o.theType & "'," & o.zfinId & ","
            If o.theType = "r" Then
                operStr(boundery) = operStr(boundery) & Replace(CDbl(o.roast), ",", ".") & ","
                If o.ground > 0 Then
                    operStr(boundery) = operStr(boundery) & Replace(CDbl(o.ground), ",", ".") & "),"
                Else
                    operStr(boundery) = operStr(boundery) & "NULL),"
                End If
            Else
                operStr(boundery) = operStr(boundery) & Replace(CDbl(o.packed), ",", ".") & ",NULL),"
            End If
            sapStr = sapStr & o.sapId & ","
            counter = counter + 1
        Next o
        sapStr = Left(sapStr, Len(sapStr) - 1)
        For counter = LBound(operStr) To UBound(operStr)
            operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
        Next counter

        'change #orders for TEMPOrders
        cSql = "CREATE TABLE #orders(sapId bigint, type nchar(1),zfinId bigint,executedMes float,executedMesGround float)"
        AdoConn.Execute cSql
        For counter = LBound(operStr) To UBound(operStr)
            iSql = "INSERT INTO #orders(sapId,type,zfinId,executedMes,executedMesGround) VALUES " & operStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId, t1.executedMes = t2.executedMes, t1.executedMesGround = t2.executedMesGround FROM tbOrders t1 INNER JOIN #orders t2 ON t1.sapId = t2.sapId"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT sapId,type,zfinId,GETDATE() as createdOn, executedMes, executedMesGround FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId is not null)"
        iSql = "INSERT INTO tbOrders (sapId,type,zfinId, createdOn, executedMes, executedMesGround) " & sSql
        AdoConn.Execute iSql
        
    End If
End With

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importMesQuantities"". Error number: " & Err.Number & ", " & Err.Description
error = True
Resume Exit_here

End Sub

Private Function newOrder(oNumber As Long, theType As String, qty As Double) As clsOrder
Dim found As Boolean
Dim o As clsOrder

found = False

If orders.count > 0 Then
    For Each o In orders
        If o.sapId = oNumber Then
            found = True
            Exit For
        End If
    Next o
End If

If found = False Then
    Set o = New clsOrder
    o.sapId = oNumber
    orders.Add o, CStr(oNumber)
End If

If theType = "PRA" Then
    o.theType = "r"
    o.roast = o.roast + qty
ElseIf theType = "MIE" Then
    o.theType = "r"
    o.ground = o.ground + qty
ElseIf theType = "PAK" Then
    o.theType = "p"
    o.packed = o.packed + qty
End If

Set newOrder = o

End Function

Public Sub saveProductionPlan()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim i As Integer
Dim dCol As Integer
Dim sCol As Integer
Dim oCol As Integer
Dim tCol As Integer
Dim mCol As Integer
Dim pCol As Integer
Dim aCol As Integer
Dim sesCol As Integer
Dim pnCol As Integer
Dim poCol As Integer
Dim mStrCol As Integer
Dim cSql As String
Dim iSql As String
Dim zfCol As Integer
Dim zfnCol As Integer
Dim d As Integer
Dim n As Integer
Dim zfinStr() As String
Dim zfinZforStr() As String
Dim operStr() As String
Dim orderStr() As String
Dim machStr() As String
Dim operData() As String
Dim operNos As String
Dim orderNos As String
Dim counter As Integer
Dim boundery As Integer
Dim zzBoundery As Integer
Dim text As String
Dim lastRow As Long
Dim error As Boolean
Dim s As Long
Dim sSql As String
Dim uSql As String
Dim rs As ADODB.Recordset
Dim dSql As String
Dim firstDate As Date
Dim lastDate As Date
Dim theType As String
Dim verId As Long
Dim zzCounter As Integer
Dim v() As String

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If dCol > 0 And sCol > 0 And oCol > 0 And tCol > 0 And mCol > 0 And pCol > 0 And aCol > 0 And pnCol > 0 And zfCol > 0 And zfnCol > 0 And mStrCol > 0 And poCol > 0 And sesCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Data zmiany"
                    dCol = n
                    Case Is = "Identyfikator zmiany"
                    sCol = n
                    Case Is = "Identyfikator operacji"
                    oCol = n
                    Case Is = "Typ operacji"
                    tCol = n
                    Case Is = "Nr maszyny"
                    mCol = n
                    Case Is = "Nr produktu"
                    pCol = n
                    Case Is = "Nazwa produktu"
                    pnCol = n
                    Case Is = "Nr blendu"
                    zfCol = n
                    Case Is = "Nazwa blendu"
                    zfnCol = n
                    Case Is = "Il. plan. [j. art.]"
                    aCol = n
                    Case Is = "Nr operacji"
                    mStrCol = n
                    Case Is = "Zlecenie produkcyjne"
                    poCol = n
                    Case Is = "Numer sesji"
                    sesCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Grafik produkcyjny"" z MES. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        'we identified all columns needed. Let's crack on
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, pnCol), Cells(lastRow, pnCol)).Replace "'", "", xlPart
        sht.Range(Cells(d, zfnCol), Cells(lastRow, zfnCol)).Replace "'", "", xlPart
        firstDate = CDate(Application.WorksheetFunction.Min(sht.Range(Cells(d, dCol), Cells(lastRow, dCol))))
        lastDate = CDate(Application.WorksheetFunction.Max(sht.Range(Cells(d, dCol), Cells(lastRow, dCol))))
        verId = newPlanVersion(firstDate, lastDate)
        updateConnection (600)
        'dSql = "DELETE tbOperations FROM tbOperations o LEFT JOIN tbOperationData od ON o.operationId = od.operationId WHERE od.plMoment >= '" & firstDate & "'"
'            dSql = "DELETE FROM tbOperationData WHERE plMoment >= '" & firstDate & "'"
        'adoConn.Execute dSql
        dSql = "DELETE FROM tbOperationData WHERE plMoment >= '" & firstDate & "' AND plMoment <= '" & lastDate & "'"
        AdoConn.Execute dSql
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Pakowanie" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, pCol).Value & ",'" & sht.Cells(s, pnCol).Value & "','zfin','pr'),"
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfCol).Value & ",'" & sht.Cells(s, zfnCol).Value & "','zfor','pr'),"
                counter = counter + 2
            ElseIf .Cells(s, tCol).Value = "" Then
                Exit For
            End If
        Next s
        
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Prażenie" Or sht.Cells(s, tCol).Value = "Mielenie" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, pCol).Value & ",'" & sht.Cells(s, pnCol).Value & "','zfor','pr'),"
                counter = counter + 1
            ElseIf .Cells(s, tCol).Value = "" Then
                For counter = LBound(zfinStr) To UBound(zfinStr)
                    zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor'"
        
'        'let's add zfin-zfor connections
'
'        counter = 0
'
'        For s = d To 50000
'            If sht.Cells(s, tCol).Value = "Pakowanie" Then
'                If counter = 0 Then
'                    boundery = 0
'                    ReDim zfinZforStr(0) As String
'                ElseIf counter Mod 1000 = 0 Then
'                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
'                    boundery = counter / 1000
'                    ReDim Preserve zfinZforStr(zzBoundery) As String
'                End If
'                zfinZforStr(boundery) = zfinZforStr(boundery) & "(" & zfins(CStr(sht.Cells(s, pCol).Value)).zfinId & "," & zfins(CStr(sht.Cells(s, zfCol).Value)).zfinId & "),"
'                counter = counter + 1
'            ElseIf .Cells(s, tCol).Value = "" Then
'                For counter = LBound(zfinZforStr) To UBound(zfinZforStr)
'                    zfinZforStr(counter) = Left(zfinZforStr(counter), Len(zfinZforStr(counter)) - 1)
'                Next counter
'                Exit For
'            End If
'        Next s
'
'        cSql = "CREATE TABLE #zfinZfor(zfinId int, zforId int)"
'        adoConn.Execute cSql
'        For counter = LBound(zfinZforStr) To UBound(zfinZforStr)
'            iSql = "INSERT INTO #zfinZfor(zfinId, zforId) VALUES " & zfinZforStr(counter)
'            adoConn.Execute iSql
'        Next counter
'        uSql = "UPDATE t1 SET t1.zforId = t2.zforId FROM tbZfinZfor t1 INNER JOIN #zfinZfor t2 ON t1.zfinId = t2.zfinId"
'        adoConn.Execute uSql
'        sSql = "SELECT DISTINCT zfinId, zforId FROM #zfinZfor WHERE zfinId NOT IN (SELECT zfinId FROM tbZfinZfor WHERE zfinId IS NOT NULL)"
'        iSql = "INSERT INTO tbZfinZfor (zfinId, zforId) " & sSql
'        adoConn.Execute iSql

        'let's add orders' data

        counter = 0
        
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Pakowanie" Or sht.Cells(s, tCol).Value = "Prażenie" Or sht.Cells(s, tCol).Value = "Mielenie" Then
                If sht.Cells(s, tCol).Value = "Prażenie" Then
                    theType = "r"
                ElseIf sht.Cells(s, tCol).Value = "Pakowanie" Then
                    theType = "p"
                ElseIf sht.Cells(s, tCol).Value = "Mielenie" Then
                    theType = "r"
                End If
                If counter = 0 Then
                    boundery = 0
                    ReDim orderStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve orderStr(boundery) As String
                End If
                orderStr(boundery) = orderStr(boundery) & "(" & sht.Cells(s, poCol).Value & ",'" & theType & "'," & zfins(CStr(sht.Cells(s, pCol).Value)).zfinId & "),"
                orderNos = orderNos & sht.Cells(s, poCol).Value & ","
                counter = counter + 1
            ElseIf .Cells(s, tCol).Value = "" Then
                For counter = LBound(orderStr) To UBound(orderStr)
                    orderNos = Left(orderNos, Len(orderNos) - 1)
                    orderStr(counter) = Left(orderStr(counter), Len(orderStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #orders(sapId bigint, type nchar(1),zfinId bigint)"
            AdoConn.Execute cSql
            For counter = LBound(orderStr) To UBound(orderStr)
                iSql = "INSERT INTO #orders(sapId,type,zfinId) VALUES " & orderStr(counter)
                AdoConn.Execute iSql
            Next counter
            sSql = "SELECT DISTINCT sapId,type,zfinId, GETDATE() as createdOn FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId is not null)"
            iSql = "INSERT INTO tbOrders (sapId,type,zfinId, createdOn) " & sSql
            AdoConn.Execute iSql
        
        downloadOrders sapStr:=orderNos
        
        'let's add operations' data

        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Pakowanie" Or sht.Cells(s, tCol).Value = "Prażenie" Or sht.Cells(s, tCol).Value = "Mielenie" Then
                If sht.Cells(s, tCol).Value = "Prażenie" Then
                    theType = "r"
                ElseIf sht.Cells(s, tCol).Value = "Pakowanie" Then
                    theType = "p"
                ElseIf sht.Cells(s, tCol).Value = "Mielenie" Then
                    theType = "g"
                End If
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                v = Split(sht.Cells(s, mStrCol).Value, "#")
                operStr(boundery) = operStr(boundery) & "(" & sht.Cells(s, oCol).Value & ",'" & theType & "'," & zfins(CStr(sht.Cells(s, pCol).Value)).zfinId & ",'" & v(0) & "'," & CInt(sht.Cells(s, sesCol).Value) & "," & orders(CStr(sht.Cells(s, poCol).Value)).orderId & "),"
                operNos = operNos & sht.Cells(s, oCol).Value & ","
                counter = counter + 1
            ElseIf .Cells(s, tCol).Value = "" Then
                For counter = LBound(operStr) To UBound(operStr)
                    operNos = Left(operNos, Len(operNos) - 1)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #operations(mesId int, type nchar(1),zfinId bigint, mesString nvarchar(50), SessionNumber int, orderId int)"
            AdoConn.Execute cSql
            For counter = LBound(operStr) To UBound(operStr)
                iSql = "INSERT INTO #operations(mesId,type,zfinId,mesString, SessionNumber, orderId) VALUES " & operStr(counter)
                AdoConn.Execute iSql
            Next counter
            uSql = "UPDATE t1 SET t1.type = t2.type, t1.zfinId = t2.zfinId, t1.mesString=t2.mesString, t1.SessionNumber = t2.SessionNumber, t1.orderId = t2.orderId FROM tbOperations t1 INNER JOIN #operations t2 ON t1.mesId = t2.mesId"
            AdoConn.Execute uSql
            sSql = "SELECT DISTINCT mesId,type,zfinId,mesString, SessionNumber, orderId, GETDATE() as createdOn FROM #operations WHERE mesId NOT IN (SELECT mesId FROM tbOperations WHERE mesId is not null)"
            iSql = "INSERT INTO tbOperations (mesId,type,zfinId,mesString, SessionNumber, orderId, createdOn) " & sSql
            AdoConn.Execute iSql
        
        downloadOperations mesStr:=operNos
        
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Pakowanie" Or sht.Cells(s, tCol).Value = "Prażenie" Then
                If sht.Cells(s, tCol).Value = "Prażenie" Then
                    theType = "r"
                ElseIf sht.Cells(s, tCol).Value = "Pakowanie" Then
                    theType = "p"
                ElseIf sht.Cells(s, tCol).Value = "Mielenie" Then
                    theType = "g"
                End If
                If counter = 0 Then
                    boundery = 0
                    ReDim machStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve machStr(boundery) As String
                End If
                machStr(boundery) = machStr(boundery) & "(" & CInt(Right(sht.Cells(s, mCol).Value, 2)) & ",'" & theType & "','" & sht.Cells(s, mCol).Value & "'),"
                counter = counter + 1
            ElseIf .Cells(s, tCol).Value = "" Then
                For counter = LBound(operStr) To UBound(operStr)
                    machStr(counter) = Left(machStr(counter), Len(machStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #machines(machineNumber int, machineType nchar(1),machineName nchar(20))"
            AdoConn.Execute cSql
            For counter = LBound(machStr) To UBound(machStr)
                iSql = "INSERT INTO #machines(machineNumber,machineType,machineName) VALUES " & machStr(counter)
                AdoConn.Execute iSql
            Next counter
            sSql = "SELECT DISTINCT machineNumber,machineType,machineName, GETDATE() as createdOn FROM #machines WHERE machineName NOT IN (SELECT machineName FROM tbMachine WHERE machineName is not null)"
            iSql = "INSERT INTO tbMachine (machineNumber,machineType,machineName, createdOn) " & sSql
            AdoConn.Execute iSql
        
        downloadMachines
        'printMachs
        
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "Pakowanie" Or sht.Cells(s, tCol).Value = "Prażenie" Or sht.Cells(s, tCol).Value = "Mielenie" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operData(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operData(boundery) As String
                End If
                operData(boundery) = operData(boundery) & "(" & operations(CStr(sht.Cells(s, oCol).Value)).operationId & "," & machs(CStr(sht.Cells(s, mCol).Value)).machId & "," & sht.Cells(s, aCol).Value & ",'" & sht.Cells(s, dCol).Value & "'," & sht.Cells(s, sCol).Value & "),"
                counter = counter + 1
            ElseIf .Cells(s, tCol).Value = "" Then
                For counter = LBound(operData) To UBound(operData)
                    operData(counter) = Left(operData(counter), Len(operData(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
        cSql = "CREATE TABLE #operData(operationId int,plMach int, plAmount float,plMoment datetime,plShift int)"
        AdoConn.Execute cSql
        For counter = LBound(operData) To UBound(operData)
            iSql = "INSERT INTO #operData(operationId,plMach,plAmount,plMoment,plShift) VALUES " & operData(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT operationId,plMach,plAmount,plMoment,plShift FROM #operData"
        iSql = "INSERT INTO tbOperationData (operationId,plMach,plAmount,plMoment,plShift) " & sSql
        AdoConn.Execute iSql
        sSql = "SELECT DISTINCT " & verId & " AS operDataVer,operationId,plMach,plAmount,plMoment,plShift FROM #operData"
        iSql = "INSERT INTO tbOperationDataHistory (operDataVer,operationId,plMach,plAmount,plMoment,plShift) " & sSql
        AdoConn.Execute iSql
        
    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
closeConnection
Exit Sub

err_trap:
error = True
MsgBox "Error in saveProductionPlan. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here


End Sub

Public Function newPlanVersion(dFrom As Date, dTo As Date) As Long
Dim iSql As String
Dim rs As ADODB.Recordset

On Error GoTo err_trap

updateConnection

If dFrom >= dTo Then
    MsgBox "Brak operacji w zakresie od " & dFrom & " do " & dTo, vbCritical + vbOKOnly, "Pusty zakres"
Else
'    rStart = DateAdd("h", -10, Week2Date(DatePart("ww", dFrom, vbSunday, vbFirstFourDays), Year(dFrom), vbSunday, vbFirstFourDays))
'    pStart = DateAdd("h", 22, Week2Date(DatePart("ww", dFrom, vbSunday, vbFirstFourDays), Year(dFrom), vbSunday, vbFirstFourDays))
    
    iSql = "INSERT INTO tbOperationDataVersions (createdOn,startRange,endRange) VALUES ('" & Now & "','" & dFrom & "','" & dTo & "');"
    AdoConn.Execute iSql
    Set rs = AdoConn.Execute("SELECT @@Identity", , adCmdText)
    newPlanVersion = rs.Fields(0)
End If

Exit_here:
closeConnection
Set rs = Nothing
Exit Function

err_trap:
MsgBox "Error in newPlanVersion. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Function

Public Sub importBOM()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim rs As ADODB.Recordset
Dim i As Integer
Dim mCol As Integer
Dim zfinCol As Integer
Dim aCol As Integer
Dim cSql As String
Dim iSql As String
Dim zfCol As Integer
Dim uCol As Integer
Dim nCol As Integer
Dim zfnCol As Integer
Dim d As Integer
Dim n As Integer
Dim zfinStr() As String
Dim bomStr() As String
Dim zfinZforStr() As String
Dim operStr() As String
Dim machStr() As String
Dim operData() As String
Dim operNos As String
Dim counter As Integer
Dim boundery As Integer
Dim zzBoundery As Integer
Dim text As String
Dim lastRow As Long
Dim error As Boolean
Dim s As Long
Dim sSql As String
Dim uSql As String
Dim dSql As String
Dim firstDate As Date
Dim lastDate As Date
Dim theType As String
Dim materialType As String
Dim verId As Long
Dim zzCounter As Integer
Dim v() As String
Dim sessionId As Integer


On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If zfinCol > 0 And aCol > 0 And uCol > 0 And mCol > 0 And nCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Material"
                    zfinCol = n
                    Case Is = "Quantity"
                    aCol = n
                    Case Is = "Un"
                    uCol = n
                    Case Is = "Component"
                    mCol = n
                    Case Is = "BOM component"
                    nCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Green coffee - BOM overview LR (BOMLR)"" z R/3 SQ01. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        'we identified all columns needed. Let's crack on
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, nCol), Cells(lastRow, nCol)).Replace "'", "", xlPart
        updateConnection (600)
        counter = 0
        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                Exit For
            Else
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                theType = ""
                If Left(CStr(sht.Cells(s, mCol).Value), 4) = "4960" Then
                    theType = "zfor"
                Else
                    theType = "zfin"
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",NULL,'" & theType & "','pr', NULL),"
                theType = ""
                materialType = "NULL"
                If Left(CStr(sht.Cells(s, mCol).Value), 4) = "4960" And InStr(1, CStr(sht.Cells(s, uCol).Value), "KG", vbTextCompare) > 0 Then
                    theType = "zcom"
                Else
                    If InStr(1, CStr(sht.Cells(s, uCol).Value), "KG", vbTextCompare) > 0 And sht.Cells(s, aCol).Value > 100 Then
                        theType = "zfor"
                    Else
                        theType = "zpkg"
                        If InStr(1, CStr(sht.Cells(s, nCol).Value), " WRO ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " FOIL ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "WRO ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "FOIL ", vbTextCompare) = 1 Then
                            materialType = "2"
                        ElseIf (InStr(1, CStr(sht.Cells(s, nCol).Value), " BX ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " BOX ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "BX ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "BOX ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " LD ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "LD ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " TR ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "TR ", vbTextCompare) = 1) And Left(CStr(sht.Cells(s, nCol).Value), 2) <> "ST" Then
                            materialType = "4"
                        End If
                    End If
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, mCol).Value & ",'" & sht.Cells(s, nCol).Value & "','" & theType & "','pr'," & materialType & "),"
                counter = counter + 2
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2), materialType int)"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            If zfinStr(counter) <> "" Then zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus,materialType) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex WHERE t2.zfinName IS NOT NULL"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,materialType,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,materialType,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor','zpkg','zcom'"

        'let's add bom reconciliation information
        
        Set rs = AdoConn.Execute("INSERT INTO tbBomReconciliation (dateAdded, createdBy) VALUES ('" & Now & "', 43);SELECT SCOPE_IDENTITY()")
        Set rs = rs.NextRecordset
        sessionId = rs.Fields(0)
        rs.Close
        Set rs = Nothing
            
        counter = 0

        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                For counter = LBound(bomStr) To UBound(bomStr)
                    bomStr(counter) = Left(bomStr(counter), Len(bomStr(counter)) - 1)
                Next counter
                Exit For
            Else
                If counter = 0 Then
                    boundery = 0
                    ReDim bomStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve bomStr(boundery) As String
                End If
                bomStr(boundery) = bomStr(boundery) & "(" & zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId & "," & sessionId & "," & zfins(CStr(sht.Cells(s, mCol).Value)).zfinId & "," & Replace(sht.Cells(s, aCol).Value, ",", ".") & ",'" & sht.Cells(s, uCol).Value & "'),"
                counter = counter + 1
            End If
        Next s

        cSql = "CREATE TABLE #boms(zfinId int, bomRecId int, materialId int, amount float, unit varchar(5))"
        AdoConn.Execute cSql
        For counter = LBound(bomStr) To UBound(bomStr)
            iSql = "INSERT INTO #boms(zfinId, bomRecId, materialId, amount, unit) VALUES " & bomStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT zfinId, bomRecId, materialId, amount, unit FROM #boms "
        iSql = "INSERT INTO tbBom (zfinId, bomRecId, materialId, amount, unit) " & sSql
        AdoConn.Execute iSql
        
    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
closeConnection
Exit Sub

err_trap:
error = True
MsgBox "Error in importBOM. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here


End Sub

Public Sub importCosting()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim rs As ADODB.Recordset
Dim i As Integer
Dim mCol As Integer
Dim zfinCol As Integer
Dim aCol As Integer
Dim cSql As String
Dim iSql As String
Dim zfCol As Integer
Dim uCol As Integer
Dim nCol As Integer
Dim clsCol As Integer 'costing lot size
Dim zfnCol As Integer
Dim tCol As Integer 'type column
Dim vCol As Integer 'verification col
Dim zfinNCol As Integer
Dim spCol As Integer
Dim d As Integer
Dim n As Integer
Dim zfinStr() As String
Dim bomStr() As String
Dim zfinZforStr() As String
Dim operStr() As String
Dim machStr() As String
Dim operData() As String
Dim operNos As String
Dim counter As Integer
Dim boundery As Integer
Dim zzBoundery As Integer
Dim text As String
Dim lastRow As Long
Dim error As Boolean
Dim s As Long
Dim sSql As String
Dim uSql As String
Dim dSql As String
Dim firstDate As Date
Dim lastDate As Date
Dim theType As String
Dim materialType As String
Dim verId As Long
Dim zzCounter As Integer
Dim v() As String
Dim sessionId As Integer


On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If zfinCol > 0 And zfinNCol > 0 And spCol > 0 And clsCol > 0 And uCol > 0 And vCol > 0 And tCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Material"
                    zfinCol = n
                    Case Is = "Material description"
                    zfinNCol = n
                    Case Is = "Standard price"
                    spCol = n
                    Case Is = "Costing lot size"
                    clsCol = n
                    Case Is = "BUn"
                    uCol = n
                    Case Is = "Ovrhd grp"
                    vCol = n
                    Case Is = "MTyp"
                    tCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Costing data (ZCOMM_HU)"" z R/3 SQ01 (group MMI). Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        'we identified all columns needed. Let's crack on
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, zfinNCol), Cells(lastRow, zfinNCol)).Replace "'", "", xlPart
        updateConnection (600)
        counter = 0
        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, vCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'" & sht.Cells(s, zfinNCol).Value & "','" & sht.Cells(s, tCol).Value & "'),"
                counter = counter + 1
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            If zfinStr(counter) <> "" Then zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName, t1.zfinType = t2.zfinType FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex WHERE t2.zfinName IS NOT NULL"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor','zpkg','zcom'"

        'let's add bom reconciliation information
        
        Set rs = AdoConn.Execute("INSERT INTO tbCostingReconciliation (dateAdded, addedBy) VALUES ('" & Now & "', 43);SELECT SCOPE_IDENTITY()")
        Set rs = rs.NextRecordset
        sessionId = rs.Fields(0)
        rs.Close
        Set rs = Nothing
            
        counter = 0

        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                For counter = LBound(bomStr) To UBound(bomStr)
                    bomStr(counter) = Left(bomStr(counter), Len(bomStr(counter)) - 1)
                Next counter
                Exit For
            ElseIf Len(sht.Cells(s, vCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim bomStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve bomStr(boundery) As String
                End If
                bomStr(boundery) = bomStr(boundery) & "(" & zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId & "," & sessionId & "," & Replace(sht.Cells(s, spCol).Value, ",", ".") & "," & Replace(sht.Cells(s, clsCol).Value, ",", ".") & ",'" & sht.Cells(s, uCol).Value & "'),"
                counter = counter + 1
            End If
        Next s

        cSql = "CREATE TABLE #costs(zfinId int, reconciliationId int, cost float, CostLotSize float, CostLotSizeUnit varchar(10))"
        AdoConn.Execute cSql
        For counter = LBound(bomStr) To UBound(bomStr)
            iSql = "INSERT INTO #costs(zfinId, reconciliationId, cost, CostLotSize, CostLotSizeUnit) VALUES " & bomStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT zfinId, reconciliationId, cost, CostLotSize, CostLotSizeUnit FROM #costs "
        iSql = "INSERT INTO tbCosting (zfinId, reconciliationId, cost, CostLotSize, CostLotSizeUnit) " & sSql
        AdoConn.Execute iSql

        
    End If
End With
    
Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
closeConnection
Exit Sub

err_trap:
error = True
MsgBox "Error in importCosting. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here


End Sub

Public Sub importRework()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim i As Integer
Dim d As Integer
Dim n As Integer
Dim text As String
Dim dCol As Integer
Dim fCol As Integer
Dim bCol As Integer
Dim tCol As Integer
Dim aCol As Integer
Dim oStr() As String
Dim rStr() As String
Dim ordersStr As String
Dim operStr() As String
Dim counter As Integer
Dim boundery As Long
Dim cSql As String
Dim iSql As String
Dim uSql As String
Dim sSql As String
Dim theType As String
Dim realType As String
Dim s As Long
Dim sapStr As String
Dim error As Boolean
Dim lastRow As Long
Dim o As clsOrder

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If dCol > 0 And tCol > 0 And fCol > 0 And aCol > 0 And bCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Data dosypu"
                    dCol = n
                    Case Is = "Nr zlecenia"
                    tCol = n
                    Case Is = "Zlecenie dosypane"
                    fCol = n
                    Case Is = "Il. dosypana"
                    aCol = n
                    Case Is = "Id. partii"
                    bCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Zlecenia z partią dosypaną""", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        updateConnection
        'continue
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing orders -------------------------------------
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, aCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim oStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve oStr(boundery) As String
                End If
                oStr(boundery) = oStr(boundery) & "(" & sht.Cells(s, tCol).Value & "),"
                oStr(boundery) = oStr(boundery) & "(" & sht.Cells(s, fCol).Value & "),"
                ordersStr = ordersStr + sht.Cells(s, tCol).Value + "," + sht.Cells(s, fCol).Value + ","
                counter = counter + 2
            End If
        Next s
            
        cSql = "CREATE TABLE #orders(sapId int)"
        AdoConn.Execute cSql
        For counter = LBound(oStr) To UBound(oStr)
            If oStr(counter) <> "" Then oStr(counter) = Left(oStr(counter), Len(oStr(counter)) - 1)
            If ordersStr <> "" Then ordersStr = Left(ordersStr, Len(ordersStr) - 1)
            iSql = "INSERT INTO #orders(sapId) VALUES " & oStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT sapId, GETDATE() as createdOn FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId IS NOT NULL)"
        iSql = "INSERT INTO tbOrders (sapId,createdOn) " & sSql
        AdoConn.Execute iSql
        
        
        downloadOrders , ordersStr
        
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add rework data -------------------------------------
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, aCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim rStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve rStr(boundery) As String
                End If
                rStr(boundery) = rStr(boundery) & "('" & sht.Cells(s, dCol).Value & "'," & orders(CStr(sht.Cells(s, fCol))).orderId & "," & orders(CStr(sht.Cells(s, tCol))).orderId & "," & Replace(sht.Cells(s, aCol), ",", ".") & "," & sht.Cells(s, bCol) & "),"
                counter = counter + 1
            End If
        Next s
            
        cSql = "CREATE TABLE #rework(RDate datetime, RFrom int, RTo int, RAmount int, RBatch bigint)"
        AdoConn.Execute cSql
        For counter = LBound(rStr) To UBound(rStr)
            If rStr(counter) <> "" Then rStr(counter) = Left(rStr(counter), Len(rStr(counter)) - 1)
            iSql = "INSERT INTO #rework(RDate, RFrom, RTo, RAmount,RBatch) VALUES " & rStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT RDate, RFrom, RTo, RAmount,RBatch, GETDATE() as creationDate, 43 as createdBy FROM #rework WHERE RBatch NOT IN (SELECT RBatch FROM tbRework WHERE RBatch IS NOT NULL)"
        iSql = "INSERT INTO tbRework (RDate, RFrom, RTo, RAmount, RBatch, creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
    End If
End With

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importRework"". Error number: " & Err.Number & ", " & Err.Description
error = True
Resume Exit_here
End Sub

Public Sub importMb51()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim i As Integer
Dim d As Integer
Dim n As Integer
Dim text As String
Dim dCol As Integer
Dim fCol As Integer
Dim bCol As Integer
Dim tCol As Integer
Dim aCol As Integer
Dim oStr() As String
Dim bStr() As String
Dim rStr() As String
Dim ordersStr As String
Dim batchStr As String
Dim operStr() As String
Dim counter As Integer
Dim boundery As Long
Dim cSql As String
Dim iSql As String
Dim uSql As String
Dim sSql As String
Dim theType As String
Dim realType As String
Dim s As Long
Dim sapStr As String
Dim error As Boolean
Dim combId As String
Dim lastRow As Long
Dim o As clsOrder

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If dCol > 0 And tCol > 0 And fCol > 0 And aCol > 0 And bCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Posting Date"
                    dCol = n
                    Case Is = "Order"
                    tCol = n
                    Case Is = "Batch"
                    fCol = n
                    Case Is = "Quantity"
                    aCol = n
                    Case Is = "Material Document"
                    bCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to MB51 w układzie M024_ROB", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        updateConnection
        'continue
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing orders -------------------------------------
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, aCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim oStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve oStr(boundery) As String
                End If
                oStr(boundery) = oStr(boundery) & "(" & sht.Cells(s, tCol).Value & "),"
                ordersStr = ordersStr + sht.Cells(s, tCol).Value + ","
                counter = counter + 1
            End If
        Next s
            
        cSql = "CREATE TABLE #orders(sapId int)"
        AdoConn.Execute cSql
        For counter = LBound(oStr) To UBound(oStr)
            If oStr(counter) <> "" Then oStr(counter) = Left(oStr(counter), Len(oStr(counter)) - 1)
            If ordersStr <> "" Then ordersStr = Left(ordersStr, Len(ordersStr) - 1)
            iSql = "INSERT INTO #orders(sapId) VALUES " & oStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT sapId, GETDATE() as createdOn FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId IS NOT NULL)"
        iSql = "INSERT INTO tbOrders (sapId,createdOn) " & sSql
        AdoConn.Execute iSql
        
        
        downloadOrders , ordersStr
        
        '--------------- Let's add missing batches -------------------------------------
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, aCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim bStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve bStr(boundery) As String
                End If
                bStr(boundery) = bStr(boundery) & "(" & DropLSU(sht.Cells(s, fCol).Value) & "),"
                batchStr = batchStr + CStr(DropLSU(sht.Cells(s, fCol).Value)) + ","
                counter = counter + 1
            End If
        Next s
            
        cSql = "CREATE TABLE #batches(batchNumber bigint)"
        AdoConn.Execute cSql
        For counter = LBound(bStr) To UBound(bStr)
            If bStr(counter) <> "" Then bStr(counter) = Left(bStr(counter), Len(bStr(counter)) - 1)
            If batchStr <> "" Then batchStr = Left(batchStr, Len(batchStr) - 1)
            iSql = "INSERT INTO #batches(batchNumber) VALUES " & bStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT batchNumber, GETDATE() as createdOn FROM #batches WHERE batchNumber NOT IN (SELECT batchNumber FROM tbBatch WHERE batchNumber IS NOT NULL)"
        iSql = "INSERT INTO tbBatch (batchNumber,createdOn) " & sSql
        AdoConn.Execute iSql
        
        
        downloadBatches batchStr
        
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add rework data -------------------------------------
        counter = 0
        For s = d To 50000
            If sht.Cells(s, tCol).Value = "" Then
                Exit For
            ElseIf Len(sht.Cells(s, aCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim rStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve rStr(boundery) As String
                End If
                combId = batches(CStr(DropLSU(sht.Cells(s, fCol)))).bId & "_" & orders(CStr(sht.Cells(s, tCol))).orderId & "_" & sht.Cells(s, bCol)
                rStr(boundery) = rStr(boundery) & "('" & sht.Cells(s, dCol).Value & "'," & batches(CStr(DropLSU(sht.Cells(s, fCol)))).bId & "," & orders(CStr(sht.Cells(s, tCol))).orderId & "," & -1 * Replace(sht.Cells(s, aCol), ",", ".") & ",'" & combId & "'),"
                counter = counter + 1
            End If
        Next s
            
        cSql = "CREATE TABLE #rework(RDate datetime, RFrom bigint, RTo int, RAmount int, RBatch nvarchar(200))"
        AdoConn.Execute cSql
        For counter = LBound(rStr) To UBound(rStr)
            If rStr(counter) <> "" Then rStr(counter) = Left(rStr(counter), Len(rStr(counter)) - 1)
            iSql = "INSERT INTO #rework(RDate, RFrom, RTo, RAmount,RBatch) VALUES " & rStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT RDate, RFrom, RTo, RAmount,RBatch, GETDATE() as creationDate, 43 as createdBy FROM #rework WHERE RBatch NOT IN (SELECT RBatch FROM tbReworkWarehouse WHERE RBatch IS NOT NULL)"
        iSql = "INSERT INTO tbReworkWarehouse (RDate, RFrom, RTo, RAmount, RBatch, creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
    End If
End With

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importRework"". Error number: " & Err.Number & ", " & Err.Description
error = True
Resume Exit_here
End Sub


Public Sub importComponentUsage()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim rs As ADODB.Recordset
Dim i As Integer
Dim mCol As Integer
Dim zfinCol As Integer
Dim aCol As Integer
Dim cSql As String
Dim iSql As String
Dim zfCol As Integer
Dim zfinNameCol As Integer
Dim matNameCol As Integer
Dim uCol As Integer
Dim nCol As Integer
Dim matCol As Integer
Dim oCol As Integer
Dim actConsCol As Integer
Dim actScrapCol As Integer
Dim bomScrapCol As Integer
Dim tarScrapCol As Integer
Dim tarConsCol As Integer
Dim oAmountCol As Integer
Dim mTypeCol As Integer
Dim zfnCol As Integer
Dim d As Integer
Dim n As Integer
Dim zfinStr() As String
Dim bomStr() As String
Dim zfinZforStr() As String
Dim operStr() As String
Dim prodCons() As String
Dim machStr() As String
Dim operData() As String
Dim operNos As String
Dim counter As Integer
Dim boundery As Integer
Dim zzBoundery As Integer
Dim text As String
Dim lastRow As Long
Dim error As Boolean
Dim s As Long
Dim sapStr As String
Dim sSql As String
Dim uSql As String
Dim dSql As String
Dim firstDate As Date
Dim lastDate As Date
Dim theType As String
Dim materialType As String
Dim verId As Long
Dim zzCounter As Integer
Dim v() As String
Dim sessionId As Integer

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If zfinCol > 0 And zfinNameCol > 0 And matCol > 0 And matNameCol > 0 And oCol > 0 And actConsCol > 0 And actScrapCol > 0 And bomScrapCol > 0 And oAmountCol > 0 And tarConsCol > 0 And tarScrapCol > 0 And mTypeCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Material number"
                    zfinCol = n
                    Case Is = "Material description"
                    zfinNameCol = n
                    Case Is = "Material Type"
                    mTypeCol = n
                    Case Is = "Component"
                    matCol = n
                    Case Is = "Component material description"
                    matNameCol = n
                    Case Is = "Production order"
                    oCol = n
                    Case Is = "Actual Consumption Qty"
                    actConsCol = n
                    Case Is = "Actual scrap % - target based"
                    actScrapCol = n
                    Case Is = "Expected scrap %"
                    bomScrapCol = n
                    Case Is = "Planned qty"
                    oAmountCol = n
                    Case Is = "Target  Qty"
                    tarConsCol = n
                    Case Is = "Target scrap %"
                    tarScrapCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Production Order Variance Report (ZZZV_VARIANCEREPORT)"" z R/3. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        Unload impForm
        wait.Show
        'we identified all columns needed. Let's crack on
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, zfinNameCol), Cells(lastRow, zfinNameCol)).Replace "'", "", xlPart
        sht.Range(Cells(d, matNameCol), Cells(lastRow, matNameCol)).Replace "'", "", xlPart
        updateConnection (600)
        counter = 0
        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                Exit For
            Else
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                theType = ""
                If InStr(1, sht.Cells(s, mTypeCol).Value, "ZFIN", vbTextCompare) > 0 Then
                    theType = "zfin"
                Else
                    theType = "zfor"
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'" & sht.Cells(s, zfinNameCol).Value & "','" & theType & "','pr'),"
                
                theType = ""
                If InStr(1, sht.Cells(s, mTypeCol).Value, "ZFOR", vbTextCompare) > 0 Then
                    theType = "zcom"
                Else
                    If Left(CStr(sht.Cells(s, matCol).Value), 1) = "1" Then
                        theType = "zpkg"
                    Else
                        theType = "zfor"
                    End If
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, matCol).Value & ",'" & sht.Cells(s, matNameCol).Value & "','" & theType & "','pr'),"
                
                counter = counter + 2
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            If zfinStr(counter) <> "" Then zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex WHERE t2.zfinName IS NOT NULL"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor','zpkg','zcom'"
'        For s = d To 50000
'            If sht.Cells(s, zfinCol).Value = "" Then
'                Exit For
'            Else
'                If counter = 0 Then
'                    boundery = 0
'                    ReDim zfinStr(0) As String
'                ElseIf counter Mod 1000 = 0 Then
'                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
'                    boundery = counter / 1000
'                    ReDim Preserve zfinStr(boundery) As String
'                End If
'                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'" & sht.Cells(s, zfinNameCol).Value & "','pr'),"
'                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, matCol).Value & ",'" & sht.Cells(s, matNameCol).Value & "','pr'),"
'                counter = counter + 2
'            End If
'        Next s
'
'        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255), prodStatus nchar(2))"
'        adoConn.Execute cSql
'        For counter = LBound(zfinStr) To UBound(zfinStr)
'            If zfinStr(counter) <> "" Then zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
'            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,prodStatus) VALUES " & zfinStr(counter)
'            adoConn.Execute iSql
'        Next counter
'        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex WHERE t2.zfinName IS NOT NULL"
'        adoConn.Execute uSql
'        sSql = "SELECT DISTINCT zfinIndex,zfinName,prodStatus,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
'        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,prodStatus,creationDate,createdBy) " & sSql
'        adoConn.Execute iSql
''
'        downloadZfins "'zfin','zfor','zpkg','zcom'"

        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing operations -------------------------------------
        counter = 0
        For s = d To 50000
            If Len(sht.Cells(s, oCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                sapStr = sapStr & sht.Cells(s, oCol).Value & ","
                operStr(boundery) = operStr(boundery) & "(" & sht.Cells(s, oCol).Value & "," & zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId & "," & Replace(CDbl(sht.Cells(s, oAmountCol).Value), ",", ".") & "),"
                counter = counter + 1
            Else
                sapStr = Left(sapStr, Len(sapStr) - 1)
                For counter = LBound(operStr) To UBound(operStr)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
        
        cSql = "CREATE TABLE #orders(sapId bigint, zfinId bigint, plannedSap float)"
        AdoConn.Execute cSql
        For counter = LBound(operStr) To UBound(operStr)
            iSql = "INSERT INTO #orders(sapId,zfinId,plannedSap) VALUES " & operStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId, t1.plannedSap = t2.plannedSap FROM tbOrders t1 INNER JOIN #orders t2 ON t1.sapId = t2.sapId"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT sapId,zfinId,plannedSap,GETDATE() as createdOn FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId IS NOT NULL)"
        iSql = "INSERT INTO tbOrders (sapId,zfinId,plannedSap,createdOn) " & sSql
        AdoConn.Execute iSql
    
        downloadOrders sapStr:=sapStr
    
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add componets usage details -------------------------------------
        counter = 0
        For s = d To 50000
            If Len(sht.Cells(s, oCol).Value) > 0 Then
                If counter = 0 Then
                    boundery = 0
                    ReDim prodCons(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve prodCons(boundery) As String
                End If
                prodCons(boundery) = prodCons(boundery) & "(" & orders(CStr(sht.Cells(s, oCol).Value)).orderId & "," & zfins(CStr(sht.Cells(s, matCol).Value)).zfinId & "," & Replace(CDbl(sht.Cells(s, actConsCol).Value), ",", ".") & "," & Replace(CDbl(sht.Cells(s, bomScrapCol).Value), ",", ".") & "," & Replace(CDbl(sht.Cells(s, actScrapCol).Value), ",", ".") & "," & Replace(CDbl(sht.Cells(s, tarConsCol).Value), ",", ".") & "," & Replace(CDbl(sht.Cells(s, tarScrapCol).Value), ",", ".") & "),"
                counter = counter + 1
            Else
                For counter = LBound(prodCons) To UBound(prodCons)
                    prodCons(counter) = Left(prodCons(counter), Len(prodCons(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
        
        cSql = "CREATE TABLE #prodCons(orderId int,componentId int,actualConsumption float,bomScrap float,actualScrap float,targetConsumption float,targetScrap float)"
        AdoConn.Execute cSql
        For counter = LBound(prodCons) To UBound(prodCons)
            iSql = "INSERT INTO #prodCons (orderId,componentId,actualConsumption,bomScrap,actualScrap,targetConsumption,targetScrap) VALUES " & prodCons(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.actualConsumption = t2.actualConsumption, t1.bomScrap = t2.bomScrap, t1.actualScrap = t2.actualScrap, t1.targetConsumption = t2.targetConsumption,t1.targetScrap = t2.targetScrap  FROM tbProductionConsumption t1 INNER JOIN #prodCons t2 ON t1.orderId=t2.orderId AND t1.componentId=t2.componentId"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT orderId,componentId,actualConsumption,bomScrap,actualScrap,targetConsumption,targetScrap,GETDATE() as createdOn FROM #prodCons t1 WHERE NOT EXISTS (SELECT * FROM tbProductionConsumption t2 WHERE t1.orderId=t2.orderId AND t1.componentId=t2.componentId)"
        iSql = "INSERT INTO tbProductionConsumption (orderId,componentId,actualConsumption,bomScrap,actualScrap,targetConsumption,targetScrap,createdOn) " & sSql
        AdoConn.Execute iSql
       
    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
eTime = Now
Unload wait
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
closeConnection
Exit Sub

err_trap:
error = True
MsgBox "Error in ""importComponentUsage"". Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here


End Sub


Public Sub importComponentScrap()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim rs As ADODB.Recordset
Dim i As Integer
Dim tCol As Integer
Dim zfinCol As Integer
Dim nCol As Integer
Dim cSql As String
Dim iSql As String
Dim zfCol As Integer
Dim uCol As Integer
Dim statCol As Integer
Dim zfnCol As Integer
Dim minCol As Integer
Dim maxCol As Integer
Dim rvCol As Integer
Dim scrapCol As Integer
Dim d As Integer
Dim n As Integer
Dim zfinStr() As String
Dim bomStr() As String
Dim zfinZforStr() As String
Dim operStr() As String
Dim machStr() As String
Dim operData() As String
Dim operNos As String
Dim counter As Integer
Dim boundery As Integer
Dim zzBoundery As Integer
Dim text As String
Dim lastRow As Long
Dim error As Boolean
Dim s As Long
Dim sSql As String
Dim uSql As String
Dim dSql As String
Dim firstDate As Date
Dim lastDate As Date
Dim theType As String
Dim materialType As String
Dim verId As Long
Dim zzCounter As Integer
Dim v() As String
Dim sessionId As Integer


On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If zfinCol > 0 And nCol > 0 And uCol > 0 And tCol > 0 And statCol > 0 And minCol > 0 And maxCol > 0 And rvCol > 0 And scrapCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 50
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Material"
                    zfinCol = n
                    Case Is = "Material description"
                    nCol = n
                    Case Is = "BUn"
                    uCol = n
                    Case Is = "MTyp"
                    tCol = n
                    Case Is = "MS"
                    statCol = n
                    Case Is = "Min.lot size"
                    minCol = n
                    Case Is = "Max. lot size"
                    maxCol = n
                    Case Is = "Rounding val."
                    rvCol = n
                    Case Is = "C.scrap"
                    scrapCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""MRP View Data (IMPECT_MRP)"" z R/3 SQ01. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        'we identified all columns needed. Let's crack on
        lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        sht.Range(Cells(d, nCol), Cells(lastRow, nCol)).Replace "'", "", xlPart
        updateConnection (600)
        counter = 0
        For s = d To 50000
            If sht.Cells(s, zfinCol).Value = "" Then
                Exit For
            Else
                If InStr(1, sht.Cells(s, tCol).Value, "zcom", vbTextCompare) = 0 And InStr(1, sht.Cells(s, tCol).Value, "zfor", vbTextCompare) = 0 And InStr(1, sht.Cells(s, tCol).Value, "zfin", vbTextCompare) = 0 And InStr(1, sht.Cells(s, tCol).Value, "zpkg", vbTextCompare) = 0 Then
                    MsgBox "Artykuł " & sht.Cells(s, zfinCol).Value & " ma nieprawidłowy typ (" & sht.Cells(s, tCol).Value & ") i zostanie pominięty"
                Else
                    If counter = 0 Then
                        boundery = 0
                        ReDim zfinStr(0) As String
                    ElseIf counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve zfinStr(boundery) As String
                    End If
                    theType = LCase(sht.Cells(s, tCol).Value)
                    materialType = "NULL"
                    If theType = "zpkg" Then
                         If InStr(1, CStr(sht.Cells(s, nCol).Value), " WRO ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " FOIL ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "WRO ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "FOIL ", vbTextCompare) = 1 Then
                            materialType = "2"
                        ElseIf (InStr(1, CStr(sht.Cells(s, nCol).Value), " BX ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " BOX ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "BX ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "BOX ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " LD ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "LD ", vbTextCompare) = 1 Or InStr(1, CStr(sht.Cells(s, nCol).Value), " TR ", vbTextCompare) > 0 Or InStr(1, CStr(sht.Cells(s, nCol).Value), "TR ", vbTextCompare) = 1) And Left(CStr(sht.Cells(s, nCol).Value), 2) <> "ST" Then
                            materialType = "4"
                        End If
                    End If
                    zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'" & sht.Cells(s, nCol).Value & "','" & theType & "','" & sht.Cells(s, statCol).Value & "'," & materialType & ",'" & sht.Cells(s, uCol).Value & "'," & Replace(CStr(sht.Cells(s, minCol).Value), ",", ".") & "," & Replace(CStr(sht.Cells(s, maxCol).Value), ",", ".") & "," & Replace(CStr(sht.Cells(s, rvCol).Value), ",", ".") & "),"
                    counter = counter + 1
                End If
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2), materialType int, basicUom nvarchar(10), minLotSize float, maxLotSize float, roundingValue float)"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            If zfinStr(counter) <> "" Then zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus,materialType,basicUom,minLotSize,maxLotSize,roundingValue) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName, t1.zfinType = t2.zfinType, t1.prodStatus = t2.prodStatus, t1.basicUom = t2.basicUom, t1.minLotSize = t2.minLotSize, t1.maxLotSize = t2.maxLotSize, t1.roundingValue = t2.roundingValue FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex WHERE t2.zfinName IS NOT NULL"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,materialType,basicUom,minLotSize,maxLotSize,roundingValue,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin WHERE zfinIndex IS NOT NULL)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,materialType,basicUom,minLotSize,maxLotSize,roundingValue,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor','zpkg','zcom'"

        'let's add scrap reconciliation information
        
        Set rs = AdoConn.Execute("INSERT INTO tbScrapReconciliation (dateAdded, createdBy) VALUES ('" & Now & "', 43);SELECT SCOPE_IDENTITY()")
        Set rs = rs.NextRecordset
        sessionId = rs.Fields(0)
        rs.Close
        Set rs = Nothing
            
        counter = 0

        For s = d To 50000
            
            If sht.Cells(s, zfinCol).Value = "" Then
                For counter = LBound(bomStr) To UBound(bomStr)
                    bomStr(counter) = Left(bomStr(counter), Len(bomStr(counter)) - 1)
                Next counter
                Exit For
            Else
                If InStr(1, sht.Cells(s, tCol).Value, "zcom", vbTextCompare) > 0 Or InStr(1, sht.Cells(s, tCol).Value, "zfor", vbTextCompare) > 0 Or InStr(1, sht.Cells(s, tCol).Value, "zfin", vbTextCompare) > 0 Or InStr(1, sht.Cells(s, tCol).Value, "zpkg", vbTextCompare) > 0 Then
                    If counter = 0 Then
                        boundery = 0
                        ReDim bomStr(0) As String
                    ElseIf counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve bomStr(boundery) As String
                    End If
                    bomStr(boundery) = bomStr(boundery) & "(" & zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId & "," & Replace(CStr(sht.Cells(s, scrapCol).Value), ",", ".") & "," & sessionId & "),"
                    counter = counter + 1
                End If
            End If
        Next s

        cSql = "CREATE TABLE #scraps(zfinId int, scrap float, scrapReconciliationId int)"
        AdoConn.Execute cSql
        For counter = LBound(bomStr) To UBound(bomStr)
            iSql = "INSERT INTO #scraps(zfinId,scrap,scrapReconciliationId) VALUES " & bomStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT zfinId,scrap,scrapReconciliationId FROM #scraps "
        iSql = "INSERT INTO tbComponentScrap (zfinId,scrap,scrapReconciliationId) " & sSql
        AdoConn.Execute iSql

    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
closeConnection
Exit Sub

err_trap:
error = True
MsgBox "Error in importComponentScrap. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here


End Sub


Public Sub importConnections()
Dim sTime As Date
Dim eTime As Date
Dim sht As Worksheet
Dim found As Boolean
Dim i As Integer
Dim d As Integer
Dim n As Integer
Dim y As Integer
Dim text As String
Dim oCol As Integer
Dim oZfinCol As Integer
Dim oZforCol As Integer
Dim zfinCol As Integer
Dim zforCol As Integer
Dim rStatCol As Integer
Dim gStatCol As Integer
Dim rMesStringCol As Integer
Dim gMesStringCol As Integer
Dim pMesStringCol As Integer
Dim zStr As String
Dim oZforStr As String
Dim oZfinStr As String
Dim oStr As String
Dim zVar As Variant
Dim oVar As Variant
Dim v() As String
Dim statCol As Integer
Dim zfinStr() As String
Dim operStr() As String
Dim counter As Integer
Dim boundery As Long
Dim cSql As String
Dim iSql As String
Dim uSql As String
Dim sSql As String
Dim s As Long
Dim sapStr As String
Dim error As Boolean
Dim var As Variant

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If oCol > 0 And zfinCol > 0 And zforCol > 0 And oZfinCol > 0 And oZforCol > 0 And statCol > 0 And rStatCol > 0 And rMesStringCol > 0 And pMesStringCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "ID oper. pak."
                    oCol = n
                    Case Is = "Nr zlec. pak."
                    oZfinCol = n
                    Case Is = "Nr zl. blendu"
                    oZforCol = n
                    Case Is = "Wyrób gotowy"
                    zfinCol = n
                    Case Is = "Nr prod. blendu"
                    zforCol = n
                    Case Is = "Status pak."
                    statCol = n
                    Case Is = "Status praż."
                    rStatCol = n
                    Case Is = "Nr oper. prażenia"
                    rMesStringCol = n
                    Case Is = "Nr oper. pak."
                    pMesStringCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        error = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Powiązania operacji"" z MES. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        updateConnection
        'continue
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfins & zfors indexes -------------------------------------
        counter = 0
        For s = d To 50000
            If .Cells(s, zfinCol).Value <> "" And .Cells(s, zforCol).Value <> "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                zStr = unEnter(sht.Cells(s, zfinCol).Value)
                zfinStr(boundery) = zfinStr(boundery) & "(" & zStr & ",'zfin','pr'),"
                counter = counter + 1
                If counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                var = unEnter(sht.Cells(s, zforCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone")
                If IsArray(var) Then
                    For y = LBound(var) To UBound(var)
                        zStr = var(y)
                        zfinStr(boundery) = zfinStr(boundery) & "(" & zStr & ",'zfor','pr'),"
                        counter = counter + 1
                        If counter Mod 1000 = 0 Then
                            'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                            boundery = counter / 1000
                            ReDim Preserve zfinStr(boundery) As String
                        End If
                    Next y
                Else
                    zStr = var
                    zfinStr(boundery) = zfinStr(boundery) & "(" & zStr & ",'zfor','pr'),"
                    counter = counter + 1
                End If
            ElseIf .Cells(s, zfinCol).Value = "" And .Cells(s, zforCol).Value = "" Then
                For counter = LBound(zfinStr) To UBound(zfinStr)
                    zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
        cSql = "CREATE TABLE #zfins(zfinIndex int,zfinType nchar(4),prodStatus nchar(2))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinType,prodStatus) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT zfinIndex,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy  FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinType,prodStatus,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin','zfor'"
        
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfors'/zfins' order data -------------------------------------
        
        counter = 0
        For s = d To 50000
            If .Cells(s, oZforCol).Value <> "" And .Cells(s, zforCol).Value <> "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                zVar = unEnter(sht.Cells(s, zforCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone", True)
                var = unEnter(sht.Cells(s, oZforCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone", True)
                If IsArray(var) Then
                    For y = LBound(var) To UBound(var)
                        oZforStr = var(y)
                        zStr = zVar(y)
                        operStr(boundery) = operStr(boundery) & "(" & oZforStr & ",'r'," & zfins(zStr).zfinId & "),"
                        sapStr = sapStr & oZforStr & ","
                        counter = counter + 1
                        If counter Mod 1000 = 0 Then
                            'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                            boundery = counter / 1000
                            ReDim Preserve operStr(boundery) As String
                        End If
                    Next y
                Else
                    oZforStr = var
                    zStr = zVar
                    operStr(boundery) = operStr(boundery) & "(" & oZforStr & ",'r'," & zfins(zStr).zfinId & "),"
                    sapStr = sapStr & oZforStr & ","
                    counter = counter + 1
                    If counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve operStr(boundery) As String
                    End If
                End If
                oZfinStr = unEnter(sht.Cells(s, oZfinCol).Value)
                zStr = unEnter(sht.Cells(s, zfinCol).Value)
                operStr(boundery) = operStr(boundery) & "(" & oZfinStr & ",'p'," & zfins(zStr).zfinId & "),"
                sapStr = sapStr & oZfinStr & ","
                counter = counter + 1
            ElseIf .Cells(s, oZforCol).Value = "" And .Cells(s, zforCol).Value = "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                sapStr = Left(sapStr, Len(sapStr) - 1)
                For counter = LBound(operStr) To UBound(operStr)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #orders(sapId bigint, type nchar(1),zfinId bigint)"
            AdoConn.Execute cSql
            For counter = LBound(operStr) To UBound(operStr)
                iSql = "INSERT INTO #orders(sapId,type,zfinId) VALUES " & operStr(counter)
                AdoConn.Execute iSql
            Next counter
            uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId FROM tbOrders t1 INNER JOIN #orders t2 ON t1.sapId = t2.sapId"
            AdoConn.Execute uSql
            sSql = "SELECT DISTINCT sapId,type,zfinId,GETDATE() as createdOn FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId is not null)"
            iSql = "INSERT INTO tbOrders (sapId,type,zfinId, createdOn) " & sSql
            AdoConn.Execute iSql
        
        downloadOrders sapStr:=sapStr
        
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfins' opearation data -------------------------------------
        counter = 0
        Erase operStr
        
        For s = d To 50000
            If .Cells(s, oZfinCol).Value <> "" And .Cells(s, zfinCol).Value <> "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                operStr(boundery) = operStr(boundery) & "(" & sht.Cells(s, oCol).Value & ",'p'," & zfins(CStr(sht.Cells(s, zfinCol).Value)).zfinId & ",'" & Now & "'," & orders(CStr(sht.Cells(s, oZfinCol).Value)).orderId & "),"
                counter = counter + 1
            ElseIf .Cells(s, oZfinCol).Value = "" And .Cells(s, zfinCol).Value = "" Then
                For counter = LBound(operStr) To UBound(operStr)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #operations(mesId int, type nchar(1), zfinId bigint, createdOn datetime, orderId int)"
            AdoConn.Execute cSql
            For counter = LBound(operStr) To UBound(operStr)
                iSql = "INSERT INTO #operations(mesId,type,zfinId,createdOn, orderId) VALUES " & operStr(counter)
                AdoConn.Execute iSql
            Next counter
            uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId, t1.orderId = t2.orderId FROM tbOperations t1 INNER JOIN #operations t2 ON t1.mesId = t2.mesId"
            AdoConn.Execute uSql
            sSql = "SELECT DISTINCT mesId,type,zfinId, createdOn, orderId FROM #operations WHERE mesId NOT IN (SELECT mesId FROM tbOperations WHERE mesId is not null)"
            iSql = "INSERT INTO tbOperations (mesId,type,zfinId, createdOn,orderId) " & sSql
            AdoConn.Execute iSql
        
         '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing sapId numbers to opearations -------------------------------------
        counter = 0
        Erase operStr
        
        For s = d To 50000
            If .Cells(s, oZfinCol).Value <> "" And .Cells(s, zfinCol).Value <> "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                var = unEnter(sht.Cells(s, rMesStringCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone", True)
                oVar = unEnter(sht.Cells(s, oZforCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone", True)
                If IsArray(var) And IsArray(oVar) = False Then
                    'perform only if there's array of mes strings and 1 zfor sapId
                    For y = LBound(var) To UBound(var)
                        operStr(boundery) = operStr(boundery) & "('" & deHash(CStr(var(y))) & "'," & orders(oVar).orderId & "),"
                        counter = counter + 1
                        If counter Mod 1000 = 0 Then
                            'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                            boundery = counter / 1000
                            ReDim Preserve operStr(boundery) As String
                        End If
                    Next y
                ElseIf IsArray(var) And IsArray(oVar) Then
                    If UBound(var) = UBound(oVar) Then
                        For y = LBound(var) To UBound(var)
                            operStr(boundery) = operStr(boundery) & "('" & deHash(CStr(var(y))) & "'," & orders(oVar(y)).orderId & "),"
                            counter = counter + 1
                            If counter Mod 1000 = 0 Then
                                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                                boundery = counter / 1000
                                ReDim Preserve operStr(boundery) As String
                            End If
                        Next y
                        'perform only if upper bound of both arrays are equal so there's sapId for each mes string
                    End If
                ElseIf IsArray(var) = False And IsArray(oVar) = False Then
                    operStr(boundery) = operStr(boundery) & "('" & deHash(CStr(var)) & "'," & orders(oVar).orderId & "),"
                    counter = counter + 1
                    If counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve operStr(boundery) As String
                    End If
                End If
                operStr(boundery) = operStr(boundery) & "('" & deHash(sht.Cells(s, pMesStringCol).Value) & "'," & orders(CStr(sht.Cells(s, oZfinCol).Value)).orderId & "),"
                counter = counter + 1
            ElseIf .Cells(s, oZfinCol).Value = "" And .Cells(s, zfinCol).Value = "" Then
                For counter = LBound(operStr) To UBound(operStr)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #operations2(mesString nchar(50), orderId int)"
            AdoConn.Execute cSql
            For counter = LBound(operStr) To UBound(operStr)
                iSql = "INSERT INTO #operations2(mesString, orderId) VALUES " & operStr(counter)
                AdoConn.Execute iSql
            Next counter
            uSql = "UPDATE t1 SET t1.orderId = t2.orderId FROM tbOperations t1 INNER JOIN #operations2 t2 ON t1.mesString = t2.mesString"
            AdoConn.Execute uSql
        '----------------------------------------------------------------------------------------------
        '--------------- Let's add missing zfins-zfor connections -------------------------------------
        counter = 0
        Erase operStr
        
        For s = d To 50000
            If .Cells(s, oZfinCol).Value <> "" And .Cells(s, oZforCol).Value <> "" And .Cells(s, statCol).Value <> "Rezygnacja" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim operStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve operStr(boundery) As String
                End If
                oZfinStr = unEnter(sht.Cells(s, oZfinCol).Value)
                var = unEnter(sht.Cells(s, oZforCol).Value, sht.Cells(s, rStatCol).Value, "Zakończone")
                If IsArray(var) Then
                    For y = LBound(var) To UBound(var)
                        oZforStr = var(y)
                        operStr(boundery) = operStr(boundery) & "(" & orders(oZforStr).orderId & "," & orders(oZfinStr).orderId & "),"
                        counter = counter + 1
                        If counter Mod 1000 = 0 Then
                            'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                            boundery = counter / 1000
                            ReDim Preserve operStr(boundery) As String
                        End If
                    Next y
                Else
                    oZforStr = var
                    operStr(boundery) = operStr(boundery) & "(" & orders(oZforStr).orderId & "," & orders(oZfinStr).orderId & "),"
                    counter = counter + 1
                End If
            ElseIf .Cells(s, oZfinCol).Value = "" And .Cells(s, oZforCol).Value = "" Then
                For counter = LBound(operStr) To UBound(operStr)
                    operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
                Next counter
                Exit For
            End If
        Next s
            
            cSql = "CREATE TABLE #orderDep(zforOrder int, zfinOrder int)"
            AdoConn.Execute cSql
            For counter = LBound(operStr) To UBound(operStr)
                iSql = "INSERT INTO #orderDep(zforOrder, zfinOrder) VALUES " & operStr(counter)
                AdoConn.Execute iSql
            Next counter
'            uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId, t1.sapId = t2.sapId FROM tbOperations t1 INNER JOIN #operations1 t2 ON t1.mesId = t2.mesId"
'            adoConn.Execute uSql
            sSql = "SELECT DISTINCT zforOrder, zfinOrder FROM #orderDep tod WHERE NOT EXISTS (SELECT * FROM tbOrderDep od WHERE od.zfinOrder=tod.zfinOrder AND od.zforOrder=tod.zforOrder)"
            iSql = "INSERT INTO tbOrderDep (zforOrder, zfinOrder) " & sSql
            AdoConn.Execute iSql
        
    End If
End With

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importConnections"". Error number: " & Err.Number & ", " & Err.Description
error = True
Resume Exit_here

End Sub

Public Function deHash(val As String) As String

If InStr(1, val, "#", vbTextCompare) > 0 Then
    deHash = Left(val, InStr(1, val, "#", vbTextCompare) - 1)
Else
    deHash = val
End If

End Function


Public Sub importBatch2order()
Dim sht As Worksheet
Dim error As Boolean
Dim found As Boolean
Dim i As Integer
Dim zfinStr() As String
Dim operStr() As String
Dim batchStr() As String
Dim counter As Integer
Dim bStr As String
Dim oStr As String
Dim boundery As Integer
Dim cSql As String
Dim iSql As String
Dim sSql As String
Dim uSql As String
Dim sTime As Date
Dim eTime As Date
Dim oCol As Integer
Dim pCol As Integer
Dim nCol As Integer
Dim uCol As Integer
Dim bCol As Integer
Dim aCol As Integer
Dim d As Integer
Dim n As Integer
Dim text As String
Dim v() As String
Dim colLetter As String
Dim lastRow As Long
Dim theType As String

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet

found = False
For i = 1 To 30
    If oCol > 0 And pCol > 0 And nCol > 0 And uCol > 0 And bCol > 0 And aCol > 0 Then
        d = i
        found = True
        Exit For
    Else
        For n = 1 To 30
            text = sht.Cells(i, n).Value
            Select Case text
                Case Is = "Order"
                oCol = n
                Case Is = "Material Number"
                pCol = n
                Case Is = "Material description"
                nCol = n
                Case Is = "Unit of measure (=GMEIN)"
                uCol = n
                Case Is = "Batch"
                bCol = n
                Case Is = "Delivered quantity (GMEIN)"
                aCol = n
            End Select
        Next n
    End If
Next i

If found Then
    v() = Split(sht.Cells(1, oCol).Address, "$", , vbTextCompare)
    colLetter = v(1)
    lastRow = sht.Range(colLetter & ":" & colLetter).Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
    sht.Range(Cells(2, nCol), Cells(lastRow, nCol)).Replace "'", "", LookAt:=xlPart
    sht.Range(Cells(2, nCol), Cells(lastRow, nCol)).Replace ".", "", LookAt:=xlPart
    updateConnection
    'continue
    '----------------------------------------------------------------------------------------------
    '--------------- Let's add missing zfins indexes -------------------------------------
    counter = 0
    For i = 1 To 30000
        theType = ""
        If (sht.Cells(i, uCol).Value = "PC" Or sht.Cells(i, uCol).Value = "KG") Then
            If sht.Cells(i, uCol).Value = "PC" Then
                theType = "zfin"
            Else
                theType = "zfor"
            End If
            If counter = 0 Then
                boundery = 0
                ReDim zfinStr(0) As String
            ElseIf counter Mod 1000 = 0 Then
                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                boundery = counter / 1000
                ReDim Preserve zfinStr(boundery) As String
            End If
            zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(i, pCol).Value & ",'" & sht.Cells(i, nCol).Value & "','" & theType & "','pr'),"
            counter = counter + 1
        ElseIf sht.Cells(i, oCol).Value = "" Or sht.Cells(i, pCol).Value = "" Then
            For counter = LBound(zfinStr) To UBound(zfinStr)
                zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            Next counter
            Exit For
        End If
    Next i
    
    cSql = "CREATE TABLE #zfins(zfinIndex int,zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2))"
    AdoConn.Execute cSql
    For counter = LBound(zfinStr) To UBound(zfinStr)
        iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
        AdoConn.Execute iSql
    Next counter
    uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex"
    AdoConn.Execute uSql
    sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
    iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
    AdoConn.Execute iSql
    
    downloadZfins "'zfin','zfor'"
    
    '----------------------------------------------------------------------------------------------
    '--------------- Let's add missing operations -------------------------------------
    counter = 0
    For i = 1 To 30000
        theType = ""
        If (sht.Cells(i, uCol).Value = "PC" Or sht.Cells(i, uCol).Value = "KG") Then
            If sht.Cells(i, uCol).Value = "PC" Then
                theType = "p"
            Else
                theType = "r"
            End If
            If counter = 0 Then
                boundery = 0
                ReDim operStr(0) As String
            ElseIf counter Mod 1000 = 0 Then
                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                boundery = counter / 1000
                ReDim Preserve operStr(boundery) As String
            End If
            operStr(boundery) = operStr(boundery) & "(" & sht.Cells(i, oCol).Value & ",'" & theType & "'," & zfins(CStr(sht.Cells(i, pCol).Value)).zfinId & "," & Replace(CDbl(sht.Cells(i, aCol).Value), ",", ".") & "),"
            counter = counter + 1
        ElseIf sht.Cells(i, oCol).Value = "" And sht.Cells(i, pCol).Value = "" Then
            For counter = LBound(operStr) To UBound(operStr)
                operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
            Next counter
            Exit For
        End If
    Next i
    
    cSql = "CREATE TABLE #orders(sapId bigint,type nchar(1), zfinId bigint, executedSap float)"
    AdoConn.Execute cSql
    For counter = LBound(operStr) To UBound(operStr)
        iSql = "INSERT INTO #orders(sapId,type,zfinId,executedSap) VALUES " & operStr(counter)
        AdoConn.Execute iSql
    Next counter
    uSql = "UPDATE t1 SET t1.zfinId = t2.zfinId, t1.executedSap = t2.executedSap FROM tbOrders t1 INNER JOIN #orders t2 ON t1.sapId = t2.sapId"
    AdoConn.Execute uSql
    sSql = "SELECT DISTINCT sapId,type,zfinId,GETDATE() as createdOn,executedSap FROM #orders WHERE sapId NOT IN (SELECT sapId FROM tbOrders WHERE sapId IS NOT NULL)"
    iSql = "INSERT INTO tbOrders (sapId,type,zfinId,createdOn,executedSap) " & sSql
    AdoConn.Execute iSql
    
Else
    error = True
    MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""COOIS"" z SAP R/3 w wariancie ""/M024ROB"". Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
End If

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importBatch2order"". Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Sub importReqs()
Dim sht As Worksheet
Dim error As Boolean
Dim found As Boolean
Dim i As Integer
Dim zfinStr() As String
Dim operStr() As String
Dim batchStr() As String
Dim counter As Integer
Dim bStr As String
Dim oStr As String
Dim boundery As Integer
Dim cSql As String
Dim iSql As String
Dim sSql As String
Dim uSql As String
Dim sTime As Date
Dim eTime As Date
Dim dCol As Integer
Dim cCol As Integer
Dim pCol As Integer
Dim oCol As Integer
Dim pdCol As Integer
Dim tCol As Integer
Dim aCol As Integer
Dim d As Integer
Dim n As Integer
Dim text As String
Dim v() As String
Dim colLetter As String
Dim lastRow As Long
Dim theType As String
Dim firstDate As Date
Dim firstRow As Integer
Dim amount As Double
Dim str() As String
Dim custStr As String
Dim loc As Variant

On Error GoTo err_trap

sTime = Now

Set sht = ActiveWorkbook.ActiveSheet

found = False
For i = 1 To 30
    If dCol > 0 And cCol > 0 And pCol > 0 And pdCol > 0 And tCol > 0 And aCol > 0 Then
        d = i
        found = True
        Exit For
    Else
        For n = 1 To 30
            text = sht.Cells(i, n).Value
            Select Case text
                Case Is = "Availability Date / Requirements Date"
                dCol = n
                Case Is = "Category Short Text"
                cCol = n
                Case Is = "Product Number"
                pCol = n
                Case Is = "Product Short Description"
                pdCol = n
                Case Is = "Target Location"
                tCol = n
                Case Is = "Receipt Quantity / Requirements Quantity"
                aCol = n
                Case Is = "Receipt Element / Requirements Element"
                oCol = n
            End Select
        Next n
    End If
Next i

If found Then
    firstRow = i
    v() = Split(sht.Cells(1, pCol).Address, "$", , vbTextCompare)
    colLetter = v(1)
    lastRow = sht.Range(colLetter & ":" & colLetter).Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
    sht.Range(Cells(firstRow, pdCol), Cells(lastRow, pdCol)).Replace "'", "", LookAt:=xlPart
    sht.Range(Cells(firstRow, pdCol), Cells(lastRow, pdCol)).Replace ".", "", LookAt:=xlPart
    updateConnection
    'continue
    '----------------------------------------------------------------------------------------------
    '--------------- Let's add missing zfins indexes -------------------------------------
    counter = 0
    For i = firstRow To 30000
        If (sht.Cells(i, cCol).Value = "SalesOrder" Or sht.Cells(i, cCol).Value = "DEP:ConRl") And sht.Cells(i, tCol) <> "M024" Then
'            If sht.Cells(i, cCol).Value = "SalesOrder" Then
'                theType = "so"
'            ElseIf sht.Cells(i, cCol).Value = "DEP:ConRl" Then
'                theType = "po"
'            Else
'                theType = ""
'            End If
            If counter = 0 Then
                boundery = 0
                ReDim zfinStr(0) As String
            ElseIf counter Mod 1000 = 0 Then
                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                boundery = counter / 1000
                ReDim Preserve zfinStr(boundery) As String
            End If
            zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(i, pCol).Value & ",'" & sht.Cells(i, pdCol).Value & "','zfin','pr'),"
            counter = counter + 1
        ElseIf sht.Cells(i, pCol).Value = "" Or sht.Cells(i, pdCol).Value = "" Then
            For counter = LBound(zfinStr) To UBound(zfinStr)
                zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
            Next counter
            Exit For
        End If
    Next i
    
    cSql = "CREATE TABLE #zfins(zfinIndex int,zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2))"
    AdoConn.Execute cSql
    For counter = LBound(zfinStr) To UBound(zfinStr)
        iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
        AdoConn.Execute iSql
    Next counter
    uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZFin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex"
    AdoConn.Execute uSql
    sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy  FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
    iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
    AdoConn.Execute iSql
    
    downloadZfins "'zfin'"
    
    '----------------------------------------------------------------------------------------------
    '--------------- Let's add missing target locations ------------------------------------------
    counter = 0
    For i = firstRow To 30000
        If Len(sht.Cells(i, tCol).Value) > 0 And sht.Cells(i, tCol).Value <> "M024" Then
            If counter = 0 Then
                boundery = 0
                ReDim batchStr(0) As String
            ElseIf counter Mod 1000 = 0 Then
                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                boundery = counter / 1000
                ReDim Preserve batchStr(boundery) As String
            End If
            batchStr(boundery) = batchStr(boundery) & "('" & sht.Cells(i, tCol).Value & "'),"
            counter = counter + 1
        ElseIf sht.Cells(i, pCol).Value = "" And sht.Cells(i, pdCol).Value = "" Then
            For counter = LBound(batchStr) To UBound(batchStr)
                batchStr(counter) = Left(batchStr(counter), Len(batchStr(counter)) - 1)
            Next counter
            Exit For
        End If
    Next i
    
    cSql = "CREATE TABLE #locations(location varchar(10))"
    AdoConn.Execute cSql
    For counter = LBound(batchStr) To UBound(batchStr)
        iSql = "INSERT INTO #locations(location) VALUES " & batchStr(counter)
        AdoConn.Execute iSql
    Next counter
    sSql = "SELECT DISTINCT location FROM #locations WHERE location NOT IN (SELECT location FROM tbCustomerString WHERE location IS NOT NULL)"
    iSql = "INSERT INTO tbCustomerString (location) " & sSql
    AdoConn.Execute iSql
    
    downloadLocations
    
    '----------------------------------------------------------------------------------------------
    '--------------- Let's add missing reqs -------------------------------------
    counter = 0
    firstDate = CDate(Application.WorksheetFunction.Min(sht.Range(Cells(2, dCol), Cells(lastRow, dCol))))

    AdoConn.Execute "DELETE FROM tbReqs WHERE deliveryDate >= '" & firstDate & "'"
    For i = firstRow To 30000
        If (sht.Cells(i, cCol).Value = "SalesOrder" Or sht.Cells(i, cCol).Value = "DEP:ConRl") And sht.Cells(i, tCol) <> "M024" Then
            If sht.Cells(i, cCol).Value = "SalesOrder" Then
                theType = "so"
            ElseIf sht.Cells(i, cCol).Value = "DEP:ConRl" Then
                theType = "po"
            Else
                theType = Null
            End If
            If IsNumeric(sht.Cells(i, aCol)) Then
                amount = Abs(sht.Cells(i, aCol))
            Else
                amount = 0
            End If
            If theType = "so" Then
                str = Split(sht.Cells(i, oCol), "/")
                If Not isArrayEmpty(str) Then
                    custStr = "'" & str(0) & "'"
                Else
                    custStr = "NULL"
                End If
            Else
                custStr = "NULL"
            End If
            If Len(sht.Cells(i, tCol) & vbNullString) = 0 Then
                loc = "NULL"
            Else
                loc = locations(CStr(sht.Cells(i, tCol))).locationId
            End If
            If Not IsNull(theType) Then
                If (theType = "so" And DateDiff("ww", Date, CDate(sht.Cells(i, dCol)), vbMonday) <= 2) Or (theType = "po" And DateDiff("ww", Date, CDate(sht.Cells(i, dCol)), vbMonday) <= 1) Then
                    If counter = 0 Then
                        boundery = 0
                        ReDim operStr(0) As String
                    ElseIf counter Mod 1000 = 0 Then
                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                        boundery = counter / 1000
                        ReDim Preserve operStr(boundery) As String
                    End If
                    operStr(boundery) = operStr(boundery) & "(" & zfins(CStr(sht.Cells(i, pCol))).zfinId & "," & amount & ",'" & theType & "'," & loc & "," & custStr & ",'" & sht.Cells(i, dCol) & "','" & Now & "'),"
                    counter = counter + 1
                End If
            End If
        ElseIf sht.Cells(i, oCol).Value = "" And sht.Cells(i, pCol).Value = "" Then
            For counter = LBound(operStr) To UBound(operStr)
                operStr(counter) = Left(operStr(counter), Len(operStr(counter)) - 1)
            Next counter
            Exit For
        End If
    Next i
    
    For counter = LBound(operStr) To UBound(operStr)
        iSql = "INSERT INTO tbReqs(zfinId,amount,type,target,custOrder,deliveryDate, creationDate) VALUES " & operStr(counter)
        AdoConn.Execute iSql
    Next counter
Else
    error = True
    MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Requirements View"" z SAP APO. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
End If

Exit_here:
closeConnection
eTime = Now
If Not error Then MsgBox "Zapis zakończony powodzeniem w " & Abs(DateDiff("s", sTime, eTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
Exit Sub

err_trap:
MsgBox "Error in ""importReqs"". Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub



Public Sub printMachs()
Dim m As clsMach

For Each m In machs
    Debug.Print m.machId & ", " & m.machStr
Next m

End Sub

Private Sub downloadZfins(typeStr As String)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newZfin As clsZfin

On Error GoTo err_trap

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

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadZfins. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadOperations(Optional fromD As Variant, Optional mesStr As Variant, Optional sapStr As Variant)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newOperation As clsOperation
Dim key As Long

On Error GoTo err_trap

n = operations.count
Do While operations.count > 0
    operations.Remove n
    n = n - 1
Loop

If Not IsMissing(fromD) Then
    sSql = "SELECT DISTINCT o.operationId, oo.orderId, o.mesId, oo.sapId FROM tbOperations o LEFT JOIN tbOrders oo ON o.orderId=oo.orderId WHERE o.createdOn >= '" & fromD & "'"
Else
    If Not IsMissing(mesStr) Then
        sSql = "SELECT DISTINCT o.operationId, oo.orderId, o.mesId, oo.sapId FROM tbOperations o LEFT JOIN tbOrders oo ON o.orderId=oo.orderId WHERE o.mesId IN (" & mesStr & ");"
    Else
        If Not IsMissing(sapStr) Then
            sSql = "SELECT DISTINCT o.operationId, oo.orderId, o.mesId, oo.sapId FROM tbOperations o RIGHT JOIN tbOrders oo ON o.orderId=oo.orderId WHERE oo.sapId IN (" & sapStr & ");"
        Else
            sSql = "SELECT DISTINCT o.operationId, oo.orderId, o.mesId, oo.sapId FROM tbOperations o RIGHT JOIN tbOrders oo ON o.orderId=oo.orderId;"
        End If
    End If
End If
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newOperation = New clsOperation
        With newOperation
            If Not IsNull(rs.Fields("operationId").Value) Then .operationId = rs.Fields("operationId").Value
            If Not IsNull(rs.Fields("orderId").Value) Then .orderId = rs.Fields("orderId").Value
            If Not IsNull(rs.Fields("sapId").Value) Then
                .sapId = rs.Fields("sapId").Value
            End If
            If Not IsNull(rs.Fields("mesId").Value) And Not IsMissing(mesStr) Then
                .mesId = rs.Fields("mesId").Value
                key = rs.Fields("mesId").Value
            End If
            operations.Add newOperation, CStr(key)
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadOperations. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadOrders(Optional mesStr As Variant, Optional sapStr As Variant)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newOrder As clsOrder
Dim key As Long

On Error GoTo err_trap

n = orders.count
Do While orders.count > 0
    orders.Remove n
    n = n - 1
Loop

If Not IsMissing(mesStr) Then
    sSql = "SELECT DISTINCT oo.orderId, oo.sapId FROM tbOperations o RIGHT JOIN tbOrders oo ON o.orderId=oo.orderId WHERE o.mesId IN (" & mesStr & ");"
Else
    If Not IsMissing(sapStr) Then
        sSql = "SELECT DISTINCT oo.orderId, oo.sapId FROM tbOperations o RIGHT JOIN tbOrders oo ON o.orderId=oo.orderId WHERE oo.sapId IN (" & sapStr & ");"
    Else
        sSql = "SELECT DISTINCT oo.orderId, oo.sapId FROM tbOperations o RIGHT JOIN tbOrders oo ON o.orderId=oo.orderId;"
    End If
End If

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newOrder = New clsOrder
        With newOrder
            If Not IsNull(rs.Fields("orderId").Value) Then .orderId = rs.Fields("orderId").Value
            If Not IsNull(rs.Fields("sapId").Value) Then
                .sapId = rs.Fields("sapId").Value
            End If
            orders.Add newOrder, CStr(rs.Fields("sapId").Value)
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadOrders. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadBatches(Optional bStr As Variant)
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newBatch As clsBatch

On Error GoTo err_trap

n = batches.count
Do While batches.count > 0
    batches.Remove n
    n = n - 1
Loop

If Not IsMissing(bStr) Then
    sSql = "SELECT batchId, batchNumber FROM tbBatch WHERE batchNumber IN (" & bStr & ");"
Else
    sSql = "SELECT batchId, batchNumber FROM tbBatch;"
End If

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newBatch = New clsBatch
        With newBatch
            .bId = rs.Fields("batchId").Value
            .bNumber = rs.Fields("batchNumber").Value
            batches.Add newBatch, CStr(rs.Fields("batchNumber").Value)
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadBatches. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadLocations()
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newLocation As clsLocation

On Error GoTo err_trap

n = locations.count
Do While locations.count > 0
    locations.Remove n
    n = n - 1
Loop

sSql = "SELECT custStringId, location FROM tbCustomerString WHERE location IS NOT NULL;"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newLocation = New clsLocation
        With newLocation
            .locationId = rs.Fields("custStringId").Value
            .locationName = Trim(rs.Fields("location").Value)
            locations.Add newLocation, Trim(rs.Fields("location").Value)
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadLocations. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadQdocs()
Dim rs As ADODB.Recordset
Dim sSql As String
Dim v() As String
Dim n As Integer
Dim newQDoc As clsQdoc
Dim theType As String

On Error GoTo err_trap

n = Qdocs.count
Do While Qdocs.count > 0
    Qdocs.Remove n
    n = n - 1
Loop

sSql = "SELECT qReconciliationId, qNumber, qType FROM tbQdocReconciliation;"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newQDoc = New clsQdoc
        With newQDoc
            .qId = rs.Fields("qReconciliationId").Value
            .qNumber = rs.Fields("qNumber").Value
            theType = Right(rs.Fields("qType").Value, 2)
            Qdocs.Add newQDoc, theType & "_" & CStr(rs.Fields("qNumber").Value)
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadQdocs. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Private Sub downloadMachines()
Dim rs As ADODB.Recordset
Dim sSql As String
Dim newMach As clsMach
Dim n As Integer

On Error GoTo err_trap

n = machs.count
Do While machs.count > 0
    machs.Remove n
    n = n - 1
Loop

sSql = "SELECT machineId, machineName FROM tbMachine;"
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Set newMach = New clsMach
        With newMach
            .machId = rs.Fields("machineId").Value
            .machStr = Trim(rs.Fields("machineName").Value)
            machs.Add newMach, Trim(CStr(rs.Fields("machineName").Value))
        End With
        rs.MoveNext
    Loop
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in downloadMachines. Error number: " & Err.Number & ", " & Err.Description
Resume Exit_here

End Sub

Public Sub updateProperty()
Dim sht As Worksheet
Dim found As Boolean
Dim aCol As Integer
Dim pCol As Integer
Dim vCol As Integer
Dim text As String
Dim n As Integer
Dim s As Long
Dim i As Long
Dim d As Long
Dim tot As Long
Dim upt As Long
Dim sSql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim zfinIdentity As Long
Dim fldName As String
Dim val As Variant
Dim sTime As Date
Dim eTime As Date

On Error GoTo err_trap

Set sht = ActiveWorkbook.ActiveSheet
With sht
    found = False
    For i = 1 To 30
        If aCol > 0 And pCol > 0 And vCol > 0 Then
            d = i
            found = True
            Exit For
        Else
            For n = 1 To 30
                text = .Cells(i, n).Value
                Select Case text
                    Case Is = "Nazwa atrybutu"
                    aCol = n
                    Case Is = "Nr produktu"
                    pCol = n
                    Case Is = "Wartość atrybutu"
                    vCol = n
                End Select
            Next n
        End If
    Next i
    If found = False Then
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Lista przypisań atrybutów do artykułów"" z MES. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    Else
        'we identified all columns needed. Let's crack on
        updateConnection (90)
        downloadZfins "'zfin', 'zfor'"
        'printMachs
        tot = 0
        For s = d To 50000 'For s = d To 50000
            If .Cells(s, aCol).Value = "Kofeinowa" Then
                fldName = "decafe?"
            ElseIf .Cells(s, aCol).Value = "Mielona" Then
                fldName = "beans?"
            ElseIf .Cells(s, aCol).Value = "ORGANICZNA" Then
                fldName = "eco?"
            ElseIf .Cells(s, aCol).Value = "" Then
                MsgBox "Dodano " & tot & " rekordów do bazy danych, zaktualizowano " & upt & " rekordów bazy danych!", vbOKOnly + vbInformation, "Powodzenie"
                Exit For
            Else
                fldName = ""
            End If
            If Len(fldName) > 0 Then
                zfinIdentity = CLng(zfins(CStr(.Cells(s, pCol).Value)).zfinId)
                If zfinIdentity > 0 Then
                    val = Null
                    sSql = "SELECT zfinId, [" & fldName & "] FROM tbZfinProperties WHERE zfinId=" & zfins(CStr(.Cells(s, pCol).Value)).zfinId & ";"
                    Set rs = New ADODB.Recordset
                    rs.Open sSql, AdoConn, adOpenKeyset, adLockOptimistic, adCmdText
                    If rs.EOF Then
                        'insert
                        rs.AddNew
                        rs.Fields("zfinId").Value = zfinIdentity
                        tot = tot + 1
                    Else
                        upt = upt + 1
                    End If
                    If .Cells(s, vCol).Value = "Tak" Then
                        val = 1
                    ElseIf .Cells(s, vCol).Value = "Nie" Then
                        val = 0
                    End If
                    If fldName = "beans?" Or fldName = "decafe?" Then
                        If val = 1 Then
                            val = 0
                        ElseIf val = 0 Then
                            val = 1
                        End If
                    End If
                    rs.Fields(fldName).Value = val
                    rs.Update
                    rs.Close
                    Set rs = Nothing
                End If
            End If
        Next s
    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
If Not rs1 Is Nothing Then
    If rs1.State = 1 Then rs1.Close
    Set rs1 = Nothing
End If
closeConnection
Exit Sub

err_trap:
If Err.Number = 5 Then
    zfinIdentity = 0
    Resume Next
Else
    MsgBox "Error in updateProperty. Error number: " & Err.Number & ", " & Err.Description
    Resume Exit_here
End If

End Sub

Public Sub exportPW_WZ()
Dim c As Range
Dim c1 As Range
Dim c2 As Range
Dim c4 As Range
Dim sht As Worksheet
Dim theType As String 'pw or wz
Dim zfinStr() As String
Dim zfin As Long
Dim batchStr() As String
Dim qDocString() As String
Dim firstAddress As String
Dim sSql As String
Dim deliv As String
Dim cSql As String
Dim eRow As Long
Dim iSql As String
Dim theRow As Long
Dim currZfin As Long
Dim rs As ADODB.Recordset
Dim i As Integer
Dim theFirstOne As Long
Dim theLastOne As Long
Dim rng As Range
Dim c3 As Range
Dim qDocDataStr() As String
Dim currType As String
Dim firstAddress1 As String
Dim sTime As Date
Dim eTime As Date
Dim isError As Boolean
Dim uSql As String
Set sht = ActiveWorkbook.ActiveSheet
Dim counter As Integer
Dim boundery As Integer
Dim bStr As String
Dim lastRow As Long
Dim prevRow As Long

On Error GoTo err_trap

sTime = Now()
isError = False

'        '----------------------------------------------------------------------------------------------
'        '--------------- Let's add missing zfins & zfors indexes -------------------------------------
'        counter = 0
'        For s = d To 50000
'            If .Cells(s, zfinCol).Value <> "" And .Cells(s, zforCol).Value <> "" Then
'                If counter = 0 Then
'                    boundery = 0
'                    ReDim zfinStr(0) As String
'                ElseIf counter Mod 1000 = 0 Then
'                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
'                    boundery = counter / 1000
'                    ReDim Preserve zfinStr(boundery) As String
'                End If
'                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zfinCol).Value & ",'zfin','pr','" & Now & "',43),"
'                zfinStr(boundery) = zfinStr(boundery) & "(" & sht.Cells(s, zforCol).Value & ",'zfor','pr','" & Now & "',43),"
'                counter = counter + 2
'            ElseIf .Cells(s, zfinCol).Value = "" And .Cells(s, zforCol).Value = "" Then
'                For counter = LBound(zfinStr) To UBound(zfinStr)
'                    zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
'                Next counter
'                Exit For
'            End If
'        Next s
lastRow = sht.Range("A:A").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
Set c3 = sht.Cells.Find("Nr magazynu:", LookIn:=xlValues, LookAt:=xlWhole)
If Not c3 Is Nothing Then
    sht.Range(Cells(c3.Row, c3.Column), Cells(lastRow, c3.Column)).Replace "'", "", LookAt:=xlPart
    Set c = sht.Cells.Find("Nr artykułu:", LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            If c.Offset(0, 1) <> "" Then
                If counter = 0 Then
                    boundery = 0
                    ReDim zfinStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve zfinStr(boundery) As String
                End If
                zfinStr(boundery) = zfinStr(boundery) & "(" & c.Offset(0, 1) & ",'" & c.Offset(2, 0) & "','zfin','pr'),"
                counter = counter + 1
            End If
            Set c = sht.Cells.FindNext(c)
        Loop While Not c Is Nothing And c.Address <> firstAddress
    Else
        isError = True
        MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Zestawienie obrotów wg artykułów"" z Qguar. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
    End If
    If Not isError Then
        For counter = LBound(zfinStr) To UBound(zfinStr)
            zfinStr(counter) = Left(zfinStr(counter), Len(zfinStr(counter)) - 1)
        Next counter
        
        updateConnection
        
        cSql = "CREATE TABLE #zfins(zfinIndex int, zfinName nvarchar(255),zfinType nchar(4),prodStatus nchar(2))"
        AdoConn.Execute cSql
        For counter = LBound(zfinStr) To UBound(zfinStr)
            iSql = "INSERT INTO #zfins(zfinIndex,zfinName,zfinType,prodStatus) VALUES " & zfinStr(counter)
            AdoConn.Execute iSql
        Next counter
        uSql = "UPDATE t1 SET t1.zfinName = t2.zfinName FROM tbZfin t1 INNER JOIN #zfins t2 ON t1.zfinIndex = t2.zfinIndex"
        AdoConn.Execute uSql
        sSql = "SELECT DISTINCT zfinIndex,zfinName,zfinType,prodStatus,GETDATE() as creationDate, 43 as createdBy  FROM #zfins WHERE zfinIndex NOT IN (SELECT zfinIndex FROM tbZfin)"
        iSql = "INSERT INTO tbZfin (zfinIndex,zfinName,zfinType,prodStatus,creationDate,createdBy) " & sSql
        AdoConn.Execute iSql
        
        downloadZfins "'zfin'"
        
        counter = 0
        
        Set c1 = sht.Cells.Find("Nr artykułu:", LookIn:=xlValues, LookAt:=xlWhole)
        If Not c1 Is Nothing Then
            firstAddress1 = c1.Address
            theFirstOne = c1.Row
            Do
                If eRow > 0 Then
                    For Each c In sht.Range("G" & eRow & ":G" & c1.Row)
                        If c.Offset(0, -4) = "Typ dokumentu:" Then
                            currType = c.Offset(0, -3)
                        End If
                        If c = "PC" Then
                            currZfin = sht.Range("D" & eRow)
                            If c.Offset(0, 1) <> "" Then
                                If counter = 0 Then
                                    boundery = 0
                                ElseIf counter Mod 1000 = 0 Then
                                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                                    boundery = counter / 1000
                                End If
                                If c.Offset(0, -6) <> "" Then
                                    If c.Offset(0, -3) = "" Then
                                        deliv = 0
                                    Else
                                        If IsNumeric(c.Offset(0, -3)) Then
                                            deliv = c.Offset(0, -3)
                                        Else
                                            deliv = 0
                                        End If
                                    End If
                                    If counter = 0 Then
                                        boundery = 0
                                        ReDim qDocString(0) As String
                                    ElseIf counter Mod 1000 = 0 Then
                                        'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                                        boundery = counter / 1000
                                        ReDim Preserve qDocString(boundery) As String
                                    End If
                                    qDocString(boundery) = qDocString(boundery) & "('" & CDate(c.Offset(0, -5)) & "','" & c.Offset(0, -6) & "'," & CDbl(deliv) & ",'" & Left(currType, 6) & "'),"
                                    counter = counter + 1
                                End If
                                i = i + 1
                            End If
                        End If
                    Next c
                End If
                eRow = c1.Row
                Set c1 = sht.Cells.FindNext(c1)
            Loop While Not c1 Is Nothing And c1.Address <> firstAddress1
            theLastOne = sht.Range("E:E").Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
            For Each c In sht.Range("G" & eRow & ":G" & theLastOne)
                If c.Offset(0, -4) = "Typ dokumentu:" Then
                    currType = c.Offset(0, -3)
                End If
                If c = "PC" Then
                    currZfin = sht.Range("D" & eRow)
                    If c.Offset(0, 1) <> "" Then
                        If counter = 0 Then
                            boundery = 0
                        ElseIf counter Mod 1000 = 0 Then
                            'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                            boundery = counter / 1000
                        End If
                        If c.Offset(0, -6) <> "" Then
                            If c.Offset(0, -3) = "" Then
                                deliv = 0
                            Else
                                deliv = c.Offset(0, -3)
                            End If
                            If counter = 0 Then
                                boundery = 0
                                ReDim qDocString(0) As String
                            ElseIf counter Mod 1000 = 0 Then
                                'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                                boundery = counter / 1000
                                ReDim Preserve qDocString(boundery) As String
                            End If
                            qDocString(boundery) = qDocString(boundery) & "('" & CDate(c.Offset(0, -5)) & "','" & c.Offset(0, -6) & "'," & CDbl(deliv) & ",'" & Left(currType, 6) & "'),"
                        End If
                        i = i + 1
                    End If
                End If
            Next c
        End If
        
        For counter = LBound(qDocString) To UBound(qDocString)
            qDocString(counter) = Left(qDocString(counter), Len(qDocString(counter)) - 1)
        Next counter
        
        
        cSql = "CREATE TABLE #qDocs(qDate datetime, qNumber nchar(9),delNumber bigint, qType nchar(6))"
        AdoConn.Execute cSql
        For counter = LBound(qDocString) To UBound(qDocString)
            iSql = "INSERT INTO #qDocs(qDate,qNumber,delNumber,qType) VALUES " & qDocString(counter)
            AdoConn.Execute iSql
        Next counter
        sSql = "SELECT DISTINCT qDate, qNumber, delNumber, qType FROM #qDocs q WHERE NOT EXISTS (SELECT 1 FROM tbQdocReconciliation qr WHERE qr.qNumber = q.qNumber AND qr.qType = q.qType)"
        iSql = "INSERT INTO tbQdocReconciliation (qDate, qNumber, delNumber, qType) " & sSql
        AdoConn.Execute iSql
        
        downloadQdocs
        
        counter = 0
        
        Set c = sht.Cells.Find("PC", LookIn:=xlValues, LookAt:=xlWhole)
        prevRow = 0
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                If c.Row > prevRow + 1 Then
                    'document changed. Check if it's PW or WZ or other
                    Set c4 = c.Offset(0, -3)
                    Do While c3.Row > prevRow
                        If c4.Offset(0, -1).Value = "Typ dokumentu:" Then
                            theType = Mid(c4.Value, 5, 2)
                            Exit Do
                        End If
                        Set c4 = c4.Offset(-1, 0)
                    Loop
                End If
                If counter = 0 Then
                    boundery = 0
                    ReDim qDocDataStr(0) As String
                ElseIf counter Mod 1000 = 0 Then
                    'we've hit maximum of 1000 elements per single upload. Let's create another string and put the rest there
                    boundery = counter / 1000
                    ReDim Preserve qDocDataStr(boundery) As String
                End If
                qDocDataStr(boundery) = qDocDataStr(boundery) & "(" & Qdocs(theType & "_" & CStr(c.Offset(0, -6))).qId & "," & batches(CStr(CDbl(Right(CStr(c.Offset(0, 1)), 10)))).bId & "," & CLng(c.Offset(0, -1)) & "," & WorksheetFunction.RoundUp(c.Offset(0, 2), 0) & "),"
                counter = counter + 1
                Set c = sht.Cells.FindNext(c)
            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If
        
        For counter = LBound(qDocDataStr) To UBound(qDocDataStr)
            qDocDataStr(counter) = Left(qDocDataStr(counter), Len(qDocDataStr(counter)) - 1)
        Next counter
        
        cSql = "CREATE TABLE #qDocsData(qReconciliationId int, batchId int,batchSize bigint, batchPal int)"
        AdoConn.Execute cSql
        For counter = LBound(qDocDataStr) To UBound(qDocDataStr)
            iSql = "INSERT INTO #qDocsData(qReconciliationId,batchId,batchSize,batchPal) VALUES " & qDocDataStr(counter)
            AdoConn.Execute iSql
        Next counter
    '    iSql = "INSERT INTO #qDocsData(qReconciliationId,batchId,batchSize,batchPal) VALUES " & qDocDataStr
    '    adoConn.Execute iSql
        sSql = "SELECT qReconciliationId,batchId,batchSize,batchPal FROM #qDocsData q WHERE NOT EXISTS (SELECT 1 FROM tbQdocData qr WHERE qr.batchId = q.batchId AND qr.qReconciliationId = q.qReconciliationId)"
        iSql = "INSERT INTO tbQdocData (qReconciliationId,batchId,batchSize,batchPal) " & sSql
        AdoConn.Execute iSql
        
        closeConnection
        eTime = Now
        MsgBox "Upload zakończony powodzeniem w " & Abs(DateDiff("s", eTime, sTime)) & " sek.", vbOKOnly + vbInformation, "Powodzenie"
    End If
Else
    isError = True
    MsgBox "Struktura raportu wydaje się zmieniona. Prawidłowy typ raportu to raport ""Zestawienie obrotów wg artykułów"" z Qguar. Żadne dane nie zostały dodane do bazy.", vbOKOnly + vbCritical, "Błędna struktura raportu"
End If
Exit_here:
closeConnection
Exit Sub

err_trap:
MsgBox "Error in ""exportPW_WZ"". Error number: " & Err.Number & ", " & Err.Description
isError = True
Resume Exit_here

End Sub









