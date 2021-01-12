Attribute VB_Name = "CustomFunctions"
Public Function PROD_WAGA_PALETY(prodNumber As Long) As Double
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.unitWeight, tbUom.pcPerPallet FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & prodNumber & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_WAGA_PALETY = Round(rs.Fields("unitWeight") * rs.Fields("pcPerPallet"), 2)
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

End Function

Public Function PROD_WAGA(prodNumber As Long) As Double
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.unitWeight FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & prodNumber & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_WAGA = rs.Fields("unitWeight")
End If
rs.Close

exit_here:
closeConnection
Set rs = Nothing



End Function

Public Function PROD_NAZWA(prodNumber As Long) As String
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT zfinName FROM tbZfin WHERE zfinIndex = " & prodNumber & ";"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_NAZWA = rs.Fields("zfinName")
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If



End Function

Public Function CS_HELP()
MsgBox "PROD_WAGA - zwraca wagę produktu w KG. Jako parametr przyjmuje index produktu." _
        & vbNewLine & "PROD_NAZWA - zwraca nazwę produktu. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_WAGA_PALETY - zwraca wagę netto palety produktu. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_SZT_NA_PALECIE - zwraca ilość sztuk na palecie. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_SZT_W_KARTONIE - zwraca ilość sztuk w kartonie. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "OSTATNIA_PELNA - zwraca adres ostatniej niepustej komórki w zakresie podanym jako parametr. " _
        & vbNewLine & "PROD_WARSTW_NA_PALECIE - zwraca ilość warstw kartonów na palecie. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_CZY_CHEP - zwraca PRAWDA jeśli produkt jest na palecie CHEP. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_CZY_EURO - zwraca PRAWDA jeśli produkt jest na palecie EURO. Jako parametr przyjmuje index produktu. " _
        & vbNewLine & "PROD_INDEX_FOLII - zwraca index folii dla produktu wybranego jako parametr." _
        & vbNewLine & "PROD_INDEX_KAWY - zwraca index ZFORa dla produktu wybranego jako parametr. " _
        & vbNewLine & "PROD_LINIA - zwraca linie produkcyjne, na których może być produkowany produkt podany jako parametr."

End Function

Public Function PROD_SZT_NA_PALECIE(prodNumber As Long) As Long
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.pcPerPallet FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & prodNumber & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_SZT_NA_PALECIE = rs.Fields("pcPerPallet")
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_SZT_W_KARTONIE(prodNumber As Long) As Integer
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.pcPerBox FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & prodNumber & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_SZT_W_KARTONIE = rs.Fields("pcPerBox")
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_WARSTW_NA_PALECIE(prodNumber As Long) As Integer
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.pcPerPallet, tbUom.pcLayer FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & prodNumber & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_WARSTW_NA_PALECIE = rs.Fields("pcPerPallet") / rs.Fields("pcLayer")
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_CZY_CHEP(prodNumber As Long) As Variant
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT palletChep FROM tbZfin LEFT JOIN (tbUom LEFT JOIN tbPallets ON tbUom.palletType = tbPallets.palletId) ON tbUom.zfinId = tbZfin.zfinId WHERE tbZfin.zfinIndex = " & prodNumber
Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_CZY_CHEP = rs.Fields("palletChep")
Else
    PROD_CZY_CHEP = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_CZY_EURO(prodNumber As Long) As Variant
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT palletLength+palletWidth as totalDim FROM tbZfin LEFT JOIN (tbUom LEFT JOIN tbPallets ON tbUom.palletType = tbPallets.palletId) ON tbUom.zfinId = tbZfin.zfinId WHERE tbZfin.zfinIndex = " & prodNumber



Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields("totalDim") = 2000 Then
        PROD_CZY_EURO = True
    Else
        PROD_CZY_EURO = False
    End If
Else
    PROD_CZY_EURO = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_SESJA_ZLECENIA(ordNumber As Long) As Variant
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT DISTINCT op.SessionNumber " _
    & "FROM tbOrders ord JOIN tbOperations op ON op.orderId=ord.orderId " _
    & "WHERE ord.sapId = " & ordNumber

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    If IsNull(rs.Fields("SessionNumber")) Then
        PROD_SESJA_ZLECENIA = "B/D"
    Else
        PROD_SESJA_ZLECENIA = rs.Fields("SessionNumber")
        rs.MoveNext
        If Not rs.EOF Then
            PROD_SESJA_ZLECENIA = "MULTI"
        End If
    End If
Else
    PROD_SESJA_ZLECENIA = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_INDEX_KAWY(prodNumber As Long) As Variant
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT zfin.zfinIndex FROM tbZfin LEFT JOIN tbZfinZfor on tbZfin.zfinId = tbZfinZfor.zfinId LEFT JOIN tbZfin zfin on zfin.zfinId=tbzfinzfor.zforId WHERE tbZfin.zfinIndex = " & prodNumber

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_INDEX_KAWY = rs.Fields("zfinIndex")
Else
    PROD_INDEX_KAWY = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If

End Function


Public Function PROD_KONWERSJA(product As Long, amount As Double, startUnit As String, endUnit) As Variant
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String
Dim denominator As Double

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbZfin.zfinIndex, tbUom.pcPerPallet, tbUom.pcPerBox, tbUom.unitWeight FROM tbUom RIGHT JOIN tbZfin ON tbUom.zfinId = tbZfin.zfinId WHERE (((tbZfin.zfinIndex)=" & product & "));"

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    Select Case startUnit
    Case Is = "pc"
        Select Case endUnit
        Case Is = "pc"
            PROD_KONWERSJA = amount
        Case Is = "kg"
            PROD_KONWERSJA = amount * rs.Fields("unitWeight")
        Case Is = "box"
            PROD_KONWERSJA = amount / rs.Fields("pcPerBox")
        Case Is = "pal"
            PROD_KONWERSJA = amount / rs.Fields("pcPerPallet")
        Case Else
            PROD_KONWERSJA = "B/D"
        End Select
    Case Is = "kg"
        Select Case endUnit
        Case Is = "pc"
            PROD_KONWERSJA = amount * (1 / rs.Fields("unitWeight"))
        Case Is = "kg"
            PROD_KONWERSJA = amount
        Case Is = "box"
            PROD_KONWERSJA = (amount * (1 / rs.Fields("unitWeight")) / rs.Fields("pcPerBox"))
        Case Is = "pal"
            PROD_KONWERSJA = (amount * (1 / rs.Fields("unitWeight")) / rs.Fields("pcPerPallet"))
        Case Else
            PROD_KONWERSJA = "B/D"
        End Select
    Case Is = "box"
        Select Case endUnit
        Case Is = "pc"
            PROD_KONWERSJA = amount * rs.Fields("pcPerBox")
        Case Is = "kg"
            PROD_KONWERSJA = amount * rs.Fields("pcPerBox") * rs.Fields("unitWeight")
        Case Is = "box"
            PROD_KONWERSJA = amount
        Case Is = "pal"
            PROD_KONWERSJA = (amount * rs.Fields("pcPerBox")) / rs.Fields("pcPerPallet")
        Case Else
            PROD_KONWERSJA = "B/D"
        End Select
    Case Is = "pal"
        Select Case endUnit
        Case Is = "pc"
            PROD_KONWERSJA = amount * rs.Fields("pcPerPallet")
        Case Is = "kg"
            PROD_KONWERSJA = amount * rs.Fields("pcPerPallet") * rs.Fields("unitWeight")
        Case Is = "box"
            PROD_KONWERSJA = (amount * rs.Fields("pcPerPallet") / rs.Fields("pcPerBox"))
        Case Is = "pal"
            PROD_KONWERSJA = amount
        Case Else
            PROD_KONWERSJA = "B/D"
        End Select
    Case Else
        PROD_KONWERSJA = "B/D"
    End Select
End If
rs.Close


exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If
End Function

Sub DisplayedToActual()
Dim rng As Range
Set rng = Application.ActiveCell
rng.Value = rng.text
rng.Value = CDbl(rng.Value)
End Sub

Public Function PROD_KLIENT(prodNumber As Long) As String
Dim oConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
' Make a query over the connection.
sSql = "SELECT tbCustomerString.custString FROM tbCustomerString RIGHT JOIN tbZFIN ON tbCustomerString.custStringId = tbZFIN.custString WHERE tbZFIN.zfinIndex = " & prodNumber & ";"
        

Set rs = New ADODB.Recordset
rs.Open sSql, AdoConn, adOpenStatic, adLockBatchOptimistic, adCmdText
If Not rs.EOF Then
    rs.MoveFirst
    PROD_KLIENT = rs.Fields("custString")
Else
    PROD_KLIENT = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If



End Function

Public Function PROD_CZY_ZIARNO(prodNumber As Long) As Variant
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
sSql = "SELECT tbzfin.zfinIndex, tbZfinProperties.[beans?] FROM tbZfinProperties JOIN tbZfin on tbZfin.zfinId = tbZfinProperties.zfinId WHERE tbZfin.zfinIndex = " & prodNumber & ";"
Set rs = AdoConn.Execute(sSql)

If Not rs.EOF Then
    rs.MoveFirst
    If rs.Fields("beans?") = 0 Then
        PROD_CZY_ZIARNO = False
    Else
        PROD_CZY_ZIARNO = True
    End If
Else
    PROD_CZY_ZIARNO = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

Public Function PROD_INDEX_FOLII(prodNumber As Long) As Variant
Dim rs As ADODB.Recordset
Dim sConn As String
Dim sSql As String
Dim dbPath As String

On Error GoTo exit_here

updateConnection
sSql = "SELECT bomy.*, mat.zfinIndex, freshBom.dateAdded FROM tbBom bomy RIGHT JOIN " _
    & "(SELECT oBom.zfinId,  MAX(oBom.bomRecId) as bomRecId, MAX(oBom.dateAdded) as dateAdded FROM " _
    & "(SELECT iBom.bomRecId, zfinId, br.dateAdded FROM tbBomReconciliation br JOIN ( " _
    & "SELECT bomRecId, bom.zfinId FROM tbBom bom GROUP BY bomRecId, bom.zfinId) iBom ON iBom.bomRecId=br.bomRecId) oBom " _
    & "WHERE oBom.dateAdded <='" & Now & "' GROUP BY oBom.zfinId) freshBom ON freshBom.zfinId=bomy.zfinId AND freshBom.bomRecId=bomy.bomRecId " _
    & "LEFT JOIN tbZfin zfin ON zfin.zfinId=bomy.zfinId LEFT JOIN tbZfin mat ON mat.zfinId=bomy.materialId " _
    & "WHERE mat.materialType=2 AND zfin.zfinIndex=" & prodNumber & ";"
Set rs = AdoConn.Execute(sSql)

If Not rs.EOF Then
    rs.MoveFirst
    PROD_INDEX_FOLII = rs.Fields("zfinIndex")
Else
    PROD_INDEX_FOLII = "B/D"
End If
rs.Close

exit_here:
closeConnection
If Not rs Is Nothing Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
End If


End Function

