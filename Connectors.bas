Attribute VB_Name = "Connectors"

Public Sub connectScada()
'Dim cmd As ADODB.Command
'Set cmd = New ADODB.Command

If ScadaConn Is Nothing Then
    Set ScadaConn = New ADODB.Connection
    ScadaConn.Provider = "SQLOLEDB"
    ScadaConn.ConnectionString = ScadaConnectionString
    ScadaConn.Open
    ScadaConn.CommandTimeout = 90
Else
    If ScadaConn.State = adStateClosed Then
        Set ScadaConn = New ADODB.Connection
        ScadaConn.Provider = "SQLOLEDB"
        ScadaConn.ConnectionString = ScadaConnectionString
        ScadaConn.Open
        ScadaConn.CommandTimeout = 90
    End If
End If


End Sub

Public Sub disconnectScada()

If Not ScadaConn Is Nothing Then
    If ScadaConn.State = 1 Then
        ScadaConn.Close
    End If
    Set ScadaConn = Nothing
End If
End Sub

Sub updateConnection(Optional timeout As Variant)
Dim tm As Integer

If Not IsMissing(timeout) Then
    tm = timeout
Else
    tm = 1200
End If

If Not AdoConn Is Nothing Then
    If AdoConn.State = 0 Then
        AdoConn.Open ConnectionString
        AdoConn.CommandTimeout = tm
    End If
Else
    Set AdoConn = New ADODB.Connection
    AdoConn.Open ConnectionString
    AdoConn.CommandTimeout = tm
End If
End Sub

Sub closeConnection()

If Not AdoConn Is Nothing Then
    If AdoConn.State = 1 Then
        AdoConn.Close
    End If
    Set AdoConn = Nothing
End If
End Sub
