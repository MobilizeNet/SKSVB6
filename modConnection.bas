Attribute VB_Name = "modConnection"
Option Explicit

Public CurrentConnection As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs2 As New ADODB.Recordset

Public Sub OpenConnection()
Set CurrentConnection = New ADODB.Connection
CurrentConnection.Open ConnectionString
End Sub

Public Sub ExecuteSql(Statement As String)
Set rs = New ADODB.Recordset
rs.Open Statement, CurrentConnection, adOpenKeyset, adLockPessimistic
End Sub

Public Sub ExecuteSql2(Statement As String)
Set rs2 = New ADODB.Recordset
rs2.Open Statement, CurrentConnection, adOpenKeyset, adLockPessimistic
End Sub





