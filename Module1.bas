Attribute VB_Name = "Module1"
Global iniid As String
Global inipass As String
Global ini_no_transaksi As String

Public koneksi As New ADODB.Connection
Public admin As ADODB.Recordset

Public Sub BukaDB()
Set koneksi = New ADODB.Connection
Set admin = New ADODB.Recordset
koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\rental.mdb"
End Sub

