Attribute VB_Name = "Conex�o"
Option Explicit

Public cnn As New ADODB.Connection

Public Rs As New ADODB.Recordset

Public Sql As String

Sub ConexaoBd()

cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\Db.accdb;Jet OLEDB:Database Password=MyDbPassword;"



End Sub

