Attribute VB_Name = "modConectarBD"
Public connectBD As ADODB.Connection
Public recordBD As ADODB.Recordset
Public myBD As String

Public Sub InitConexao()

     Set connectBD = New ADODB.Connection
     Set recordBD = New ADODB.Recordset

     myBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\bd_utilitarios.mdb"

     connectBD.Open myBD

     If connectBD.State = adStateOpen Then
          Else
          MsgBox "Não foi possível conectar ao banco de dados", vbCritical, "Erro"
     End If

End Sub
