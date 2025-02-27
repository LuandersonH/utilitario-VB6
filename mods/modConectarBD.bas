Attribute VB_Name = "modConectarBD"
Public connectBD As ADODB.Connection
Public recordBD As ADODB.Recordset
Public caminhoBD As String
Public myBD As String

Public Sub InitConexao(frm)
On Error GoTo ErroAoIniciarConexao
     Open App.Path & "\caminhoBD.txt" For Input As #1
          Do While Not EOF(1)
               Line Input #1, caminhoBD
          Loop
     Close #1

     Set connectBD = New ADODB.Connection
     Set recordBD = New ADODB.Recordset

     myBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & caminhoBD

     connectBD.Open myBD

Exit Sub

ErroAoIniciarConexao:

     If connectBD.State = adStateOpen Then
          connectBD.Close
     End If

     Select Case Err.Number
          Case 53
               MsgBox "O arquivo caminhoBD.txt não existe na raiz do programa.", vbExclamation, "Arquivo não encontrado"
          Case -2147217843
               MsgBox "O caminho do banco de dados está incorreto. Verifique o arquivo caminhoBD.txt.", vbExclamation, "Caminho inválido"
          Case Else
               MsgBox "Erro ao iniciar conexão com o Banco de Dados: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
          End Select

     Unload frm

End Sub

