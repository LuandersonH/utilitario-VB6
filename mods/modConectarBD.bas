Attribute VB_Name = "modConectarBD"
Public connectBD As ADODB.Connection
Public recordBD As ADODB.Recordset
Public caminhoBD As String
Public MyBD As String
Public fecharDireto As Boolean

Public Sub InitConexao(frm)
On Error GoTo ErroAoIniciarConexao
Dim caminhoBD As String
Dim MyBD As String

If Dir(App.Path & "\caminhoBD.txt") <> "" Then

     Open App.Path & "\caminhoBD.txt" For Input As #1
          If Not EOF(1) Then
               Line Input #1, caminhoBD
          End If
     Close #1

     If Trim(caminhoBD) = "" Then
          MsgBox "O arquivo 'caminhoBD.txt' da raiz de Utilitarios est� vazio, preencha com o caminho do banco de dados, exemplo: C:\apps\Utilitarios\BD_Utilitarios.mdb", vbExclamation, "Arquivo Inv�lido"
          fecharDireto = 1
           Unload frm
          Exit Sub
     End If

     Set connectBD = New ADODB.Connection
     Set recordBD = New ADODB.Recordset

     MyBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & caminhoBD

     connectBD.Open MyBD
Else
     MsgBox "Crie o arquivo 'caminhoBD.txt' da raiz de Utilitarios com o caminho do banco de dados, exemplo: C:\apps\Utilitarios\BD_Utilitarios.mdb"
     fecharDireto = 1
     Unload frm
     Set frm = Nothing
     Exit Sub
End If

Exit Sub

ErroAoIniciarConexao:

     Close #1

     If connectBD.State = adStateOpen Then
          connectBD.Close
     End If

     Select Case Err.Number
          Case 53
               MsgBox "O arquivo caminhoBD.txt n�o existe na raiz do programa.", vbExclamation, "Arquivo n�o encontrado"
          Case -2147217843
               MsgBox "O caminho do banco de dados est� incorreto. Verifique o arquivo caminhoBD.txt.", vbExclamation, "Caminho inv�lido"
          Case -2147467259
               MsgBox "O caminho do banco de dados em 'caminhoBD.txt' est� incorreto.Verifique e tente novamente.", vbExclamation, "Caminho inv�lido"
          Case Else
               MsgBox "Erro ao iniciar conex�o com o Banco de Dados: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
          End Select

           fecharDireto = 1
           Unload frm

End Sub

