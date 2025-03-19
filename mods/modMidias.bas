Attribute VB_Name = "modMidias"
Public Function UnionFilmesSeriesMusicas()
 'campos  totais no BD, apos Union All: CODIGO - NOME - DIRETOR - ATORES - TEMPORADAS - GENERO - NOTA - OBSERVA��O - ARTISTA - PARTICIPANTES - ALBUM - DURACAO - TIPO
UnionFilmesSeriesMusicas = "SELECT Codigo, Nome, Diretor, Atores, 0 AS Temporadas, Genero, Nota, Observacao, Null AS Artista, Null AS Participantes, Null AS Album, Duracao, Grupo, Excluido FROM Filmes " & _
     "UNION ALL " & _
     "SELECT Codigo, Nome, Diretor, Atores, Temporadas, Genero, Nota, Observacao, Null AS Artista, Null AS Participantes, Null AS Album, Null AS Duracao, Grupo, Excluido FROM Series " & _
     "UNION ALL " & _
      "SELECT Codigo, Nome, Null AS Diretor, Null AS Atores, Null AS Temporadas, Genero, Nota, Observacao, Artista, Participantes, Album, Null AS Duracao, Grupo, Excluido FROM Musicas"

End Function

Public Function setarColunasIniciaisDoGridMedia(frm)
     With frm.GridMedia
               .Clear
               .Cols = 13
               .Rows = 1
               .TextMatrix(0, 0) = "Codigo"
               .TextMatrix(0, 1) = "Nome"
               .TextMatrix(0, 2) = "Diretor"
               .TextMatrix(0, 3) = "Atores"
               .TextMatrix(0, 4) = "Temporadas"
               .TextMatrix(0, 5) = "Genero"
               .TextMatrix(0, 6) = "Nota"
               .TextMatrix(0, 7) = "Observacao"
               .TextMatrix(0, 8) = "Artista"
               .TextMatrix(0, 9) = "Participantes"
               .TextMatrix(0, 10) = "Album"
               .TextMatrix(0, 11) = "Duracao"
               .TextMatrix(0, 12) = "Grupo"

               .ColWidth(0) = frm.Width / 12
               .ColWidth(1) = frm.Width / 12
               .ColWidth(2) = frm.Width / 12
               .ColWidth(3) = frm.Width / 12
               .ColWidth(4) = frm.Width / 12
               .ColWidth(5) = frm.Width / 12
               .ColWidth(6) = frm.Width / 12
               .ColWidth(7) = frm.Width / 12
               .ColWidth(8) = frm.Width / 12
               .ColWidth(9) = frm.Width / 12
               .ColWidth(10) = frm.Width / 12
               .ColWidth(11) = frm.Width / 12
     End With
End Function

Public Function inserirDadosDoRecordSetNoGridMedia(frm)
     Dim linhaAtualMedia As Integer
     linhaAtualMedia = 1

     While Not recordBD.EOF
          With frm.GridMedia
               

               If recordBD!Excluido <> 1 Then
                    .Rows = frm.GridMedia.Rows + 1

                    .TextMatrix(linhaAtualMedia, 0) = IIf(IsNull(recordBD!Codigo), 0, recordBD!Codigo)
                    .TextMatrix(linhaAtualMedia, 1) = IIf(IsNull(recordBD!Nome), "", recordBD!Nome)
                    .TextMatrix(linhaAtualMedia, 2) = IIf(IsNull(recordBD!Diretor), "", recordBD!Diretor)
                    .TextMatrix(linhaAtualMedia, 3) = IIf(IsNull(recordBD!Atores), "", recordBD!Atores)
                    .TextMatrix(linhaAtualMedia, 4) = IIf(IsNull(recordBD!Temporadas), 0, recordBD!Temporadas)
                    .TextMatrix(linhaAtualMedia, 5) = IIf(IsNull(recordBD!Genero), "", recordBD!Genero)
                    .TextMatrix(linhaAtualMedia, 6) = IIf(IsNull(recordBD!Nota), 0, recordBD!Nota)
                    .TextMatrix(linhaAtualMedia, 7) = IIf(IsNull(recordBD!Observacao), "", recordBD!Observacao)
                    .TextMatrix(linhaAtualMedia, 8) = IIf(IsNull(recordBD!Artista), "", recordBD!Artista)
                    .TextMatrix(linhaAtualMedia, 9) = IIf(IsNull(recordBD!Participantes), "", recordBD!Participantes)
                    .TextMatrix(linhaAtualMedia, 10) = IIf(IsNull(recordBD!Album), "", recordBD!Album)
                    .TextMatrix(linhaAtualMedia, 11) = IIf(IsNull(recordBD!Duracao), "", recordBD!Duracao)
                    .TextMatrix(linhaAtualMedia, 12) = IIf(IsNull(recordBD!Grupo), "", recordBD!Grupo)
     
                    linhaAtualMedia = linhaAtualMedia + 1
               End If

               recordBD.MoveNext

          End With
      Wend
End Function

Public Function CarregarTodasAsMedias(frm)
On Error GoTo erroAoCarregarMidias
          
     If connectBD.State = adStateClosed Then connectBD.Open

     If recordBD.State = adStateOpen Then recordBD.Close
          recordBD.Open UnionFilmesSeriesMusicas, connectBD, adOpenStatic, adLockReadOnly

     Call setarColunasIniciaisDoGridMedia(frm)
     Call inserirDadosDoRecordSetNoGridMedia(frm)

     recordBD.Close
     Exit Function

erroAoCarregarMidias:
MsgBox "Erro ao carregar midias: " & Err.Number & " - " & Err.Description, vbCritical, "ERRO!"
 If recordBD.State = adStateOpen Then recordBD.Close
End Function

Public Function pesquisarNoInputMediaFilterComLike(frm)
     Dim textoDoInputMedia As String
     Dim queryInputMediaFilter As String
     Dim queryUnion As String

     'query parametrizada
      Dim cmdInputMedia As New ADODB.Command
      Set cmdInputMedia = New ADODB.Command

     'conecta ao BD
      If connectBD.State = adStateClosed Then connectBD.Open
     'conect o commnd ao bd
     cmdInputMedia.ActiveConnection = connectBD
     
     queryUnion = UnionFilmesSeriesMusicas
     cmdInputMedia.CommandText = "SELECT * FROM (" & queryUnion & ") WHERE Nome LIKE ?"
   
     cmdInputMedia.Parameters.Append cmdInputMedia.CreateParameter(, adVarChar, adParamInput, 255, "%" & frm.inputMediaFilter.Text & "%")

     If recordBD.State = adStateOpen Then recordBD.Close
     Set recordBD = cmdInputMedia.Execute

      Call setarColunasIniciaisDoGridMedia(frm)
     Call inserirDadosDoRecordSetNoGridMedia(frm)

     recordBD.Close
End Function


'frm CADASTRO DE MIDIAS ABAIXO
'frm CADASTRO DE MIDIAS ABAIXO
'frm CADASTRO DE MIDIAS ABAIXO

Public Function removerEspacosEmBranco(frm)
    frm.txtNome.Text = Trim(frm.txtNome.Text)
    frm.txtAtoresParticipantes.Text = Trim(frm.txtAtoresParticipantes.Text)
    frm.txtDuracaoTemporadasAlbum.Text = Trim(frm.txtDuracaoTemporadasAlbum.Text)
    frm.txtGenero.Text = Trim(frm.txtGenero.Text)
    frm.txtDiretorArtista.Text = Trim(frm.txtDiretorArtista.Text)
    frm.txtObservacao.Text = Trim(frm.txtObservacao.Text)
End Function




Public Function AtualizarCamposPorTipo(frm)
     ' Trim vai remover os espacos em branco
     ' UCase vai deixar os textos maiusculos, caso seja digitado de outra forma
     Select Case Trim(UCase(frm.cboTipo.Text))
          Case "FILME"
               frm.lblNome.Caption = "Nome do filme"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"
               frm.txtDiretorArtista.Text = ""

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"
               frm.txtAtoresParticipantes.Text = ""

               ' album para Duracao
               frm.lblDuracao.Caption = "Duracao"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagDuracao"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 5
                frm.txtDuracaoTemporadasAlbum.Text = ""
               
          Case "SERIE"
               frm.lblNome.Caption = "Nome da serie"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"
               frm.txtDiretorArtista.Text = ""

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"
               frm.txtAtoresParticipantes.Text = ""

               ' Duracao para Temporadas
               frm.lblDuracao.Caption = "Temporadas"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagTemporadas"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 2
               frm.txtDuracaoTemporadasAlbum.Text = ""

          Case "MUSICA"
               frm.lblNome.Caption = "Nome da musica"

               ' Diretor para Artista
               frm.lblDiretor.Caption = "Artista"
               frm.txtDiretorArtista.Tag = "tagArtista"
               frm.txtDiretorArtista.Text = ""

               ' Atores para Participantes
               frm.lblAtores.Caption = "Participantes"
               frm.txtAtoresParticipantes.Tag = "tagParticipantes"
               frm.txtAtoresParticipantes.Text = ""

               ' Duracao para album
               frm.lblDuracao.Caption = "�lbum"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagAlbum"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 255
               frm.txtDuracaoTemporadasAlbum.Text = ""
     End Select
End Function

'frm EXCLUIDOS ABAIXO
'frm EXCLUIDOS ABAIXO
'frm EXCLUIDOS ABAIXO

Public Function inserirDadosExcluidosDoRecordSetNoGridMedia(frm)
     Dim linhaAtualMedia As Integer
     linhaAtualMedia = 1

     While Not recordBD.EOF
          With frm.GridMedia
               
               If recordBD!Excluido <> 0 Then
                    .Rows = frm.GridMedia.Rows + 1

                    .TextMatrix(linhaAtualMedia, 0) = IIf(IsNull(recordBD!Codigo), 0, recordBD!Codigo)
                    .TextMatrix(linhaAtualMedia, 1) = IIf(IsNull(recordBD!Nome), "", recordBD!Nome)
                    .TextMatrix(linhaAtualMedia, 2) = IIf(IsNull(recordBD!Diretor), "", recordBD!Diretor)
                    .TextMatrix(linhaAtualMedia, 3) = IIf(IsNull(recordBD!Atores), "", recordBD!Atores)
                    .TextMatrix(linhaAtualMedia, 4) = IIf(IsNull(recordBD!Temporadas), 0, recordBD!Temporadas)
                    .TextMatrix(linhaAtualMedia, 5) = IIf(IsNull(recordBD!Genero), "", recordBD!Genero)
                    .TextMatrix(linhaAtualMedia, 6) = IIf(IsNull(recordBD!Nota), 0, recordBD!Nota)
                    .TextMatrix(linhaAtualMedia, 7) = IIf(IsNull(recordBD!Observacao), "", recordBD!Observacao)
                    .TextMatrix(linhaAtualMedia, 8) = IIf(IsNull(recordBD!Artista), "", recordBD!Artista)
                    .TextMatrix(linhaAtualMedia, 9) = IIf(IsNull(recordBD!Participantes), "", recordBD!Participantes)
                    .TextMatrix(linhaAtualMedia, 10) = IIf(IsNull(recordBD!Album), "", recordBD!Album)
                    .TextMatrix(linhaAtualMedia, 11) = IIf(IsNull(recordBD!Duracao), "", recordBD!Duracao)
                    .TextMatrix(linhaAtualMedia, 12) = IIf(IsNull(recordBD!Grupo), "", recordBD!Grupo)
     
                    linhaAtualMedia = linhaAtualMedia + 1
               End If

               recordBD.MoveNext

          End With
      Wend
End Function

Public Function pesquisarExcluidosNoInputMediaFilterComLike(frm)
     Dim textoDoInputMedia As String
     Dim queryInputMediaFilter As String
     Dim queryUnion As String

     'query parametrizada
      Dim cmdInputMedia As New ADODB.Command
      Set cmdInputMedia = New ADODB.Command

     'conecta ao BD
      If connectBD.State = adStateClosed Then connectBD.Open
     'conect o commnd ao bd
     cmdInputMedia.ActiveConnection = connectBD
     
     queryUnion = UnionFilmesSeriesMusicas
     cmdInputMedia.CommandText = "SELECT * FROM (" & queryUnion & ") WHERE Nome LIKE ?"
   
     cmdInputMedia.Parameters.Append cmdInputMedia.CreateParameter(, adVarChar, adParamInput, 255, "%" & frm.inputMediaFilter.Text & "%")

     If recordBD.State = adStateOpen Then recordBD.Close
     Set recordBD = cmdInputMedia.Execute

     Call setarColunasIniciaisDoGridMedia(frm)
     Call inserirDadosExcluidosDoRecordSetNoGridMedia(frm)

     recordBD.Close
End Function

Public Function CarregarTodasAsMediasExcluidas(frm)
On Error GoTo erroAoCarregarMidiasExcluidas
          
     If connectBD.State = adStateClosed Then connectBD.Open

     If recordBD.State = adStateOpen Then recordBD.Close
          recordBD.Open UnionFilmesSeriesMusicas, connectBD, adOpenStatic, adLockReadOnly

     Call setarColunasIniciaisDoGridMedia(frm)
     Call inserirDadosExcluidosDoRecordSetNoGridMedia(frm)

     recordBD.Close
     Exit Function

erroAoCarregarMidiasExcluidas:
MsgBox "Erro ao carregar midias: " & Err.Number & " - " & Err.Description, vbCritical, "ERRO!"
 If recordBD.State = adStateOpen Then recordBD.Close
End Function
