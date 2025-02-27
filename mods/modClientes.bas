Attribute VB_Name = "modClientes"
Public Function AtualizarCamposPorTipo(frm)
     ' Trim vai remover os espaços em branco
     ' UCase vai deixar os textos maiúsculos, caso seja digitado de outra forma
     Select Case Trim(UCase(frm.cboTipo.Text))
          Case "FILME"
               frm.lblNome.Caption = "Nome do filme"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"

               ' Álbum para Duração
               frm.lblDuracao.Caption = "Duração"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagDuracao"
               
          Case "SERIE"
               frm.lblNome.Caption = "Nome da série"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"

               ' Duração para Temporadas
               frm.lblDuracao.Caption = "Temporadas"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagTemporadas"

          Case "MUSICA"
               frm.lblNome.Caption = "Nome da música"

               ' Diretor para Artista
               frm.lblDiretor.Caption = "Artista"
               frm.txtDiretorArtista.Tag = "tagArtista"

               ' Atores para Participantes
               frm.lblAtores.Caption = "Participantes"
               frm.txtAtoresParticipantes.Tag = "tagParticipantes"

               ' Duração para Álbum
               frm.lblDuracao.Caption = "Álbum"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagAlbum"
               
     End Select
End Function

