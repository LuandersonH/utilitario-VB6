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
               frm.txtDiretorArtista.Text = ""

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"
               frm.txtAtoresParticipantes.Text = ""

               ' Álbum para Duração
               frm.lblDuracao.Caption = "Duração"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagDuracao"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 5
                frm.txtDuracaoTemporadasAlbum.Text = ""
               
          Case "SERIE"
               frm.lblNome.Caption = "Nome da série"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"
               frm.txtDiretorArtista.Text = ""

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"
               frm.txtAtoresParticipantes.Text = ""

               ' Duração para Temporadas
               frm.lblDuracao.Caption = "Temporadas"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagTemporadas"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 2
               frm.txtDuracaoTemporadasAlbum.Text = ""

          Case "MUSICA"
               frm.lblNome.Caption = "Nome da música"

               ' Diretor para Artista
               frm.lblDiretor.Caption = "Artista"
               frm.txtDiretorArtista.Tag = "tagArtista"
               frm.txtDiretorArtista.Text = ""

               ' Atores para Participantes
               frm.lblAtores.Caption = "Participantes"
               frm.txtAtoresParticipantes.Tag = "tagParticipantes"
               frm.txtAtoresParticipantes.Text = ""

               ' Duração para Álbum
               frm.lblDuracao.Caption = "Álbum"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagAlbum"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 255
               frm.txtDuracaoTemporadasAlbum.Text = ""
     End Select
End Function

