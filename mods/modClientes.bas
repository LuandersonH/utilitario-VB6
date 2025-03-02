Attribute VB_Name = "modClientes"
Public Function AtualizarCamposPorTipo(frm)
     ' Trim vai remover os espa�os em branco
     ' UCase vai deixar os textos mai�sculos, caso seja digitado de outra forma
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

               ' �lbum para Dura��o
               frm.lblDuracao.Caption = "Dura��o"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagDuracao"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 5
                frm.txtDuracaoTemporadasAlbum.Text = ""
               
          Case "SERIE"
               frm.lblNome.Caption = "Nome da s�rie"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretorArtista.Tag = "tagDiretor"
               frm.txtDiretorArtista.Text = ""

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtoresParticipantes.Tag = "tagAtores"
               frm.txtAtoresParticipantes.Text = ""

               ' Dura��o para Temporadas
               frm.lblDuracao.Caption = "Temporadas"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagTemporadas"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 2
               frm.txtDuracaoTemporadasAlbum.Text = ""

          Case "MUSICA"
               frm.lblNome.Caption = "Nome da m�sica"

               ' Diretor para Artista
               frm.lblDiretor.Caption = "Artista"
               frm.txtDiretorArtista.Tag = "tagArtista"
               frm.txtDiretorArtista.Text = ""

               ' Atores para Participantes
               frm.lblAtores.Caption = "Participantes"
               frm.txtAtoresParticipantes.Tag = "tagParticipantes"
               frm.txtAtoresParticipantes.Text = ""

               ' Dura��o para �lbum
               frm.lblDuracao.Caption = "�lbum"
               frm.txtDuracaoTemporadasAlbum.Tag = "tagAlbum"
               frm.txtDuracaoTemporadasAlbum.MaxLength = 255
               frm.txtDuracaoTemporadasAlbum.Text = ""
     End Select
End Function

