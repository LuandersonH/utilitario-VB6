Attribute VB_Name = "modClientes"
Public Function AtualizarCamposPorTipo(frm)
     ' Trim vai remover os espa�os em branco
     ' UCase vai deixar os textos mai�sculos, caso seja digitado de outra forma
     Select Case Trim(UCase(frm.cboTipo.Text))
          Case "FILME"
               frm.lblNome.Caption = "Nome do filme"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretor.Tag = "tagDiretor"

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtores.Tag = "tagAtores"

               ' �lbum para Dura��o
               frm.lblDuracao.Caption = "Dura��o"
               frm.txtDuracao.Tag = "tagDuracao"
               
          Case "SERIE"
               frm.lblNome.Caption = "Nome da s�rie"

               ' Artista para Diretor
               frm.lblDiretor.Caption = "Diretor"
               frm.txtDiretor.Tag = "tagDiretor"

               ' Participantes para Atores
               frm.lblAtores.Caption = "Atores"
               frm.txtAtores.Tag = "tagAtores"

               ' Dura��o para Temporadas
               frm.lblDuracao.Caption = "Temporadas"
               frm.txtDuracao.Tag = "tagTemporadas"

          Case "MUSICA"
               frm.lblNome.Caption = "Nome da m�sica"

               ' Diretor para Artista
               frm.lblDiretor.Caption = "Artista"
               frm.txtDiretor.Tag = "tagArtista"

               ' Atores para Participantes
               frm.lblAtores.Caption = "Participantes"
               frm.txtAtores.Tag = "tagParticipantes"

               ' Dura��o para �lbum
               frm.lblDuracao.Caption = "�lbum"
               frm.txtDuracao.Tag = "tagAlbum"
               
     End Select
End Function

