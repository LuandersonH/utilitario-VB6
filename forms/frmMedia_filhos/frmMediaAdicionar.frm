VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMidia_Cadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin lvButton.lvButtons_H lvCadastroVoltar 
      Height          =   1005
      Left            =   5010
      TabIndex        =   17
      Top             =   2940
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "VOLTAR"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMediaAdicionar.frx":0000
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H lvCadastroAdicionar 
      Height          =   1005
      Left            =   2340
      TabIndex        =   16
      Top             =   2940
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "ADICIONAR"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483625
      cFHover         =   -2147483625
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMediaAdicionar.frx":25DA
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin VB.TextBox txtDuracaoTemporadasAlbum 
      Alignment       =   2  'Center
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   45
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "tagDuracao"
      Top             =   1800
      Width           =   1185
   End
   Begin VB.ComboBox cboNota 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      ItemData        =   "frmMediaAdicionar.frx":4BB4
      Left            =   7260
      List            =   "frmMediaAdicionar.frx":4BB6
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "tagNota"
      Top             =   1800
      Width           =   1830
   End
   Begin VB.TextBox txtObservacao 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Tag             =   "tagObservacao"
      Top             =   2535
      Width           =   9645
   End
   Begin VB.TextBox txtGenero 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Tag             =   "tagGenero"
      Top             =   1800
      Width           =   5565
   End
   Begin VB.TextBox txtAtoresParticipantes 
      Height          =   315
      Left            =   4000
      TabIndex        =   4
      Tag             =   "tagAtores"
      Top             =   1095
      Width           =   5685
   End
   Begin VB.TextBox txtDiretorArtista 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Tag             =   "tagDiretor"
      Top             =   1110
      Width           =   3705
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   2445
      MaxLength       =   80
      TabIndex        =   2
      Tag             =   "tagNome"
      Top             =   405
      Width           =   7245
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmMediaAdicionar.frx":4BB8
      Left            =   105
      List            =   "frmMediaAdicionar.frx":4BC5
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "tagTipo"
      Top             =   400
      Width           =   1830
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observacao"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4110
      TabIndex        =   15
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Label lblNota 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7335
      TabIndex        =   14
      Top             =   1600
      Width           =   810
   End
   Begin VB.Label lblGenero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1545
      TabIndex        =   13
      Top             =   1605
      Width           =   690
   End
   Begin VB.Label lblDuracao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duracao"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   100
      TabIndex        =   12
      Top             =   1600
      Width           =   810
   End
   Begin VB.Label lblAtores 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4035
      TabIndex        =   11
      Top             =   870
      Width           =   1065
   End
   Begin VB.Label lblDiretor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diretor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   10
      Top             =   885
      Width           =   1065
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2490
      TabIndex        =   9
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   420
   End
End
Attribute VB_Name = "frmMidia_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnVoltar_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMidia.Show
Unload Me

End Sub

Private Sub cboTipo_Click()
     Call AtualizarCamposPorTipo(Me)
End Sub

Private Sub Form_Load()
     'Adicionando as opcoes no cboNota - notas de 1 a 5
    Dim i As Integer
     For i = 1 To 5
          cboNota.AddItem i
     Next i

     cboNota.Text = 1
     cboTipo.Text = "Filme"

     Call centralizarForm(Me)

End Sub

Private Sub lvCadastroAdicionar_Click()

 If Trim(txtNome.Text) = "" Then
   MsgBox "O campo de nome nao pode estar vazio.", vbCritical, "PREENCHA OS CAMPOS"
   Exit Sub
End If


     Select Case cboTipo.List(cboTipo.ListIndex)
          Case "FILME"
          On Error GoTo ErroNoCadastroDeFilme

               'corrige o campo "Duracao"
               If Len(txtDuracaoTemporadasAlbum.Text) = 0 Then
               txtDuracaoTemporadasAlbum = "00:00"
               End If
     

     
               ' Garante que a conexao esta aberta
               If connectBD.State = adStateClosed Then connectBD.Open
     
               'Declaracao de uma var do tipo ADODB.Command, que fara a SQL
               Dim cmdFilme As New ADODB.Command
               Set cmdFilme = New ADODB.Command
     
               'Conectando esse comando ao banco
               cmdFilme.ActiveConnection = connectBD
     
               'A query que tera os valores substituidos:
               cmdFilme.CommandText = "INSERT INTO Filmes (Nome, Diretor, Atores, Duracao, Genero, Nota, Observacao, Grupo, Excluido) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"

               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 5, txtDuracaoTemporadasAlbum.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 6, "Filmes")
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adInteger, adParamInput, , 0)


               cmdFilme.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               Unload Me
               Load frmMidia
          Exit Sub

ErroNoCadastroDeFilme:
          MsgBox "Erro no cadastro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
     
     Case "SERIE"
     On Error GoTo ErroNoCadastroDeSerie
     
               ' Garante que a conexao esta aberta
               If connectBD.State = adStateClosed Then connectBD.Open
     
               'Declaracao de uma var do tipo ADODB.Command, que fara a SQL
               Dim cmdSerie As New ADODB.Command
               Set cmdSerie = New ADODB.Command
     
               'Conectando esse comando ao banco
               cmdSerie.ActiveConnection = connectBD
     
               'A query que tera os valores substituidos:
               cmdSerie.CommandText = "INSERT INTO Series (Nome, Diretor, Atores, Temporadas, Genero, Nota, Observacao, Grupo, Excluido) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adInteger, adParamInput, , CInt(txtDuracaoTemporadasAlbum.Text))
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 6, "Series")
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adInteger, adParamInput, 1, CInt(0))
               cmdSerie.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               Unload Me
               Load frmMidia
          Exit Sub

ErroNoCadastroDeSerie:
          MsgBox "Erro no cadastro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
     
     Case "MUSICA"
          
     On Error GoTo ErroNoCadastroDeMusica
     
               ' Garante que a conexao esta aberta
               If connectBD.State = adStateClosed Then connectBD.Open
     
               'Declaracao de uma var do tipo ADODB.Command, que fara a SQL
               Dim cmdMusica As New ADODB.Command
               Set cmdMusica = New ADODB.Command
     
               'Conectando esse comando ao banco
               cmdMusica.ActiveConnection = connectBD
     
               'A query que tera os valores substituidos:
               cmdMusica.CommandText = "INSERT INTO Musicas (Nome, Artista, Participantes, Album, Genero, Nota, Observacao, Grupo, Excluido) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"
               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdMusica.Parameters.Append cmdFilme.CreateParameter(, adInteger, adParamInput, , CInt(txtDuracaoTemporadasAlbum.Text))
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 7, "Musicas")
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adInteger, adParamInput, 1, 0)
               cmdMusica.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               frmMidia.Show
               Unload Me
          Exit Sub

ErroNoCadastroDeMusica:
          MsgBox "Erro no cadastro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
          
     Case Else
          MsgBox "SELECIONE UM TIPO DE MIDIA PARA SER CADASTRADA", vbExclamation, "SELECIONE UM TIPO"
     
     End Select
End Sub

Private Sub lvCadastroVoltar_Click()
frmMidia.Show
Unload Me
End Sub

Private Sub txtDuracaoTemporadasAlbum_KeyPress(KeyAscii As Integer)

If txtDuracaoTemporadasAlbum.Tag = "tagDuracao" Or txtDuracaoTemporadasAlbum.Tag = "tagTemporadas" Then
     ' Permitir numeros de 0 a 9 (ASCII 49 a 57), dois pontos ":" (ASCII 58) e Backspace (ASCII 8)
     If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then 'Or KeyAscii = 58 - antes permitia 2 pontos
               If KeyAscii = 8 And Len(txtDuracaoTemporadasAlbum.Text) = 3 And txtDuracaoTemporadasAlbum.SelStart = Len(txtDuracaoTemporadasAlbum.Text) Then
                    txtDuracaoTemporadasAlbum.Text = Left(txtDuracaoTemporadasAlbum.Text, 1)
                    
               End If

          Exit Sub  ' Permite a tecla pressionada
     End If

     If (KeyAscii >= 49 And KeyAscii <= 57) Then
     ' permitir numeros de 1 a 9
          Exit Sub
     End If

     ' permitir a tecla Backspace
     If KeyAscii = 8 Then
          Exit Sub
     End If

     KeyAscii = 0 ' impedir que a tecla seja processada
End If
          
    
End Sub

Private Sub txtDuracaoTemporadasAlbum_Change()
     If txtDuracaoTemporadasAlbum.Tag = "tagDuracao" Then
         Dim valor As String
         'txtDuracaoTemporadasAlbum.Text = Replace(txtDuracaoTemporadasAlbum.Text, ":", "")
         valor = Replace(txtDuracaoTemporadasAlbum.Text, ":", "") ' Remove os ":"

         ' Se tiver pelo menos 2 dígitos, insere os ":"
         If Len(valor) >= 2 Then
             txtDuracaoTemporadasAlbum.Text = Left(valor, 2) & ":" & Mid(valor, 3, 2)
             txtDuracaoTemporadasAlbum.SelStart = Len(txtDuracaoTemporadasAlbum.Text) ' Mantém o cursor no final
         End If
     End If
End Sub
