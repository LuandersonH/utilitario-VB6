VERSION 5.00
Begin VB.Form frmMediaAdicionar 
   Caption         =   "Cadastro"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9750
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
   ScaleHeight     =   3705
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   17
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
      ItemData        =   "frmMediaAdicionar.frx":0000
      Left            =   7260
      List            =   "frmMediaAdicionar.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Tag             =   "tagNota"
      Top             =   1800
      Width           =   1830
   End
   Begin VB.CommandButton btnVoltar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "VOLTAR"
      Height          =   675
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2955
      Width           =   2190
   End
   Begin VB.CommandButton btnAdicionar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADICIONAR"
      Height          =   675
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2940
      Width           =   2190
   End
   Begin VB.TextBox txtObservacao 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Tag             =   "tagObservacao"
      Top             =   2535
      Width           =   9645
   End
   Begin VB.TextBox txtGenero 
      Height          =   315
      Left            =   1500
      TabIndex        =   10
      Tag             =   "tagGenero"
      Top             =   1800
      Width           =   5565
   End
   Begin VB.TextBox txtAtoresParticipantes 
      Height          =   315
      Left            =   4000
      TabIndex        =   7
      Tag             =   "tagAtores"
      Top             =   1095
      Width           =   5550
   End
   Begin VB.TextBox txtDiretorArtista 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Tag             =   "tagDiretor"
      Top             =   1110
      Width           =   3705
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   2445
      TabIndex        =   3
      Tag             =   "tagNome"
      Top             =   405
      Width           =   7125
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmMediaAdicionar.frx":0004
      Left            =   105
      List            =   "frmMediaAdicionar.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "tagTipo"
      Top             =   400
      Width           =   1830
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1200
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
      TabIndex        =   11
      Top             =   1600
      Width           =   810
   End
   Begin VB.Label lblGenero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gênero"
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
      TabIndex        =   9
      Top             =   1605
      Width           =   810
   End
   Begin VB.Label lblDuracao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duração"
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   2
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
Attribute VB_Name = "frmMediaAdicionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdicionar_Click()
     Select Case cboTipo.List(cboTipo.ListIndex)
          Case "FILME"
          On Error GoTo ErroNoCadastroDeFilme
     
               ' Garante que a conexao esta aberta
               If connectBD.State = adStateClosed Then connectBD.Open
     
               'Declaracao de uma var do tipo ADODB.Command, que fara a SQL
               Dim cmdFilme As New ADODB.Command
               Set cmdFilme = New ADODB.Command
     
               'Conectando esse comando ao banco
               cmdFilme.ActiveConnection = connectBD
     
               'A query que tera os valores substituidos:
               cmdFilme.CommandText = "INSERT INTO Filmes (Nome, Diretor, Atores, Duracao, Genero, Nota, Observacao) VALUES (?, ?, ?, ?, ?, ?, ?)"
               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adDate, adParamInput, 255, txtDuracaoTemporadasAlbum.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdFilme.Parameters.Append cmdFilme.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               
               cmdFilme.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               Unload Me
               Load frmMedia
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
               cmdSerie.CommandText = "INSERT INTO Series (Nome, Diretor, Atores, Temporadas, Genero, Nota, Observacao) VALUES (?, ?, ?, ?, ?, ?, ?)"
               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adInteger, adParamInput, , CInt(txtDuracaoTemporadasAlbum.Text))
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdSerie.Parameters.Append cmdSerie.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               
               cmdSerie.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               Unload Me
               Load frmMedia
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
               cmdMusica.CommandText = "INSERT INTO Musicas (Nome, Artista, Participantes, Album, Genero, Nota, Observacao) VALUES (?, ?, ?, ?, ?, ?, ?)"
               
               'Substituicao dos parametros pelos dados dos inputs:
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtNome.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtDiretorArtista.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtAtoresParticipantes.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtDuracaoTemporadasAlbum.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtGenero.Text)
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adInteger, adParamInput, , CInt(cboNota.Text))
               cmdMusica.Parameters.Append cmdMusica.CreateParameter(, adVarChar, adParamInput, 255, txtObservacao.Text)
               
               cmdMusica.Execute
               
               MsgBox "Cadastro realizado com sucesso", vbExclamation, "SUCESSO"
               Unload Me
               Load frmMedia
          Exit Sub

ErroNoCadastroDeMusica:
          MsgBox "Erro no cadastro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
          
     Case Else
          MsgBox "SELECIONE UM TIPO DE MÍDIA PARA SER CADASTRADA", vbExclamation, "SELECIONE UM TIPO"
     
     End Select

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

End Sub

Private Sub txtDuracaoTemporadasAlbum_KeyPress(KeyAscii As Integer)

If txtDuracaoTemporadasAlbum.Tag = "tagDuracao" Or txtDuracaoTemporadasAlbum.Tag = "tagTemporadas" Then
     ' Permitir números de 0 a 9 (ASCII 49 a 57), dois pontos ":" (ASCII 58) e Backspace (ASCII 8)
     If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 58 Or KeyAscii = 8 Then
          Exit Sub  ' Permite a tecla pressionada
     End If

     ' Permitir apenas números de 1 a 9 (ASCII 49 a 57)

     If txtDuracaoTemporadasAlbum.Tag = "tagDuracao" Then
    If KeyAscii = 58 Then ' Código ASCII do ":"
     Exit Sub
          End If
     End If

     If (KeyAscii >= 49 And KeyAscii <= 57) Then
     ' Permitir números de 1 a 9
          Exit Sub
     End If

     ' Permitir a tecla Backspace (Código ASCII 8)
     If KeyAscii = 8 Then
          Exit Sub
     End If

     KeyAscii = 0 ' Impede que a tecla seja processada
End If
          
    
End Sub
