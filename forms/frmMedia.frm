VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmMidia 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Midias"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14790
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
   ScaleHeight     =   7755
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H lvEstorno 
      Height          =   1000
      Left            =   12500
      TabIndex        =   7
      Top             =   6725
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "RECUPERAR MIDIAS"
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
      Image           =   "frmMedia.frx":0000
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H lvVoltar 
      Height          =   1005
      Left            =   9480
      TabIndex        =   6
      Top             =   6725
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
      Image           =   "frmMedia.frx":25DA
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H lvExcluir 
      Height          =   1005
      Left            =   6500
      TabIndex        =   5
      Top             =   6726
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "EXCLUIR"
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
      Image           =   "frmMedia.frx":4BB4
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H lvRecarregar 
      Height          =   1005
      Left            =   3500
      TabIndex        =   4
      Top             =   6725
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "RECARREGAR"
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
      Image           =   "frmMedia.frx":718E
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H lvCadastrar 
      Height          =   1005
      Left            =   500
      TabIndex        =   3
      Top             =   6725
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1773
      Caption         =   "CADASTRAR"
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
      Image           =   "frmMedia.frx":9768
      ImgSize         =   32
      cBack           =   14737632
   End
   Begin VB.TextBox inputMediaFilter 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   450
      Width           =   14775
   End
   Begin MSFlexGridLib.MSFlexGrid GridMedia 
      Height          =   5415
      Left            =   30
      TabIndex        =   1
      Top             =   1290
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   700
      BackColor       =   14737632
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12632256
      BackColorBkg    =   12648447
      GridColor       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblMediaInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESQUISE PELAS MIDIAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   5430
      TabIndex        =   2
      Top             =   60
      Width           =   3660
   End
End
Attribute VB_Name = "frmMidia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo erroLoadAoRecarregarGridMedia
     Call centralizarForm(Me)
     Call setarColunasIniciaisDoGridMedia(Me)
     
     Call UnionFilmesSeriesMusicas
     Call CarregarTodasAsMedias(Me)
     Exit Sub

erroLoadAoRecarregarGridMedia:
     MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "E R R O !"
End Sub


Private Sub inputMediaFilter_Change()
Call pesquisarNoInputMediaFilterComLike(Me)
End Sub

Private Sub lvCadastrar_Click()
frmMidia_Cadastro.Show
Unload Me
End Sub

Private Sub lvEstorno_Click()
frmMidia_Excluidos.Show
End Sub

Private Sub lvExcluir_Click()
On Error GoTo erroDeleteMedia
   Dim codigoMediaSelecionada As Integer
   Dim grupoMediaSelecionada As String
   Dim nomeMediaSelecionada As String
   Dim queryDeletarMedia As String

   If GridMedia.Rows <= 0 Or GridMedia.Row <= 0 Then
      MsgBox "Selecione uma midia para excluir!", vbExclamation
      Exit Sub
   End If

   codigoMediaSelecionada = GridMedia.TextMatrix(GridMedia.Row, 0)
   grupoMediaSelecionada = GridMedia.TextMatrix(GridMedia.Row, 12)
   nomeMediaSelecionada = GridMedia.TextMatrix(GridMedia.Row, 1)

   If MsgBox("Realmente deseja excluir " & nomeMediaSelecionada & " De " & grupoMediaSelecionada & "?", vbYesNo, "E X C L U I R ?") = vbNo Then
      Exit Sub
   Else
      If connectBD.State = adStateClosed Then connectBD.Open
      queryDeletarMedia = "UPDATE " & grupoMediaSelecionada & " Set Excluido = 1 Where Codigo = " & codigoMediaSelecionada
      connectBD.Execute queryDeletarMedia
     Call CarregarTodasAsMedias(Me)
   End If
   Exit Sub

erroDeleteMedia:
   MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "E R R O !"
End Sub

Private Sub lvRecarregar_Click()
   On Error GoTo erroAoRecarregarGridMedia

   Call UnionFilmesSeriesMusicas
   Call CarregarTodasAsMedias(Me)
   Exit Sub

erroAoRecarregarGridMedia:
     MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "E R R O !"
End Sub

Private Sub lvVoltar_Click()
frmHome.Show
Unload Me
End Sub
