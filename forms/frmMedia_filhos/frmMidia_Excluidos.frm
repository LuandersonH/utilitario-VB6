VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmMidia_Excluidos 
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
   LinkTopic       =   "frmMidia_Excluidos"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnVoltar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAIR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   9780
      Picture         =   "frmMidia_Excluidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6690
      Width           =   1830
   End
   Begin VB.CommandButton btnEstornarMedia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   6660
      Picture         =   "frmMidia_Excluidos.frx":25CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1860
   End
   Begin VB.CommandButton btnReloadList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECARREGAR LISTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3735
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMidia_Excluidos.frx":4B94
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6735
      Width           =   1875
   End
   Begin VB.TextBox inputMediaFilter 
      BackColor       =   &H00C0C0FF&
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
      Left            =   45
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
      BackColorBkg    =   12632319
      GridColor       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblMediaExcluidaInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESQUISE PELAS MIDIAS EXCLUIDAS"
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
      Left            =   4635
      TabIndex        =   2
      Top             =   60
      Width           =   5415
   End
End
Attribute VB_Name = "frmMidia_Excluidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEstornarMedia_Click()
On Error GoTo erroEstornoDeMidia
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

   If MsgBox("Realmente deseja estornar: " & nomeMediaSelecionada & ", do grupo: " & grupoMediaSelecionada & "?", vbYesNo, "E X C L U I R ?") = vbNo Then
      Exit Sub
   Else
      If connectBD.State = adStateClosed Then connectBD.Open
      queryDeletarMedia = "UPDATE " & grupoMediaSelecionada & " Set Excluido = 0 Where Codigo = " & codigoMediaSelecionada
      connectBD.Execute queryDeletarMedia
   End If
   Exit Sub

erroEstornoDeMidia:
   MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "E R R O !"
End Sub

Private Sub btnReloadList_Click()
   On Error GoTo erroAoRecarregarGridMediaExcluido

   Call UnionFilmesSeriesMusicas
   Call CarregarTodasAsMediasExcluidas(Me)
   Exit Sub

erroAoRecarregarGridMediaExcluido:
     MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "E R R O !"
End Sub

Private Sub btnVoltar_Click()
frmMidia.Show
Unload Me
End Sub

Private Sub Form_Load()
     Call centralizarForm(Me)
     Call setarColunasIniciaisDoGridMedia(Me)
End Sub

Private Sub inputMediaFilter_Change()
Call pesquisarExcluidosNoInputMediaFilterComLike(Me)
End Sub

