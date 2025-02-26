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
   Begin VB.CommandButton btnVoltar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "VOLTAR"
      Height          =   675
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2955
      Width           =   2190
   End
   Begin VB.CommandButton btnAdicionar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADICIONAR"
      Height          =   675
      Left            =   1965
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2940
      Width           =   2190
   End
   Begin VB.TextBox txtObservacao 
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Tag             =   "tagObservacao"
      Top             =   2535
      Width           =   9645
   End
   Begin VB.TextBox txtNota 
      Height          =   315
      Left            =   7335
      TabIndex        =   13
      Tag             =   "tagNota"
      Top             =   1785
      Width           =   2160
   End
   Begin VB.TextBox txtGenero 
      Height          =   315
      Left            =   4005
      TabIndex        =   11
      Tag             =   "tagGenero"
      Top             =   1770
      Width           =   3255
   End
   Begin VB.TextBox txtDuracao 
      Height          =   315
      Left            =   90
      TabIndex        =   9
      Tag             =   "tagDuracao"
      Top             =   1770
      Width           =   3705
   End
   Begin VB.TextBox txtAtores 
      Height          =   315
      Left            =   4000
      TabIndex        =   7
      Tag             =   "tagAtores"
      Top             =   1095
      Width           =   5550
   End
   Begin VB.TextBox txtDiretor 
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
      ItemData        =   "frmMediaAdicionar.frx":0000
      Left            =   105
      List            =   "frmMediaAdicionar.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
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
      TabIndex        =   14
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
      TabIndex        =   12
      Top             =   1560
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
      Left            =   4050
      TabIndex        =   10
      Top             =   1545
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
      Top             =   1550
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
MsgBox "FILME"

Case "SERIE"
MsgBox "SERIE"

Case "MUSICA"
MsgBox "MUSICA"

Case Else
 MsgBox "SELECIONE UM TIPO DE MÍDIA PARA SER CADASTRADA", vbExclamation, "SELECIONE UM TIPO"

End Select


End Sub

Private Sub cboTipo_Click()
Call AtualizarCamposPorTipo(Me)
End Sub
