VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVButton.ocx"
Begin VB.Form frmHome 
   BackColor       =   &H00E0E0E0&
   Caption         =   "UTILITARIO"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameMidias 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MIDIAS FAVORITAS"
      Height          =   3540
      Left            =   10200
      TabIndex        =   2
      Top             =   3000
      Width           =   2800
      Begin lvButton.lvButtons_H lvEntrar 
         Height          =   960
         Index           =   2
         Left            =   75
         TabIndex        =   5
         Top             =   2520
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   1693
         Caption         =   "ENTRAR"
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
         Image           =   "frmHome.frx":94CA
         ImgSize         =   32
         cBack           =   14737632
      End
      Begin VB.Image imgFavs 
         Height          =   1995
         Left            =   555
         Picture         =   "frmHome.frx":BAA4
         Stretch         =   -1  'True
         Top             =   465
         Width           =   1800
      End
   End
   Begin VB.Frame frameToDoList 
      BackColor       =   &H8000000E&
      Caption         =   "TODOLIST"
      Height          =   3540
      Left            =   5600
      TabIndex        =   1
      Top             =   3000
      Width           =   2800
      Begin lvButton.lvButtons_H lvEntrar 
         Height          =   960
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   2520
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   1693
         Caption         =   "ENTRAR"
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
         Image           =   "frmHome.frx":1ACBA
         ImgSize         =   32
         cBack           =   14737632
      End
      Begin VB.Image imgToDoList 
         Height          =   1995
         Left            =   510
         Picture         =   "frmHome.frx":1D294
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.Frame frameCalculadora 
      BackColor       =   &H8000000E&
      Caption         =   "CALCULADORA"
      Height          =   3540
      Left            =   1000
      TabIndex        =   0
      Top             =   3000
      Width           =   2800
      Begin lvButton.lvButtons_H lvEntrar 
         Height          =   960
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   2520
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   1693
         Caption         =   "ENTRAR"
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
         Image           =   "frmHome.frx":1E941
         ImgSize         =   32
         cBack           =   14737632
      End
      Begin VB.Image imgCalculator 
         Height          =   1995
         Left            =   510
         Picture         =   "frmHome.frx":20F1B
         Stretch         =   -1  'True
         Top             =   420
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("ENCERRAR O PROGRAMA?", vbYesNo + vbQuestion, "E N C E R R A R") = vbYes Then
      End
     Else
      Cancel = 1
     End If

End Sub

Private Sub Form_Load()
    Me.Width = (Screen.Width - 8000)
    Me.Height = (Screen.Height - 5000)

     Call centralizarForm(Me)

     'Conecta ao banco de dados
     Call InitConexao(Me)
End Sub

Private Sub Form_Paint()
'Me.Line (x1, y1) - (x2, y2), cor, [opcao]
'-(x1, y1): Coordenadas do ponto inicial (canto superior esquerdo do retangulo).
'-(x2, y2): Coordenadas do ponto final (canto inferior direito do retangulo).
'-[op��o]: Se for B, desenha apenas a borda. Se for BF, preenche o retangulo.

     'limpa o paint
     Cls

    'primeira metade da tela (azul)
    Me.ForeColor = vbBlue
    Me.Line (0, 0)-(Me.ScaleWidth, Me.ScaleHeight / 2), vbBlue, BF

    'segunda metade da tela (verde)
    Me.ForeColor = vbGreen
    Me.Line (0, Me.ScaleHeight / 2)-(Me.ScaleWidth, Me.ScaleHeight), vbGreen, BF
End Sub

Private Sub Form_Resize()
    Static lastWidth As Integer
    Static lastHeight As Integer

    Dim minWidth As Integer
    Dim minHeight As Integer

    minWidth = 14000
    minHeight = 8000

    ' Se estiver menor que o minimo, retorna ao tamanho anterior
     If Me.WindowState = vbNormal Then
          If Me.Width <= minWidth Then Me.Width = minWidth
          If Me.Height <= minHeight Then Me.Height = minHeight
     End If

    ' Atualiza o tamanho salvo
    lastWidth = Me.Width
    lastHeight = Me.Height

    'frameToDoList no meio da tela
    frameToDoList.Left = (Me.ScaleWidth - frameToDoList.Width) \ 2
    frameToDoList.Top = (Me.ScaleHeight - frameToDoList.Height) \ 2

    'frameCalculadora a esquerda
    frameCalculadora.Left = (frameToDoList.Left - frameCalculadora.Width - 2000)
    frameCalculadora.Top = (Me.ScaleHeight - frameCalculadora.Height) \ 2

     'frameMidias a direita
    frameMidias.Left = (frameToDoList.Left + frameCalculadora.Width + 2000)
    frameMidias.Top = (Me.ScaleHeight - frameMidias.Height) \ 2
     
End Sub

Private Sub imgCalculator_Click()
frmCalculator.Show
End Sub

Private Sub imgFavs_Click()
frmMidia.Show
End Sub

Private Sub imgToDoList_Click()
frmConsultableToDoList.Show
End Sub

Private Sub lvEntrar_Click(Index As Integer)
Select Case Index
Case 0
frmCalculator.Show
Case 1
frmConsultableToDoList.Show
Case 2
frmMidia.Show
End Select
End Sub
