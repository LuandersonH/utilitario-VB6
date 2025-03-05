VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00E0E0E0&
   Caption         =   "UTILITÁRIO"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameFavoritos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FAVORITOS"
      Height          =   3540
      Left            =   10200
      TabIndex        =   2
      Top             =   3000
      Width           =   2800
      Begin VB.CommandButton btnEntrarFavoritos 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ENTRAR"
         Height          =   800
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2620
         Width           =   2500
      End
      Begin VB.Image imgFavs 
         Height          =   1995
         Left            =   555
         Picture         =   "frmHome.frx":0000
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
      Begin VB.CommandButton btnEntrarTodolist 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ENTRAR"
         Height          =   800
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2620
         Width           =   2500
      End
      Begin VB.Image imgToDoList 
         Height          =   1995
         Left            =   510
         Picture         =   "frmHome.frx":F216
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
      Begin VB.CommandButton btnEntrarCalculadora 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ENTRAR"
         Height          =   800
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2620
         Width           =   2500
      End
      Begin VB.Image imgCalculator 
         Height          =   1995
         Left            =   510
         Picture         =   "frmHome.frx":108C3
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
Private Sub btnEntrarCalculadora_Click()
frmCalculator.Show
End Sub

Private Sub btnEntrarFavoritos_Click()
frmFavorites.Show
End Sub


Private Sub lblHomeIntro_Click()

End Sub

Private Sub btnEntrarTodolist_Click()
frmConsultableToDoList.Show

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("ENCERRAR O PROGRAMA?", vbYesNo + vbQuestion, "E N C E R R A R") = vbNo Then
     Cancel = 1
     Else
          End
     End If

End Sub

Private Sub Form_Load()
     'Abre maximizado
    Me.Left = 0
    Me.Top = 0
    Me.Width = (Screen.Width - 8000)
    Me.Height = (Screen.Height - 5000)

     'Conecta ao banco de dados
     Call InitConexao(Me)
End Sub

Private Sub Form_Paint()
'Me.Line (x1, y1) - (x2, y2), cor, [opção]
'-(x1, y1): Coordenadas do ponto inicial (canto superior esquerdo do retângulo).
'-(x2, y2): Coordenadas do ponto final (canto inferior direito do retângulo).
'-[opção]: Se for B, desenha apenas a borda. Se for BF, preenche o retângulo.

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

    ' Se estiver menor que o mínimo, retorna ao tamanho anterior
    If Me.Width <= minWidth Then Me.Width = minWidth
    If Me.Height <= minHeight Then Me.Height = minHeight

    ' Atualiza o tamanho salvo
    lastWidth = Me.Width
    lastHeight = Me.Height

    'frameToDoList no meio da tela
    frameToDoList.Left = (Me.ScaleWidth - frameToDoList.Width) \ 2
    frameToDoList.Top = (Me.ScaleHeight - frameToDoList.Height) \ 2

    'frameCalculadora a esquerda
    frameCalculadora.Left = (frameToDoList.Left - frameCalculadora.Width - 2000)
    frameCalculadora.Top = (Me.ScaleHeight - frameCalculadora.Height) \ 2

     'frameFavoritos a direita
    frameFavoritos.Left = (frameToDoList.Left + frameCalculadora.Width + 2000)
    frameFavoritos.Top = (Me.ScaleHeight - frameFavoritos.Height) \ 2
     
End Sub
