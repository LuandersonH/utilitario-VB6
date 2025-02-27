VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   Caption         =   "UTILITÁRIO"
   ClientHeight    =   7980
   ClientLeft      =   0
   ClientTop       =   0
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnEntrarFavoritos 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ENTRAR"
      Height          =   810
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5650
      Width           =   2595
   End
   Begin VB.CommandButton btnEntrarTodolist 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ENTRAR"
      Height          =   810
      Left            =   5685
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5650
      Width           =   2595
   End
   Begin VB.Frame frameFavoritos 
      BackColor       =   &H8000000E&
      Caption         =   "FAVORITOS"
      Height          =   3540
      Left            =   10200
      TabIndex        =   5
      Top             =   3000
      Width           =   2800
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
      TabIndex        =   4
      Top             =   3000
      Width           =   2800
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
      TabIndex        =   3
      Top             =   3000
      Width           =   2800
      Begin VB.CommandButton btnEntrarCalculadora 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ENTRAR"
         Height          =   810
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2650
         Width           =   2595
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
   Begin VB.Frame frameHome 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   6045
      Left            =   15
      TabIndex        =   1
      Top             =   -30
      Width           =   14000
      Begin VB.Label lblAppName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UTILITÁRIO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   630
         Left            =   4965
         TabIndex        =   2
         Top             =   585
         Width           =   3615
      End
   End
   Begin VB.Frame frameIntro 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   -60
      TabIndex        =   0
      Top             =   -30
      Width           =   18465
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

Private Sub btnEntrarTodolist_Click()
ConsultableToDoList.Show
End Sub

Private Sub lblHomeIntro_Click()

End Sub

Private Sub Form_Load()
Call InitConexao(Me)
End Sub

