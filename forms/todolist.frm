VERSION 5.00
Begin VB.Form frmTodolist 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14175
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
   ScaleHeight     =   8490
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClearAll 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "LIMPAR TUDO"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10005
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5910
      Width           =   1725
   End
   Begin VB.ListBox listTasks 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      ItemData        =   "todolist.frx":0000
      Left            =   1905
      List            =   "todolist.frx":0002
      TabIndex        =   4
      Top             =   3345
      Width           =   8080
   End
   Begin VB.CommandButton btnFinishedTask 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "CONCLUÍDA"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1250
      Left            =   10020
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4650
      Width           =   1725
   End
   Begin VB.CommandButton btnDeleteTask 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "EXCLUIR"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1250
      Left            =   10020
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3345
      Width           =   1725
   End
   Begin VB.CommandButton btnInsertTask 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADICIONAR"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   10020
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1995
      Width           =   1725
   End
   Begin VB.Label lblTodolist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA DE TAREFAS"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   3420
      TabIndex        =   0
      Top             =   960
      Width           =   6810
   End
End
Attribute VB_Name = "frmTodolist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
