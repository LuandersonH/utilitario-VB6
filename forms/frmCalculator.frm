VERSION 5.00
Begin VB.Form frmCalculator 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   7455
   ClientLeft      =   10140
   ClientTop       =   2760
   ClientWidth     =   4680
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
   ScaleHeight     =   7455
   ScaleWidth      =   4680
   Begin VB.CommandButton btnConfirm 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   3360
      TabIndex        =   18
      Top             =   5100
      Width           =   1000
   End
   Begin VB.CommandButton btnSub 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3360
      TabIndex        =   17
      Top             =   4050
      Width           =   1000
   End
   Begin VB.CommandButton btnMult 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3360
      TabIndex        =   16
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton btnDivision 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3360
      TabIndex        =   15
      Top             =   1950
      Width           =   1000
   End
   Begin VB.CommandButton btnPercantage 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   10
      Left            =   2310
      TabIndex        =   14
      Top             =   1950
      Width           =   1000
   End
   Begin VB.CommandButton btnSome 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2310
      TabIndex        =   13
      Top             =   6150
      Width           =   1000
   End
   Begin VB.CommandButton btnPoint 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1260
      TabIndex        =   12
      Top             =   6150
      Width           =   1000
   End
   Begin VB.CommandButton btnClear 
      BackColor       =   &H0080C0FF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1950
      Width           =   2050
   End
   Begin VB.CommandButton digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   9
      Left            =   210
      TabIndex        =   10
      Top             =   6150
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   8
      Left            =   2310
      TabIndex        =   9
      Top             =   5100
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   7
      Left            =   1260
      TabIndex        =   8
      Top             =   5100
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   6
      Left            =   210
      TabIndex        =   7
      Top             =   5100
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   5
      Left            =   2310
      TabIndex        =   6
      Top             =   4050
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   4
      Left            =   1260
      TabIndex        =   5
      Top             =   4050
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   4050
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   2
      Left            =   2310
      TabIndex        =   3
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   1
      Left            =   1260
      TabIndex        =   2
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "@MingLiU-ExtB"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   3000
      Width           =   1000
   End
   Begin VB.TextBox display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4690
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim var1 As Double, var2 As Double
Dim opern As String


Private Sub btnClear_Click()
display.Text = ""
End Sub

Private Sub btnConfirm_Click()
var2 = Val(display.Text)
Select Case opern
   'som e sub
   Case "+": display.Text = var1 + var2
   Case "-": display.Text = var1 - var2
   'div e mult
   Case "/": display.Text = var1 / var2
   Case "*": display.Text = var1 * var2
   'percent
   Case "%": display.Text = (var1 * 1 / 100) * var2
   
End Select

End Sub

Private Sub btnDivision_Click()
var1 = Val(display.Text)
opern = "/"
display.Text = ""
End Sub

Private Sub btnMult_Click()
var1 = Val(display.Text)
opern = "*"
display.Text = ""
End Sub

Private Sub btnPercantage_Click(Index As Integer)
var1 = Val(display.Text)
opern = "%"
display.Text = ""
End Sub

Private Sub btnPoint_Click()
display.Text = display.Text + "."
End Sub

Private Sub btnSome_Click()
var1 = Val(display.Text)
opern = "+"
display.Text = ""
End Sub

Private Sub btnSub_Click()
var1 = Val(display.Text)
opern = "-"
display.Text = ""
End Sub

Private Sub digits_Click(Index As Integer)
display.Text = display.Text + digits(Index).Caption
End Sub

Private Sub display_KeyPress(KeyAscii As Integer)
       If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If

End Sub


