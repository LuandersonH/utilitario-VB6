VERSION 5.00
Begin VB.Form frmFavorites 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Favoritos"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15960
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
   ScaleHeight     =   8070
   ScaleWidth      =   15960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnMedia 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "MÍDIAS"
      Height          =   1485
      Left            =   8385
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3375
      Width           =   3315
   End
   Begin VB.Frame frameChoiceFav 
      BackColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   1680
      TabIndex        =   0
      Top             =   1260
      Width           =   12795
      Begin VB.CommandButton btnVehicles 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "VEICULOS"
         Height          =   1485
         Left            =   2550
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2145
         Width           =   3315
      End
      Begin VB.Label lblMyFavs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GERENCIAMENTO DE FAVORITOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1635
         TabIndex        =   1
         Top             =   510
         Width           =   9810
      End
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMedia_Click()
frmMedia.Show
End Sub

Private Sub btnVehicles_Click()
frmVehicles.Show
End Sub
