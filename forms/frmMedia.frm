VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmMedia 
   Caption         =   "Mídias"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   405
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
   ScaleHeight     =   7755
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDeleteMídia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXCLUÍR MÍDIA"
      Height          =   930
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Width           =   3225
   End
   Begin VB.CommandButton btnAddMedia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADICIONAR MÍDIA"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   3285
   End
   Begin VB.CommandButton btnReloadList 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RECARREGAR LISTA"
      Height          =   960
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6735
      Width           =   3285
   End
   Begin VB.TextBox inputMediaFilter 
      BackColor       =   &H00C0C0C0&
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
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      RowHeightMin    =   500
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblMediaInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESQUISE PELAS MÍDIAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5430
      TabIndex        =   1
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMedia_Click()

End Sub

Private Sub lblVehicles_Click()

End Sub

Private Sub comboTipo_Change()

End Sub

Private Sub btnAddMedia_Click()
frmMediaAdicionar.Show
End Sub

Private Sub btnReloadList_Click()
On Error GoTo erroAoCarregarMidias
     Dim queryUnionFilmesSeriesMusicas As String
     Dim linhaAtualMedia As Integer
     
     'campos no BD: CODIGO - NOME - DIRETOR - ATORES - TEMPORADAS - GENERO - NOTA - OBSERVAÇÃO - ARTISTA - PARTICIPANTES - ALBUM - DURAÇÃO

     queryUnionFilmesSeriesMusicas = "SELECT Codigo, Nome, Diretor, Atores, 0 AS Temporadas, Genero, Nota, Observacao, Null AS Artista, Null AS Participantes, Null AS Album, Duracao FROM Filmes " & _
     "UNION ALL " & _
     "SELECT Codigo, Nome, Diretor, Atores, Temporadas, Genero, Nota, Observacao, Null AS Artista, Null AS Participantes, Null AS Album, Null AS Duracao FROM Series " & _
     "UNION ALL " & _
      "SELECT Codigo, Nome, Null AS Diretor, Null AS Atores, Null AS Temporadas, Genero, Nota, Observacao, Artista, Participantes, Album, Null AS Duracao FROM Musicas"

     If connectBD.State = adStateClosed Then connectBD.Open

     If recordBD.State = adStateOpen Then recordBD.Close
          recordBD.Open queryUnionFilmesSeriesMusicas, connectBD, adOpenStatic, adLockReadOnly

     With GridMedia
               .Clear
               .Cols = 12
               .Rows = 1
               .TextMatrix(0, 0) = "Codigo"
               .TextMatrix(0, 1) = "Nome"
               .TextMatrix(0, 2) = "Diretor"
               .TextMatrix(0, 3) = "Atores"
               .TextMatrix(0, 4) = "Temporadas"
               .TextMatrix(0, 5) = "Genero"
               .TextMatrix(0, 6) = "Nota"
               .TextMatrix(0, 7) = "Observacao"
               .TextMatrix(0, 8) = "Artista"
               .TextMatrix(0, 9) = "Participantes"
               .TextMatrix(0, 10) = "Album"
               .TextMatrix(0, 11) = "Duracao"
     End With

     linhaAtualMedia = 1

     While Not recordBD.EOF
          With GridMedia
               
               .Rows = GridMedia.Rows + 1
               .TextMatrix(linhaAtualMedia, 0) = IIf(IsNull(recordBD!Codigo), 0, recordBD!Codigo)
               .TextMatrix(linhaAtualMedia, 1) = IIf(IsNull(recordBD!Nome), "", recordBD!Nome)
               .TextMatrix(linhaAtualMedia, 2) = IIf(IsNull(recordBD!Diretor), "", recordBD!Diretor)
               .TextMatrix(linhaAtualMedia, 3) = IIf(IsNull(recordBD!Atores), "", recordBD!Atores)
               .TextMatrix(linhaAtualMedia, 4) = IIf(IsNull(recordBD!Temporadas), 0, recordBD!Temporadas)
               .TextMatrix(linhaAtualMedia, 5) = IIf(IsNull(recordBD!Genero), "", recordBD!Genero)
               .TextMatrix(linhaAtualMedia, 6) = IIf(IsNull(recordBD!Nota), 0, recordBD!Nota)
               .TextMatrix(linhaAtualMedia, 7) = IIf(IsNull(recordBD!Observacao), "", recordBD!Observacao)
               .TextMatrix(linhaAtualMedia, 8) = IIf(IsNull(recordBD!Artista), "", recordBD!Artista)
               .TextMatrix(linhaAtualMedia, 9) = IIf(IsNull(recordBD!Participantes), "", recordBD!Participantes)
               .TextMatrix(linhaAtualMedia, 10) = IIf(IsNull(recordBD!Album), "", recordBD!Album)
               .TextMatrix(linhaAtualMedia, 11) = IIf(IsNull(recordBD!Duracao), "", recordBD!Duracao)

               recordBD.MoveNext
               linhaAtualMedia = linhaAtualMedia + 1
          End With
      Wend

     recordBD.Close
     Exit Sub

erroAoCarregarMidias:
MsgBox "Erro ao carregar midias: " & Err.Number & " - " & Err.Description, vbCritical, "ERRO!"
 If recordBD.State = adStateOpen Then recordBD.Close
End Sub

