VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M�dias"
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
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDeleteM�dia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXCLU�R M�DIA"
      Height          =   930
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Width           =   3225
   End
   Begin VB.CommandButton btnAddMedia 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADICIONAR M�DIA"
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
      TabIndex        =   1
      Top             =   1320
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   1
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblMediaInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESQUISE PELAS M�DIAS"
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
      TabIndex        =   2
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
Unload Me
End Sub

Private Sub btnReloadList_Click()
Call UnionFilmesSeriesMusicas
Call CarregarTodasAsMedias(Me)
GridMedia.BackColor = vbWhite 'grid sem ser o fixed
'GridMedia.GridColorFixed = vbBlack
GridMedia.BackColorFixed = vbRed
End Sub

'Private Sub inputMediaFilter_Change()
'Dim textoDoInputMedia As String
'textoDoInputMedia = inputMediaFilter.Text

'Call UnionFilmesSeriesMusicas

'MsgBox textoDoInputMedia

'queryInputMediaFilter = "SELECT * From " & UnionFilmesSeriesMusicas & " WHERE Nome LIKE " * " & textoDoInputMedia & " * ""
'queryInputMediaFilter = "SELECT * FROM (" & UnionFilmesSeriesMusicas & ") AS Midia WHERE Nome LIKE '*" & textoDoInputMedia & "*'"
'End Sub

Private Sub inputMediaFilter_Change()
     Dim textoDoInputMedia As String
     Dim queryInputMediaFilter As String
     Dim queryUnion As String
     Dim linhaAtualMedia As Integer

    textoDoInputMedia = inputMediaFilter.Text
    Call UnionFilmesSeriesMusicas

    ' Corrigir a construcao do filtro com LIKE
    queryInputMediaFilter = "SELECT * FROM (" & UnionFilmesSeriesMusicas & ") AS Midia WHERE Nome LIKE '%" & textoDoInputMedia & "%'"


    ' Mensagem de depuracao para verificar se a query final est� correta
    Debug.Print queryInputMediaFilter
    
     If connectBD.State = adStateClosed Then connectBD.Open

     If recordBD.State = adStateOpen Then recordBD.Close
          recordBD.Open queryInputMediaFilter, connectBD, adOpenStatic, adLockReadOnly

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

               .ColWidth(0) = Width / 12
               .ColWidth(1) = Width / 12
               .ColWidth(2) = Width / 12
               .ColWidth(3) = Width / 12
               .ColWidth(4) = Width / 12
               .ColWidth(5) = Width / 12
               .ColWidth(6) = Width / 12
               .ColWidth(7) = Width / 12
               .ColWidth(8) = Width / 12
               .ColWidth(9) = Width / 12
               .ColWidth(10) = Width / 12
               .ColWidth(11) = Width / 12
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
End Sub

