VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmConsultableToDoList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TodoList"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "ConsultableToDoList"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstabToDoList 
      Height          =   9045
      Left            =   -120
      TabIndex        =   0
      Top             =   -75
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   15954
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1147
      TabMaxWidth     =   5292
      BackColor       =   8388608
      TabCaption(0)   =   "Lista de tarefas"
      TabPicture(0)   =   "ConsultableToDoList.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTodolist"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "listTasks"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tboxInsertTask"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnInsertTask"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnDeleteTask"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnFinishedTask"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnClearAll"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "ConsultableToDoList.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblHistoryInput"
      Tab(1).Control(1)=   "GridHistorico"
      Tab(1).Control(2)=   "inputHistoryFilter"
      Tab(1).ControlCount=   3
      Begin VB.TextBox inputHistoryFilter 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74880
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1170
         Width           =   14775
      End
      Begin MSFlexGridLib.MSFlexGrid GridHistorico 
         Height          =   6660
         Left            =   -74895
         TabIndex        =   8
         Top             =   2040
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   11748
         _Version        =   393216
         Rows            =   1
         RowHeightMin    =   500
         BackColor       =   14737632
         BackColorFixed  =   14737632
         BackColorSel    =   16777215
         ForeColorSel    =   16777215
         BackColorBkg    =   14737632
         GridColor       =   0
         WordWrap        =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.CommandButton btnClearAll 
         BackColor       =   &H000000FF&
         Caption         =   "EXCLUIR TUDO"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   11055
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6705
         Width           =   1725
      End
      Begin VB.CommandButton btnFinishedTask 
         BackColor       =   &H0080FF80&
         Caption         =   "CONCLUIR"
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   11100
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4185
         Width           =   1725
      End
      Begin VB.CommandButton btnDeleteTask 
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
         Height          =   1095
         Left            =   11070
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5520
         Width           =   1725
      End
      Begin VB.CommandButton btnInsertTask 
         BackColor       =   &H0080C0FF&
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
         Left            =   11055
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1875
         Width           =   1725
      End
      Begin VB.TextBox tboxInsertTask 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2925
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1935
         Width           =   8080
      End
      Begin VB.ListBox listTasks 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2925
         TabIndex        =   1
         Top             =   4245
         Width           =   8080
      End
      Begin VB.Label lblHistoryInput 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESQUISE POR TAREFAS"
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
         Left            =   -69195
         TabIndex        =   10
         Top             =   750
         Width           =   3570
      End
      Begin VB.Label lblTodolist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA DE TAREFAS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   630
         Left            =   4545
         TabIndex        =   7
         Top             =   840
         Width           =   5565
      End
   End
End
Attribute VB_Name = "frmConsultableToDoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim handleTaskValue As String

'TODOLIST
'TODOLIST
Private Sub Form_Load()
     Call centralizarForm(Me)
     Call InitConexao(Me)
     Call reloadListTasks(Me)
     Call historicoConsultarTasks(Me)
End Sub

Private Sub btnInsertTask_Click()
     Call addTasks(Me)
End Sub

Private Sub btnFinishedTask_Click()
     Call endTasks(Me)
End Sub

Private Sub btnDeleteTask_Click()
     Call deleteTasks(Me)
End Sub

Private Sub btnClearAll_Click()
     resposta = MsgBox("Isso apagara TODAS AS TAREFAS e nao podera ser desfeito, deseja continuar?", vbExclamation, "Apagar todas as tarefas")
     
     If resposta = vbOK Then
          Call deleteAllTasks(Me)
     End If
End Sub


'HISTORICO
'HISTORICO
Private Sub Form_Resize()
      With GridHistorico
       
          .TextMatrix(0, 0) = "Tarefa"
          .TextMatrix(0, 1) = "Status"
      
          .ColWidth(0) = 10000
          .ColWidth(1) = 15000
    End With
End Sub

Private Sub inputHistoryFilter_Change()
    Call historicoConsultarTasks(Me)
End Sub

Private Sub sstabToDoList_Click(PreviousTab As Integer)
    'Ao clicar na tab Historico, carrega as tasks
    If sstabToDoList.Tab = 1 Then
        Call historicoConsultarTasks(Me)
    End If
End Sub

