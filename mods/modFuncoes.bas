Attribute VB_Name = "mod_Funcoes_ToDoList"
Public Function centralizarForm(frm)
    frm.Left = (Screen.Width / 2) - (frm.Width / 2)
    frm.Top = (Screen.Height / 2) - (frm.Height / 2)
End Function

Public Function reloadListTasks(frm As Object)
    Dim querySelectTasksPendentes As String
   querySelectTasksPendentes = "SELECT Codigo, Descricao FROM Tasks WHERE Status = 'PENDENTE' ORDER BY Codigo ASC"

    On Error GoTo reloadErro

    If recordBD.State = adStateOpen Then recordBD.Close
    recordBD.Open querySelectTasksPendentes, connectBD, adOpenStatic, adLockReadOnly

    frm.listTasks.Clear

    If Not recordBD.EOF Then
        While Not recordBD.EOF
            'add a descricao na listTasks para ser exibida
            frm.listTasks.AddItem recordBD.Fields("Descricao").Value
            
            'ItemData = informação unica e oculta ao usuario. Cada tarefa recebe seu Codigo do banco de dados
            frm.listTasks.ItemData(frm.listTasks.NewIndex) = recordBD.Fields("Codigo").Value

             'avança p prox. registro
            recordBD.MoveNext
        Wend
      End If
      

    recordBD.Close
    Exit Function
reloadErro:
   MsgBox "Erro ao carregar a lista: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function

Public Function addTasks(frm As Object)
    Dim queryAddTask As String
    Dim tarefaClicadaDesc As String
    Dim newTask As String

    On Error GoTo addTasksErro
    ' Verifica se uma tarefa foi selecionada
    If Trim(frm.tboxInsertTask.Text) <> "" Then
        'corta os espaços
        newTask = Trim(frm.tboxInsertTask.Text)
        'troca uma aspa simples por 2 aspas simples para ser interpretado certo no banco, caso o usuario digite uma aspa simples na descricao da tarefa
        newTask = Replace(newTask, "'", "''")

        ' Query correta para o INSERT
        queryAddTask = "INSERT INTO Tasks (descricao, status) VALUES ('" & newTask & "', 'PENDENTE')"

        ' Executa a query no BD
        If connectBD.State = adStateClosed Then connectBD.Open
        connectBD.Execute queryAddTask

        MsgBox "Tarefa adicionada com sucesso!", vbInformation, "Sucesso"

        frm.tboxInsertTask.Text = ""
        
        
        Call reloadListTasks(frm)
        Exit Function
    Else
        MsgBox "Crie uma tarefa antes de adicionar a lista!", vbExclamation, "Aviso"
        Exit Function
    End If

addTasksErro:
        MsgBox "Erro ao adicionar tarefa: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function

Public Function ConsultarTasks(frm As Object)
    Dim queryInputHistory As String

    If frm.inputHistoryFilter.Text = "" Then
        MsgBox "Preencha o campo de pesquisa antes de consultar!", vbExclamation, "Aviso"
        Exit Function
    End If

    queryInputHistory = "SELECT * FROM Tasks WHERE Descricao = '" & Replace(frm.inputHistoryFilter.Text, "'", "''") & "'"

    If recordBD.State = adStateOpen Then recordBD.Close

    recordBD.Open queryInputHistory, connectBD, adOpenStatic, adLockReadOnly

    frm.inputHistoryFilter.Text = ""
 
    ' Verifica se encontrou resultados antes de acessar os campos
    If Not recordBD.EOF Then
        MsgBox "Tarefa encontrada: " & recordBD.Fields("Descricao").Value, vbInformation, "Resultado"
    Else
        MsgBox "Nenhuma tarefa encontrada!", vbExclamation, "Aviso"
    End If

    ' Fecha o Recordset
    recordBD.Close
End Function

Public Function endTasks(frm As Object)
    Dim queryUpdateTasks As String
    Dim tarefaClicadaCodigo As Integer

    On Error GoTo erroEndTasks
    
    ' Verifica se uma tarefa foi selecionada
    If frm.listTasks.ListIndex <> -1 Then
        tarefaClicadaCodigo = frm.listTasks.ItemData(frm.listTasks.ListIndex)

        queryUpdateTasks = "UPDATE Tasks SET Status = 'CONCLUIDA' WHERE Codigo = " & tarefaClicadaCodigo
        
        
        If connectBD.State = adStateClosed Then connectBD.Open
        connectBD.Execute queryUpdateTasks

        MsgBox "Tarefa concluída com sucesso!", vbInformation, "Sucesso"
        Call reloadListTasks(frm)

        Exit Function
    Else
    
        MsgBox "Selecione uma tarefa a ser concluida!", vbExclamation, "Aviso"
        
        Exit Function
    End If

erroEndTasks:
    MsgBox "Erro ao concluir tarefa: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function


Public Function deleteTasks(frm As Object)
On Error GoTo erroDeleteTask

     Dim deleteTarefaSelecionada As Integer

     If frm.listTasks.ListIndex = -1 Then
          MsgBox "Nenhuma tarefa selecionada para excluir!", vbExclamation, "Aviso"
          Exit Function
     End If

           deleteTarefaSelecionada = frm.listTasks.ItemData(frm.listTasks.ListIndex)

          queryDeleteTaskSelecionada = "DELETE FROM Tasks WHERE Codigo = " & deleteTarefaSelecionada

          If connectBD.State = adStateClosed Then connectBD.Open

          connectBD.Execute queryDeleteTaskSelecionada
     
          Call reloadListTasks(frm)

          Exit Function

erroDeleteTask:
     MsgBox "Erro ao excluir tarefa: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function

Public Function deleteAllTasks(frm As Object)

     On Error GoTo erroDeleteAllTasks

          queryDeleteAllTasks = "DELETE FROM Tasks WHERE Status = 'Pendente'"

          If connectBD.State = adStateClosed Then connectBD.Open

          connectBD.Execute queryDeleteAllTasks
          
          Call reloadListTasks(frm)
          Exit Function

erroDeleteAllTasks:
     MsgBox "Erro ao excluir todas as tarefas: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function


Public Function historicoConsultarTasks(frm As Object)
    Dim queryHistoricoConsultarTasks As String
    Dim linhaGrid As Integer
    Dim filtro As String

    On Error GoTo erroConsultarHistorico

    filtro = Trim(frm.inputHistoryFilter.Text)

     
    If filtro = "" Then
        'Se o input estiver vazio, carrega todas as tasks
        queryHistoricoConsultarTasks = "SELECT * FROM Tasks WHERE Status IN ('Pendente', 'Concluida') ORDER BY Status, Descricao"
    Else
        'Se o input tiver algum caractere, faz a busca no BD
        queryHistoricoConsultarTasks = "SELECT * FROM Tasks WHERE (Status IN ('Pendente', 'Concluida')) AND (Descricao LIKE '%" & filtro & "%') ORDER BY Status, Descricao"
    End If

    If connectBD.State = adStateClosed Then connectBD.Open

    If recordBD.State = adStateOpen Then recordBD.Close
    recordBD.Open queryHistoricoConsultarTasks, connectBD, adOpenStatic, adLockReadOnly

    frm.GridHistorico.Clear
    frm.GridHistorico.Rows = 1
    frm.GridHistorico.TextMatrix(0, 0) = "Tarefa"
    frm.GridHistorico.TextMatrix(0, 1) = "Status"

    linhaGrid = 1

    While Not recordBD.EOF
        frm.GridHistorico.Rows = frm.GridHistorico.Rows + 1
        frm.GridHistorico.TextMatrix(linhaGrid, 0) = recordBD.Fields("Descricao").Value
        frm.GridHistorico.TextMatrix(linhaGrid, 1) = recordBD.Fields("Status").Value

        recordBD.MoveNext
        linhaGrid = linhaGrid + 1
    Wend

    recordBD.Close
    Exit Function

erroConsultarHistorico:
    MsgBox "Erro ao consultar histórico: " & Err.Number & " - " & Err.Description, vbCritical, "Erro"
End Function


