Attribute VB_Name = "Module6"
Option Explicit


Public Sub get_Task(name As String, state As String, due As String, priority As String, remain As String, mod_btn As Button, del_btn As Button, index As Integer)

Dim new_task As New Task ' Task Class Definition


With new_task
    .name = name
    .state = state
    .due = due
    .priority = priority
    .remain = remain
    .index = index
End With

Set new_task.mod_btn = mod_btn
Set new_task.del_btn = del_btn

Task_Collection.Add new_task ' get to collection

End Sub

Public Sub del_Task(name As String)

Dim i As Integer

For i = 1 To Task_Collection.Count
    If (Task_Collection(i).name = name) Then
        Task_Collection.Remove (i)
        
        Exit For
    End If
Next

End Sub

Public Function get_state_number()
'Public Const tasklist_name_range As String = "D15:D115" 'Task List Range
Dim i As Integer

Dim not_started_num As Integer
Dim in_progress_num As Integer
Dim complete_num As Integer

not_started_num = 0
in_progress_num = 0
complete_num = 0

Dim result As New Collection:


For i = 1 To Task_Collection.Count
    If (Task_Collection(i).state = "Not Started") Then
        not_started_num = not_started_num + 1
        
    ElseIf (Task_Collection(i).state = "In Progress") Then
        in_progress_num = in_progress_num + 1
       
    ElseIf (Task_Collection(i).state = "Complete") Then
        complete_num = complete_num + 1
    Else
        not_started_num = not_started_num + 1
    End If
Next


result.Add not_started_num, Key:="Not Started"
result.Add in_progress_num, Key:="In Progress"
result.Add complete_num, Key:="Complete"

Set get_state_number = result

End Function

