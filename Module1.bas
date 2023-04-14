Attribute VB_Name = "Module1"
'[전역변수 설정]
Option Explicit

'-------------------------------------------------------------------------------------------
Public Const MAX_TASK_NUMBER As Integer = 100

'-------------------------------------------------------------------------------------------
'Task List Range(Max : 100)
Public Const tasklist_name_range As String = "D15:D115" 'Task List Range
Public Const tasklist_state_range As String = "H15:H115" 'Task List Range

'-------------------------------------------------------------------------------------------

Public Const Add_task_name As String = "C11"

Public Const Add_task_due As String = "I11"

Public Total_task_num As Integer

Public Not_started_task_num As Integer
Public In_Progress_task_num As Integer
Public Complete_task_num As Integer


Public Const Add_task_state As String = "H11"


Public Const Add_task_priority As String = "K11"

Public Const Task_Start_Cell As String = "C14"

Public Const Task_Name_Loc As String = "D14"

Public Const task_state_Loc As String = "H14"

Public Const Total_Task_Loc As String = "E5" 'Total Task Number

Public Const Not_stared_loc As String = "H5" 'Not Started Task

Public Const In_progress_loc As String = "H6" 'In Progress Task

Public Const Complete_loc As String = "H7" 'Complete Task


Public Const Not_stared_task As String = "P5" 'Not Started Task

Public Const In_progress_task As String = "U5" 'In Progress Task

Public Const Complete_task As String = "Z5" 'Complete Task

'-------------------------------------------------------------------------------------------

'old
Public Const Timeline_Task_Loc As String = "B5" ' Timeline Sheet Start Loc
Public Const Timeline_Today_Loc As String = "K3" ' Timeline Today Loc

'new
Public Const Timeline_day_Loc As String = "K5" ' Timeline Today Loc
Public Const Timeline_month_Loc As String = "K2" ' Timeline Today Loc

'-------------------------------------------------------------------------------------------

'Public Userform_task_name As String
'Public Userform_task_state As String
'Public Userform_task_due As String
'Public Userform_task_priority As String

Public mod_button_line As Integer

'-------------------------------------------------------------------------------------------

Public Const Lock_Password As String = "123123" 'Wonho_scheduler_V0
'-------------------------------------------------------------------------------------------

Public Task_Collection As New Collection

'-------------------------------------------------------------------------------------------
Public Low_max_num As Integer
Public Normal_max_num As Integer
Public Urgent_max_num As Integer

Public Const Priority_Bar As String = "O5"

Public Const Low_Loc As String = "J5"
Public Const Normal_Loc As String = "J6"
Public Const Urgent_Loc As String = "J7"

'-------------------------------------------------------------------------------------------

Public Const Today_Task_Loc As String = "E6"
Public Const Delayed_Task_Loc As String = "E7"





Public Sub mod_btn_Click()
  

    mod_button_line = CInt(Replace(Application.Caller, "Mod_Btn", ""))
   
   UserForm1.Show
End Sub
Public Sub Del_btn_Click()
 Debug.Print "---------------------------------------------------------------1 "
   ActiveSheet.Shapes(Application.Caller).Delete ' Delete Del Button
 
   Dim btn_name As String
   btn_name = Application.Caller
   
   ActiveSheet.Shapes(Replace(btn_name, "Del", "Mod")).Delete ' Delete Mod Button
   
   
   Dim delete_index As Integer
   delete_index = CInt(Replace(btn_name, "Del_Btn", ""))

    
    ActiveSheet.Range(Task_Start_Cell).offset(delete_index).Resize(, 9).ClearContents

    Dim i As Integer

    If (delete_index < Task_Collection.Count) Then
        '[1] Contents Delete & Modify
        For i = delete_index + 1 To Task_Collection.Count
            Debug.Print "delete2 => " + CStr(i)
            ActiveSheet.Shapes("Mod_Btn" + CStr(i)).Delete ' Delete Del Button
            ActiveSheet.Shapes("Del_Btn" + CStr(i)).Delete ' Delete Del Button
            ActiveSheet.Range(Task_Start_Cell).offset(i).Resize(, 9).ClearContents
            
            Task_Collection(i).index = Task_Collection(i).index - 1
            
        Next
    
    
    
        '[2] Regenerate
        For i = delete_index + 1 To Task_Collection.Count
      
            
            Dim mod_btn As Button
            Dim del_btn As Button
            
            Dim mod_btn_loc As Range
            Set mod_btn_loc = Range(Task_Start_Cell).offset(Task_Collection(i).index, 9)
            
            Set mod_btn = ActiveSheet.Buttons.Add(mod_btn_loc.Left, mod_btn_loc.Top, mod_btn_loc.Width, mod_btn_loc.Height)
            
            With mod_btn
                .OnAction = "mod_btn_Click"
                .Caption = "Mod"
                .name = "Mod_Btn" + CStr(Task_Collection(i).index)
                
            End With
            
            
            
            Dim del_btn_loc As Range
            Set del_btn_loc = Range(Task_Start_Cell).offset(Task_Collection(i).index, 10)
            
            Set del_btn = ActiveSheet.Buttons.Add(del_btn_loc.Left, del_btn_loc.Top, del_btn_loc.Width, del_btn_loc.Height)
        
            With del_btn
                .OnAction = "Del_btn_Click"
                .Caption = "Del"
                .name = "Del_Btn" + CStr(Task_Collection(i).index)
            End With
            
            
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 0).Value = Task_Collection(i).index
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 1).Value = Task_Collection(i).name
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 6).Value = Task_Collection(i).due
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 5).Value = Task_Collection(i).state
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 7).Value = Task_Collection(i).priority
            Range(Task_Start_Cell).offset(Task_Collection(i).index, 8).Value = Task_Collection(i).remain
        Next
    
    End If
    
    
' Delete Collection
 Task_Collection.Remove (delete_index)

Call Copy_TaskData
Call Calc_taskNum

End Sub

Public Sub sort_by_priority()

Dim i As Integer

Dim low_not_start As Integer
Dim low_progress As Integer
Dim low_complete As Integer

low_not_start = 0
low_progress = 0
low_complete = 0

Dim normal_not_start As Integer
Dim normal_progress As Integer
Dim normal_complete As Integer

normal_not_start = 0
normal_progress = 0
normal_complete = 0

Dim urgent_not_start As Integer
Dim urgent_progress As Integer
Dim urgent_complete As Integer

urgent_not_start = 0
urgent_progress = 0
urgent_complete = 0



For i = 1 To Task_Collection.Count
    ' State : Not Started
    If (Task_Collection(i).state = "Not Started") Then
        If (Task_Collection(i).priority = "Urgent") Then
            urgent_not_start = urgent_not_start + 1
        ElseIf (Task_Collection(i).priority = "Normal") Then
            normal_not_start = normal_not_start + 1
        Else ' priority= "Low"
            low_not_start = low_not_start + 1
        End If
        
    ' State : In Progress
    ElseIf (Task_Collection(i).state = "In Progress") Then
        If (Task_Collection(i).priority = "Urgent") Then
            urgent_progress = urgent_progress + 1
        ElseIf (Task_Collection(i).priority = "Normal") Then
            normal_progress = normal_progress + 1
        Else ' priority= "Low"
            low_progress = low_progress + 1
        End If

    ' State : Complete
    ElseIf (Task_Collection(i).state = "Complete") Then
        If (Task_Collection(i).priority = "Urgent") Then
            urgent_complete = urgent_complete + 1
        ElseIf (Task_Collection(i).priority = "Normal") Then
            normal_complete = normal_complete + 1
        Else ' priority= "Low"
            low_complete = low_complete + 1
        End If
    
    End If
Next


Low_max_num = WorksheetFunction.Max(low_not_start, low_progress, low_complete)
Normal_max_num = WorksheetFunction.Max(normal_not_start, normal_progress, normal_complete)
Urgent_max_num = WorksheetFunction.Max(urgent_not_start, urgent_progress, urgent_complete)


ActiveSheet.Range(Low_Loc).Value = Low_max_num
ActiveSheet.Range(Normal_Loc).Value = Normal_max_num
ActiveSheet.Range(Urgent_Loc).Value = Urgent_max_num


'Priority_Bar Draw
ActiveSheet.Range(Priority_Bar).offset(1).Resize(MAX_TASK_NUMBER * 2).ClearContents ' Clear Line
ActiveSheet.Range(Priority_Bar).offset(1).Resize(MAX_TASK_NUMBER * 2).ClearFormats ' Clear Line

ActiveSheet.Range(Priority_Bar).offset(1).Resize(Urgent_max_num).Interior.color = RGB(255, 0, 0)
ActiveSheet.Range(Priority_Bar).offset(Urgent_max_num + 2).Resize(Normal_max_num).Interior.color = RGB(0, 255, 0)
ActiveSheet.Range(Priority_Bar).offset(Urgent_max_num + 3 + Normal_max_num).Resize(Low_max_num).Interior.color = RGB(0, 0, 255)


End Sub
Public Sub Copy_TaskData()

' Sort With Priority
Call sort_by_priority '

'Clear All Contents
ActiveSheet.Range(Not_stared_task).offset(1).Resize(MAX_TASK_NUMBER * 2, 400).ClearContents ' Clear Line
ActiveSheet.Range(In_progress_task).offset(1).Resize(MAX_TASK_NUMBER * 2, 400).ClearContents ' Clear Line
ActiveSheet.Range(Not_stared_task).offset(1).Resize(MAX_TASK_NUMBER * 2, 400).ClearContents ' Clear Line

Dim i As Integer

Dim not_start_low As Integer
Dim not_start_normal As Integer
Dim not_start_urgent As Integer

not_start_low = 0
not_start_normal = 0
not_start_urgent = 0

Dim inprogress_low As Integer
Dim inprogress_normal As Integer
Dim inprogress_urgent As Integer

inprogress_low = 0
inprogress_normal = 0
inprogress_urgent = 0

Dim complete_low As Integer
Dim complete_normal As Integer
Dim complete_urgent As Integer

complete_low = 0
complete_normal = 0
complete_urgent = 0


For i = 1 To Task_Collection.Count

    ' (1) ---------------------- Not Started ----------------------
    If (Task_Collection(i).state = "Not Started") Then

        ' (1-1) Urgent
        If (Task_Collection(i).priority = "Urgent") Then
            not_start_urgent = not_start_urgent + 1
            ActiveSheet.Range(Not_stared_task).offset(not_start_urgent).Value = Task_Collection(i).name
            
            ' font setting
            ActiveSheet.Range(Not_stared_task).offset(not_start_urgent).Font.color = set_Fontcolor(Task_Collection(i).due)
        
        ' (1-2) Normal
        ElseIf (Task_Collection(i).priority = "Normal") Then
            not_start_normal = not_start_normal + 1
            ActiveSheet.Range(Not_stared_task).offset(Urgent_max_num + 1 + not_start_normal).Value = Task_Collection(i).name
        
            ' font setting
            ActiveSheet.Range(Not_stared_task).offset(Urgent_max_num + 1 + not_start_normal).Font.color = set_Fontcolor(Task_Collection(i).due)
            
        ' (1-3) Low
        Else ' priority= "Low"
            not_start_low = not_start_low + 1
            ActiveSheet.Range(Not_stared_task).offset(Urgent_max_num + Normal_max_num + 2 + not_start_low).Value = Task_Collection(i).name
            
            ' font setting
            ActiveSheet.Range(Not_stared_task).offset(Urgent_max_num + Normal_max_num + 2 + not_start_low).Font.color = set_Fontcolor(Task_Collection(i).due)
            
        End If
    
        
    ' (2) ---------------------- In_progress_task ----------------------
    ElseIf (Task_Collection(i).state = "In Progress") Then

        ' (2-1) Urgent
        If (Task_Collection(i).priority = "Urgent") Then
            inprogress_urgent = inprogress_urgent + 1
            ActiveSheet.Range(In_progress_task).offset(inprogress_urgent).Value = Task_Collection(i).name
        
            ' font setting
            ActiveSheet.Range(In_progress_task).offset(inprogress_urgent).Font.color = set_Fontcolor(Task_Collection(i).due)
        ' (2-2) Normal
        ElseIf (Task_Collection(i).priority = "Normal") Then
            inprogress_normal = inprogress_normal + 1
            ActiveSheet.Range(In_progress_task).offset(Urgent_max_num + 1 + inprogress_normal).Value = Task_Collection(i).name
        
            ' font setting
            ActiveSheet.Range(In_progress_task).offset(Urgent_max_num + 1 + inprogress_normal).Font.color = set_Fontcolor(Task_Collection(i).due)
        ' (2-3) Low
        Else ' priority= "Low"
            not_start_low = not_start_low + 1
            ActiveSheet.Range(In_progress_task).offset(Urgent_max_num + Normal_max_num + 2 + inprogress_low).Value = Task_Collection(i).name
            
            ' font setting
            ActiveSheet.Range(In_progress_task).offset(Urgent_max_num + Normal_max_num + 2 + inprogress_low).Font.color = set_Fontcolor(Task_Collection(i).due)
        End If

        
    ' (3) ---------------------- Complete_task ----------------------
    ElseIf (Task_Collection(i).state = "Complete") Then
   
        ' (3-1) Urgent
        If (Task_Collection(i).priority = "Urgent") Then
            complete_urgent = complete_urgent + 1
            ActiveSheet.Range(Complete_task).offset(complete_urgent).Value = Task_Collection(i).name
        
            ' font setting
            ActiveSheet.Range(Complete_task).offset(complete_urgent).Font.color = set_Fontcolor(Task_Collection(i).due)
        ' (3-2) Normal
        ElseIf (Task_Collection(i).priority = "Normal") Then
            complete_normal = complete_normal + 1
            ActiveSheet.Range(Complete_task).offset(Urgent_max_num + 1 + complete_normal).Value = Task_Collection(i).name
        
            ' font setting
            ActiveSheet.Range(Complete_task).offset(Urgent_max_num + 1 + complete_normal).Font.color = set_Fontcolor(Task_Collection(i).due)
        ' (3-3) Low
        Else ' priority= "Low"
            complete_low = complete_low + 1
            ActiveSheet.Range(Complete_task).offset(Urgent_max_num + Normal_max_num + 2 + complete_low).Value = Task_Collection(i).name
            
            ' font setting
            ActiveSheet.Range(Complete_task).offset(Urgent_max_num + Normal_max_num + 2 + complete_low).Font.color = set_Fontcolor(Task_Collection(i).due)
        End If
    
    End If
    
Next
    
    
End Sub

Public Sub find_todayTask()

Dim i As Integer

Dim today_task_num As Integer
today_task_num = 0

Dim delayed_task_num As Integer
delayed_task_num = 0

For i = 1 To Task_Collection.Count
    '(1) find today task
    If (Task_Collection(i).due = Date) Then
        today_task_num = today_task_num + 1
        
    End If
    
    '(2) find delayed task
    If (DateDiff("d", Date, Task_Collection(i).due) < 0) Then
        delayed_task_num = delayed_task_num + 1
    End If
Next


'Today Task Show
ActiveSheet.Range(Today_Task_Loc).Value = today_task_num
ActiveSheet.Range(Today_Task_Loc).Font.color = RGB(255, 128, 0) ' Orange
ActiveSheet.Range(Today_Task_Loc).HorizontalAlignment = xlCenter


'Delayed Task Show
ActiveSheet.Range(Delayed_Task_Loc).Value = delayed_task_num
ActiveSheet.Range(Delayed_Task_Loc).Font.color = RGB(255, 0, 0) ' Red
ActiveSheet.Range(Delayed_Task_Loc).HorizontalAlignment = xlCenter


End Sub


Public Function Get_TaskList_taskInfo(task_property As String, offset As Integer)

If (task_property = "State") Then
    Get_TaskList_taskInfo = ActiveSheet.Range(task_state_Loc).offset(offset).Value
ElseIf (task_property = "Name") Then
    Get_TaskList_taskInfo = ActiveSheet.Range(Task_Name_Loc).offset(offset).Value
End If

End Function


