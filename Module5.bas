Attribute VB_Name = "Module5"
Option Explicit

Public Sub Calc_taskNum()

Call Unlock_Activesheet ' Unlock Active Sheet

Dim state_numInfo As New Collection:

Set state_numInfo = get_state_number() ' Get Each State Number

Not_started_task_num = state_numInfo("Not Started")
In_Progress_task_num = state_numInfo("In Progress")
Complete_task_num = state_numInfo("Complete")

Total_task_num = Not_started_task_num + In_Progress_task_num + Complete_task_num



' Display Count Number

ActiveSheet.Range(Total_Task_Loc).Value = Total_task_num
    
ActiveSheet.Range(Not_stared_loc).Value = Not_started_task_num
ActiveSheet.Range(In_progress_loc).Value = In_Progress_task_num
ActiveSheet.Range(Complete_loc).Value = Complete_task_num

Call Lock_Activesheet ' Lock Active Sheet

End Sub

Public Sub Copy_Timeline()

Dim i As Integer

Dim line_number As Integer

'Delete All Contents
ActiveSheet.Range(Timeline_Task_Loc).offset(1).Resize(MAX_TASK_NUMBER * 2, 400).ClearContents ' Clear Line
ActiveSheet.Range(Timeline_Task_Loc).offset(1).Resize(MAX_TASK_NUMBER * 2, 400).Interior.color = xlNone

For i = 1 To Task_Collection.Count
    line_number = i * 2 - 1 ' ÇÑÄ­¾¿ ¶ç¿ö¼­ »öÄ¥


    ' Timeline Task Name : Offset : 0
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 0).Value = Task_Collection(i).name
    ' set font color
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 0).Font.color = set_Fontcolor(Task_Collection(i).due)
    
    ' Timeline Task State : Offset : 4
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 4).Value = Task_Collection(i).state
    
    ' Timeline Task Priority : Offset : 5
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 5).Value = Task_Collection(i).priority
    
    ' Timeline Task Due : Offset : 6
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 6).Value = Task_Collection(i).due
    ' set font align
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 6).HorizontalAlignment = xlCenter
    
    ' Timeline Task Remain : Offset : 7
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 7).Value = Task_Collection(i).remain
    ActiveSheet.Range(Timeline_Task_Loc).offset(line_number, 7).HorizontalAlignment = xlCenter

    ActiveSheet.Range(Timeline_Today_Loc).offset(line_number + 1).Resize(, CInt(Task_Collection(i).remain) + 1).Interior.color = get_gantt_color(i)
    



Next


End Sub
