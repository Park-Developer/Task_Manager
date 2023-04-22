Attribute VB_Name = "Module3"
Option Explicit

Public Function get_weekday(day_num As Integer)

Select Case (day_num)
    Case 1: get_weekday = "Sun"
    Case 2: get_weekday = "Mon"
    Case 3: get_weekday = "Tue"
    Case 4: get_weekday = "Wed"
    Case 5: get_weekday = "Thu"
    Case 6: get_weekday = "Fri"
    Case 7: get_weekday = "Sat"
End Select

End Function
Public Sub sort_weekTask()
' This_week_Loc



'(1) erase func

ActiveSheet.Range(This_week_Loc).Resize(200, 2).UnMerge
ActiveSheet.Range(This_week_Loc).Resize(200, 2).ClearFormats
ActiveSheet.Range(This_week_Loc).Resize(200, 2).ClearContents
    
Debug.Print get_weekday(3)
Debug.Print DateAdd("d", 1, Date)

Dim loop_cnt As Integer
Dim start_day As Integer

start_day = Weekday(Date)
loop_cnt = 8 - start_day


Dim i As Integer
Dim j As Integer

Dim t As Integer ' task count

Dim day_offset As Integer
day_offset = 0

Dim task_date As Date


'(2) View This Week Task
For i = 1 To loop_cnt
    task_date = DateAdd("d", i - 1, Date)
    ActiveSheet.Range(This_week_Loc).offset(day_offset).Value = CStr(get_weekday(Weekday(task_date))) + "(" + CStr(day(task_date)) + ")" ' 夸老
   ActiveSheet.Range(This_week_Loc).offset(day_offset).Font.Bold = True
    ActiveSheet.Range(This_week_Loc).offset(day_offset).Resize(, 2).Interior.color = RGB(224, 224, 224)
    t = 1 'task cnt ini
    
    For j = 1 To Task_Collection.Count
        
        
        If Task_Collection(j).due = task_date Then
            ActiveSheet.Range(This_week_Loc).offset(day_offset + t).Value = Task_Collection(j).name  ' 夸老
            
            t = t + 1
            
        End If
        
    Next
    
    day_offset = day_offset + t
Next


'(3) Divider
ActiveSheet.Range(This_week_Loc).offset(day_offset).Value = "Next Week"
ActiveSheet.Range(This_week_Loc).offset(day_offset).Resize(, 2).Merge
ActiveSheet.Range(This_week_Loc).offset(day_offset).Font.color = RGB(0, 0, 255)
ActiveSheet.Range(This_week_Loc).offset(day_offset).Font.Bold = True
ActiveSheet.Range(This_week_Loc).offset(day_offset).HorizontalAlignment = xlCenter
ActiveSheet.Range(This_week_Loc).offset(day_offset).Interior.color = RGB(255, 255, 102)
day_offset = day_offset + 1
    

'(4) View Next Week Task
For i = 1 To 7
 task_date = DateAdd("d", i - 1 + loop_cnt, Date)
    ActiveSheet.Range(This_week_Loc).offset(day_offset).Value = CStr(get_weekday(Weekday(task_date))) + "(" + CStr(day(task_date)) + ")" ' 夸老
ActiveSheet.Range(This_week_Loc).offset(day_offset).Font.Bold = True
     ActiveSheet.Range(This_week_Loc).offset(day_offset).Resize(, 2).Interior.color = RGB(224, 224, 224)
    t = 1 'task cnt ini
    
    For j = 1 To Task_Collection.Count
       
        
        If Task_Collection(j).due = task_date Then
            ActiveSheet.Range(This_week_Loc).offset(day_offset + t).Value = Task_Collection(j).name  ' 夸老
            
            t = t + 1
            
        End If
        
    Next
    
    day_offset = day_offset + t
Next


End Sub
