Attribute VB_Name = "Module2"
Option Explicit


Function get_gantt_color(task_number As Integer)

Select Case task_number Mod 7
    Case 1: get_gantt_color = RGB(0, 0, 255) ' blue
    Case 2: get_gantt_color = RGB(0, 255, 0) 'lime
    Case 3: get_gantt_color = RGB(255, 0, 0) 'Red
    Case 4: get_gantt_color = RGB(255, 255, 0) 'Yellow
    Case 5: get_gantt_color = RGB(255, 0, 255) 'Magneta
    Case 6: get_gantt_color = RGB(128, 0, 128) 'Purple
    Case 7: get_gantt_color = RGB(0, 0, 128) 'dark blue
    

End Select

End Function

Public Function state_check(state As String)
    If (state <> "Not Started" And state <> "In Progress" And state <> "Complete") Then
        state_check = "Not Started"
    
    Else
        state_check = state
    End If
    
End Function

Public Function remain_check(remain As String)


    
End Function
Public Function name_check(name As String) ' task name 중복 및 공백 여부 확인

If (name = "") Then '입력값이 없는 경우
    name_check = False
Else
    Dim i As Integer
    
        For i = 1 To Task_Collection.Count
            
            If (Task_Collection(i).name = name) Then
                name_check = False
                Exit Function
                
            End If
      
        Next
        
        name_check = True


End If

End Function
Function priority_check(priority As String)

    If (priority <> "Urgent" And priority <> "Normal" And priority <> "Low") Then
        priority_check = "Low"
    
    Else
        priority_check = priority
    End If
    
End Function


Public Function Check_dataForm(input_data As String)
If (IsEmpty(input_data) = True) Then
    Check_dataForm = True

Else
    Dim num As Integer
    
    num = Len(input_data) - Len(Replace(input_data, "-", ""))
    
    If (num = 2) Then
    Check_dataForm = True
    Else
    Check_dataForm = False
    End If

End If


End Function


Public Function change_color_withPriority(priority As String)
    If (priority = "Urgent") Then
        change_color_withPriority = RGB(255, 0, 0)
    ElseIf (priority = "Normal") Then
        change_color_withPriority = RGB(0, 255, 0)
    Else ' priority= "Low"
        change_color_withPriority = RGB(0, 0, 255)
    End If


End Function

Public Function set_Fontcolor(due As String)

    If (due = Date) Then
       set_Fontcolor = RGB(255, 128, 0) ' Orange
        
    ElseIf (DateDiff("d", Date, due) < 0) Then
        set_Fontcolor = RGB(255, 0, 0)  ' Red
    Else
        set_Fontcolor = RGB(0, 0, 0) 'black
    
    End If
    
End Function
