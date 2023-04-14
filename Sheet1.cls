VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub ComboBox1_Change() ' Task State

Range(Add_task_state).Value = ComboBox1.Value

End Sub

Private Sub ComboBox2_Change()  ' Task Priority

Range(Add_task_priority).Value = ComboBox2.Value
    
End Sub

Private Sub CommandButton1_Click() 'Add

'(1) Unlock Sheet
Call Unlock_Activesheet

'(2) Task Name Check
Debug.Print "name_c" + CStr(name_check(Range(Add_task_name).Value))
Debug.Print "IsEmpty(name)" + CStr(IsEmpty(Range(Add_task_name).Value))
Debug.Print "Range(Add_task_name).Value" + CStr(Range(Add_task_name).Value)

If (name_check(Range(Add_task_name).Value) = False) Then
        MsgBox "[Error1] Invalid Task Name!"
        Exit Sub
End If

Dim new_task_name As String
new_task_name = Range(Add_task_name).Value 'Offset 1 : Task Name


'(3) Task Due Check
Dim new_task_due As String

If (IsEmpty(Range(Add_task_due).Value) = True) Then 'Offset 5 : Due Date
    new_task_due = Date ' default

Else
    If (Check_dataForm(Range(Add_task_due).Value) = True) Then
        new_task_due = Range(Add_task_due).Value
    
    Else
         MsgBox "[Error2] Invalid Task Due!"
    
        Exit Sub
    End If
End If


'Offset 4 : Task State
Dim new_task_state As String
new_task_state = state_check(Range(Add_task_state).Value)


'Offset 6 : Priority
Dim new_task_priority As String
new_task_priority = priority_check(Range(Add_task_priority).Value)


'Offset 7 : Remain
Dim new_task_ramain As String
new_task_remain = DateDiff("d", Date, new_task_due)

'Offset 8 : Modiify Button
Dim mod_btn As Button

Dim mod_btn_loc As Range
Set mod_btn_loc = Range(Task_Start_Cell).offset(Task_Collection.Count + 1, 9)

Set mod_btn = ActiveSheet.Buttons.Add(mod_btn_loc.Left, mod_btn_loc.Top, mod_btn_loc.Width, mod_btn_loc.Height)

With mod_btn
    .OnAction = "mod_btn_Click"
    .Caption = "Mod"
    .name = "Mod_Btn" + CStr(Task_Collection.Count + 1)
    
End With

'Offset 9 : Delete Button
Dim del_btn As Button

Dim del_btn_loc As Range
Set del_btn_loc = Range(Task_Start_Cell).offset(Task_Collection.Count + 1, 10)

Set del_btn = ActiveSheet.Buttons.Add(del_btn_loc.Left, del_btn_loc.Top, del_btn_loc.Width, del_btn_loc.Height)

With del_btn
    .OnAction = "Del_btn_Click"
    .Caption = "Del"
    .name = "Del_Btn" + CStr(Task_Collection.Count + 1)
End With


Call get_Task(CStr(new_task_name), CStr(new_task_state), CStr(new_task_due), CStr(new_task_priority), CStr(new_task_remain), mod_btn, del_btn, Task_Collection.Count + 1)

' Number Update

Call Calc_taskNum


'Write Task Info
Range(Task_Start_Cell).offset(Total_task_num, 0).Value = Total_task_num
Range(Task_Start_Cell).offset(Total_task_num, 0).HorizontalAlignment = xlCenter
Range(Task_Start_Cell).offset(Total_task_num, 0).Font.Bold = True

Range(Task_Start_Cell).offset(Total_task_num, 1).Value = new_task_name

Range(Task_Start_Cell).offset(Total_task_num, 6).Value = new_task_due
Range(Task_Start_Cell).offset(Total_task_num, 6).HorizontalAlignment = xlCenter

Range(Task_Start_Cell).offset(Total_task_num, 5).Value = new_task_state
Range(Task_Start_Cell).offset(Total_task_num, 7).Value = new_task_priority
Range(Task_Start_Cell).offset(Total_task_num, 8).Value = new_task_remain

'Copy Task
Call Copy_TaskData

' ComboBox Clear
ComboBox1.Value = ""
ComboBox2.Value = ""
Range(Add_task_name).Value = ""
Range(Add_task_due).Value = ""

' Lock Active Sheet
Call Lock_Activesheet

End Sub
Private Sub task_info_init() ' 고민을 해보자!!!!!!!!!!!!!!
'Public Const Task_Start_Cell As String = "C14"
' MAX_TASK_NUMBER

Dim recoreded_task_num As Integer
recoreded_task_num = WorksheetFunction.CountA(Range(Task_Start_Cell).offset(1).Resize(MAX_TASK_NUMBER))

Dim i As Integer


Dim name As String
Dim state As String
Dim due As String
Dim remain As String
Dim priority As String
Dim index As Integer

Dim mod_btn As Button
Dim del_btn As Button

'If (recoreded_task_num <> Task_Collection.Count) Then
    '(1) Reset Task_Collection
    Debug.Print "recoreded_task_num" + CStr(recoreded_task_num)
    
    If Task_Collection.Count > 0 Then
        Debug.Print "Task_Collection.Countttt" + CStr(Task_Collection.Count)
        
        While Task_Collection.Count > 0
            Task_Collection.Remove (1)
        Wend
        
    End If
    
    
    '(2) Regenerate Task_Collection referencing board sheet
    For i = 1 To recoreded_task_num
        index = Range(Task_Start_Cell).offset(i, 0).Value
        name = Range(Task_Start_Cell).offset(i, 1).Value
        due = Range(Task_Start_Cell).offset(i, 6).Value
        state = state_check(Range(Task_Start_Cell).offset(i, 5).Value)
        priority = priority_check(Range(Task_Start_Cell).offset(i, 7).Value)
                
        'remain check
        If IsEmpty(Range(Task_Start_Cell).offset(i, 8).Value) = True Then
            If (due = "") Then
                remain = "1"
            Else
                remain = DateDiff("d", Date, due)
            End If
            
            
            
        Else
            remain = Range(Task_Start_Cell).offset(i, 8).Value
        End If



        Set mod_btn = ActiveSheet.Buttons("Mod_Btn" + CStr(i))
        Set del_btn = ActiveSheet.Buttons("Del_Btn" + CStr(i))


        Call get_Task(name, state, due, priority, remain, mod_btn, del_btn, index)
    Next
    
'End If

Debug.Print "Task_Collection.Coun1t" + CStr(Task_Collection.Count)
    
    
Call Calc_taskNum
Call Copy_TaskData

End Sub

Private Sub CommandButton2_Click() ' Setting button


End Sub

Private Sub Worksheet_Activate()
'[Combobox Setting]
ComboBox1.List = Array("Not Started", "In Progress", "Complete")
ComboBox2.List = Array("Urgent", "Normal", "Low")

Call task_info_init

'[Task Config]
Call Calc_taskNum
Call Copy_TaskData

'[Today Task]
Call find_todayTask

' Lock Active Sheet
Call Lock_Activesheet


End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

End Sub
