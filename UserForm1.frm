VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '家蜡磊 啊款单
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public Userform_task_name As String
'Public Userform_task_state As String
'Public Userform_task_due As String
'Public Userform_task_priority As String
'  Worksheets("Board").
Private Sub CommandButton1_Click() ' User info OK Button Click

'(1) Task Name Check

If (name_check(TextBox1.Text) = False) Then

    MsgBox "[Error1] Invalid Task name!"
    Unload UserForm1
    Exit Sub
End If


'Due Format Check
If (Check_dataForm(TextBox2.Text) = True) Then

    '(1) Name
    Task_Collection(mod_button_line).name = TextBox1.Text
    Range(Task_Start_Cell).offset(mod_button_line, 1).Value = TextBox1.Text
    
    
    '(2) State
    Task_Collection(mod_button_line).state = state_check(ComboBox2.Value)  ' state check
    Range(Task_Start_Cell).offset(mod_button_line, 5).Value = state_check(ComboBox2.Value) ' state check
    
    
    '(3) Due
    If (IsEmpty(TextBox2.Text) = True) Then
        Task_Collection(mod_button_line).due = Date ' Default 利侩
        Range(Task_Start_Cell).offset(mod_button_line, 6).Value = Date
    Else
        Task_Collection(mod_button_line).due = TextBox2.Text
        Range(Task_Start_Cell).offset(mod_button_line, 6).Value = TextBox2.Text
    End If
    
    
    '(4) Remain
    Task_Collection(mod_button_line).remain = DateDiff("d", Date, TextBox2.Text)
    Range(Task_Start_Cell).offset(mod_button_line, 8).Value = DateDiff("d", Date, TextBox2.Text)
    
    
    '(5) Priority
    Task_Collection(mod_button_line).priority = priority_check(ComboBox1.Value)
    Range(Task_Start_Cell).offset(mod_button_line, 7).Value = ComboBox1.Value

    
    Call Copy_TaskData ' copy data to Work Progress
    
    Unload UserForm1
    
Else
    MsgBox "[Error2] Invalid Task Due!"
    Unload UserForm1
    Exit Sub
End If



End Sub

Private Sub CommandButton2_Click()
'Cancel

Unload UserForm1

End Sub

Private Sub TextBox1_Change()
'task name

End Sub

Private Sub TextBox2_Change()
'task due

End Sub

' This is the Initialize event procedure for UserForm1
Private Sub UserForm_Initialize()
MsgBox "asd"


' Task State Setting
ComboBox2.List = Array("Not Started", "In Progress", "Complete")

' Task Priority Setting
ComboBox1.List = Array("Urgent", "Normal", "Low")


End Sub
