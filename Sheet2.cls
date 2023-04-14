VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Write_Calaneder()

Worksheets("Timeline").Activate 'Activate

ActiveSheet.Range(Timeline_Today_Loc).offset(1).Resize(2, 366).ClearFormats ' Clear Format
ActiveSheet.Range(Timeline_Today_Loc).offset(1).Resize(2, 366).UnMerge
ActiveSheet.Range(Timeline_Today_Loc).offset(1).Resize(2, 366).ClearContents ' Clear Line

Dim today_day As Integer
today_day = CInt(day(Now))

Dim today_month As Integer
today_month = CInt(month(Now))

Dim merge_start As Integer
Dim merge_end As Integer

 
'Write  Today Month & Today Day(day)
ActiveSheet.Range(Timeline_Today_Loc).offset(1).Value = today_month

ActiveSheet.Range(Timeline_Today_Loc).offset(2).Value = today_day

Dim this_month_day As Integer
this_month_day = day(Application.WorksheetFunction.EoMonth(Now, 0))

Dim month_loop As Integer
month_loop = 12 - today_month

Dim day_loop As Integer


Dim day_offset, month_offset As Integer

day_offset = 0 ' default day offset value


' This Month
Dim day_cnt, month_cnt As Integer

' Write This Month Day
For day_cnt = today_day To this_month_day
    ActiveSheet.Range(Timeline_day_Loc).offset(0, day_offset).Value = day_cnt
    ActiveSheet.Range(Timeline_day_Loc).offset(0, day_offset).HorizontalAlignment = xlCenter
    day_offset = day_offset + 1
Next

ActiveSheet.Range(Timeline_day_Loc).Resize(, day_offset).Interior.color = RGB(192, 192, 192)

'Write This Month
ActiveSheet.Range(Timeline_Today_Loc).offset(0).Resize(2, day_offset).Merge
ActiveSheet.Range(Timeline_Today_Loc).offset(0).Resize(, day_offset).Interior.color = RGB(52, 153, 255)
ActiveSheet.Range(Timeline_Today_Loc).offset(0).Resize(, day_offset).Font.Size = 24


ActiveSheet.Range(Timeline_Today_Loc).offset(0).Resize(, day_offset).HorizontalAlignment = xlCenter
' Loop Month
Dim loop_cnt As Integer
Dim month_start_point As Integer

For month_cnt = today_month + 1 To 12
   
   merge_start = day_offset  'merge start
   


    day_cnt = 1
    
    If (day_cnt = 1) Then
        ' Write Month
        ActiveSheet.Range(Timeline_month_Loc).offset(1, day_offset).Value = month_cnt
        ActiveSheet.Range(Timeline_month_Loc).offset(1, day_offset).Font.Size = 24
        month_start_point = day_offset
        
    Else
    
    End If
    
    day_loop = day(Application.WorksheetFunction.EoMonth(Now, month_cnt - today_month))
    loop_cnt = 0
    
    For day_cnt = 1 To day_loop
        ActiveSheet.Range(Timeline_day_Loc).offset(, day_offset).Value = day_cnt
        ActiveSheet.Range(Timeline_day_Loc).offset(, day_offset).HorizontalAlignment = xlCenter
        day_offset = day_offset + 1
        loop_cnt = loop_cnt + 1
        
    Next
    
    merge_end = day_offset - 1 'merge end
        
        
    'Write This Month
    ActiveSheet.Range(Timeline_month_Loc).offset(1, month_start_point).Resize(2, loop_cnt).Merge
    ActiveSheet.Range(Timeline_month_Loc).offset(1, month_start_point).Resize(2, loop_cnt).HorizontalAlignment = xlCenter
    ActiveSheet.Range(Timeline_month_Loc).offset(1, month_start_point).Resize(2, loop_cnt).Interior.color = RGB(52, 153, 255)
    
    ActiveSheet.Range(Timeline_day_Loc).offset(0, month_start_point).Resize(, loop_cnt).Interior.color = RGB(192, 192, 192)
Next


End Sub



Private Sub Worksheet_Activate()
Call Write_Calaneder

Call Copy_Timeline

End Sub
