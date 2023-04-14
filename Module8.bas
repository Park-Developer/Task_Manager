Attribute VB_Name = "Module8"
Option Explicit

Public Sub reset()

Dim i As Integer

For i = 1 To Task_Collection.Count
     Task_Collection.Remove (i)
      ActiveSheet.Range(Task_Start_Cell).offset(i).Resize(, 9).ClearContents
Next
    

   
   


For i = 1 To 10

      ActiveSheet.Range(Task_Start_Cell).offset(i).Resize(, 9).ClearContents

Next
    

   
   
   
End Sub

