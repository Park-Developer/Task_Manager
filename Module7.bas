Attribute VB_Name = "Module7"
Option Explicit
Sub sad()


Dim new_task As New Task ' Task Class Definition


With new_task
    .name = "name1"
    .state = "state1"
    .due = "djue11"
    .priority = "proiroty"
End With



Task_Collection.Add new_task ' get to collectio

With new_task
    .name = "name2"
    .state = "stat3e1"
    .due = "djue141"
    .priority = "p3roiroty"
End With



Task_Collection.Add new_task ' get to collectio





 End Sub
 
 
 

