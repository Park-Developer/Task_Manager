Attribute VB_Name = "Module4"

Option Explicit


Public Sub Lock_Activesheet()

With ActiveSheet
    .Protect Password:=Lock_Password, DrawingObjects:=True, UserInterfaceOnly:=True, Contents:=True
    
    
    
End With


End Sub





Public Sub Unlock_Activesheet()


ActiveSheet.Unprotect Lock_Password

End Sub

