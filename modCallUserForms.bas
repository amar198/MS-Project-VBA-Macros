Attribute VB_Name = "modCallUserForms"
Public Sub Call_AddToTimeline()
        
    On Error GoTo err_AddToTimeline
    
    frmTimelines.Show
    
    Exit Sub
    
err_AddToTimeline:
    MsgBox "Error Description: " & Err.Description & vbCrLf & _
        "Error Number: " & Err.Number
End Sub

