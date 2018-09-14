VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimelines 
   Caption         =   "Select Start and End Date"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6165
   OleObjectBlob   =   "frmTimelines.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimelines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmTimelines.Hide
End Sub

Private Sub cmdOk_Click()
    
    On Error GoTo err_cmdOk_Click
    
    Dim oCurrentTask As Task
    Dim oFirstTask As Task
    Dim dtStartDate As Date, dtEndDate As Date
    
    dtStartDate = mvStartDate.Value
    dtEndDate = mvEndDate.Value
    
    For Each oCurrentTask In ActiveProject.Tasks
      If Not (oCurrentTask Is Nothing) Then
        
        'Checking if the Start & Finish dates are in the selected range & if the Duration is 0.
        If (Format(oCurrentTask.Start, "mm/dd/yyyy") >= dtStartDate And Format(oCurrentTask.Finish, "mm/dd/yyyy") <= dtEndDate) _
          And oCurrentTask.Duration = 0 Then
            
            SelectRow Row:=oCurrentTask.ID, RowRelative:=False      'Selects the row.
            TaskOnTimeline                                          'Adds the selected row to the timeline.
            
        ElseIf chkRemoveExistingTasksFromTimeline.Value = True Then
            SelectRow Row:=oCurrentTask.ID, RowRelative:=False              'Selects the row.
            TaskOnTimeline Remove:=True                                     'Removes the row from the timelines
            
        End If
      
      End If
    Next oCurrentTask
    
    'Hiding the selection Form
    frmTimelines.Hide
    
    'Exiting the sub procedure to avoid below code getting executed.
    Exit Sub
    
err_cmdOk_Click:
    MsgBox "Error Description: " & Err.Description & vbCrLf & _
        "Error Number: " & Err.Number
End Sub
