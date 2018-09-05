Attribute VB_Name = "modPERTFunctionality"
Private Const OptimisticDuration As Long = 188743694
Private Const MostLikely As Long = 188744967
Private Const Pessimistic As Long = 188744965

Dim oOrganisationDetails As New clsOrganizationDetails

Public Sub PERT()
    Dim oCurrentTask As Task
    Dim oFirstTask As Task
    Dim FoundBadWeights As Boolean
    Dim UseDefaultWeights As Boolean
    Dim dOptimistic As Double, dPessimistic As Double, dMostLikely As Double, dTotalWeight As Double, iCtr As Integer
    Dim dStandardDeviation As Double, dTotalProjectVariance As Double
    
    UseDefaultWeights = True
    FoundBadWeights = False
    
    iCtr = 0
    dTotalProjectVariance = 0
    
        
    Set oFirstTask = ActiveProject.Tasks(1)
    For Each oCurrentTask In ActiveProject.Tasks
      If Not (oCurrentTask Is Nothing) Then
        If oCurrentTask.PercentComplete = 0 And oCurrentTask.PercentWorkComplete = 0 Then
            'oCurrentTask.Duration1 is Optimistic Duration
            'oCurrentTask.Duration2 is Most Likely Duration
            'oCurrentTask.Duration3 is Pessimistic Duration
            
            ''Getting the weights assigned to the task.
            ''If no weights are assigned, then default weights will be taken.
            ''' Default Weight for Optimistic=1, Most Likely=4, and Pessimistic=1
            
            'Getting Optimistic, Most Likely, & Pessimistic Weight
            If Trim$(Str$(oCurrentTask.Number1)) <> "0" Then
                dOptimistic = oCurrentTask.Number1
            Else
                dOptimistic = 1
            End If
            
            If Trim$(Str$(oCurrentTask.Number2)) <> "0" Then
                dMostLikely = oCurrentTask.Number2
            Else
                dMostLikely = 4
            End If
            
            If Trim$(Str$(oCurrentTask.Number3)) <> "0" Then
                dPessimistic = oCurrentTask.Number3
            Else
                dPessimistic = 1
            End If
            
            ''Getting total of weights for calculating Estimated Activity Duration (EAD) & Standard Deviation
            dTotalWeight = dOptimistic + dMostLikely + dPessimistic
            
            
            'Implementing PERT formula
            ' (O + 4M + P) / 6
            oCurrentTask.Duration = _
                ((oCurrentTask.Duration1 * dOptimistic) + (oCurrentTask.Duration2 * dMostLikely) + (oCurrentTask.Duration3 * dPessimistic)) / dTotalWeight
                
            'oCurrentTask.Name.Select
            
'                If Right$(Str$(oCurrentTask.Duration), 1) = "?" Then
'                    oCurrentTask.Duration = Mid$(Str$(oCurrentTask.Duration), 1, Len(Str$(oCurrentTask.Duration)) - 1)
'                End If
                
            'Implementing the standard deviation formula
            '(P - O) / 6
            oCurrentTask.Duration4 = (oCurrentTask.Duration3 - oCurrentTask.Duration1) / dTotalWeight
            
            'Calculating the Variance
            'SquareOf((P - O)/6) or Std. Dev. Square
            dStandardDeviation = oCurrentTask.Duration4
            oCurrentTask.Duration5 = (dStandardDeviation * dStandardDeviation)
            
            If oCurrentTask.Duration1 > 0 Then
                iCtr = iCtr + 1
                dTotalProjectVariance = dTotalProjectVariance + oCurrentTask.Duration5
            End If
                
        Else
            oCurrentTask.Text30 = "Not Calc'd: Task In Progress or Complete"
            
        End If
      End If
    Next oCurrentTask
    
    If FoundBadWeights = True Then
        MsgBox Prompt:="Some Tasks Weight Values were found to be incorrect." & _
        Chr(13) & "Check the Text30 fields for details.", Buttons:=vbCritical, _
        Title:=oOrganisationDetails.OrganisationName & " -- WorkPERT Weights Error"
        
    Else
        'MsgBox "Done---!!!", vbOKOnly, oOrganisationDetails.OrganisationName
    End If
    
    oFirstTask.Duration4 = Sqr((dTotalProjectVariance / iCtr))
   
    Set oFirstTask = Nothing
    Set oCurrentTask = Nothing
End Sub

Public Sub GetDurationAsPerSigmaValues()

    Dim oFirstTask As Task
    Dim d1Sigma As Double
    Dim sMessageText As String
    
    Set oFirstTask = ActiveProject.Tasks(1)
    d1Sigma = oFirstTask.Duration4
    
    With oFirstTask
        If .Duration4 > 0 Then
            sMessageText = "Sigma Values for this project are: " & vbCrLf & _
                "1 Sigma = " & Round((((.Duration - .Duration4) / 60) / 8), 2) & " to " & Round((((.Duration + .Duration4) / 60) / 8), 2) & vbCrLf & _
                "2 Sigma = " & Round((((.Duration - (.Duration4 * 2)) / 60) / 8), 2) & " to " & Round((((.Duration + (.Duration4 * 2)) / 60) / 8), 2) & vbCrLf & _
                "3 Sigma = " & Round((((.Duration - (.Duration4 * 3)) / 60) / 8), 2) & " to " & Round((((.Duration + (.Duration4 * 3)) / 60) / 8), 2) & vbCrLf & _
                "4 Sigma = " & Round((((.Duration - (.Duration4 * 4)) / 60) / 8), 2) & " to " & Round((((.Duration + (.Duration4 * 4)) / 60) / 8), 2) & vbCrLf & _
                "5 Sigma = " & Round((((.Duration - (.Duration4 * 5)) / 60) / 8), 2) & " to " & Round((((.Duration + (.Duration4 * 5)) / 60) / 8), 2) & vbCrLf & _
                "6 Sigma = " & Round((((.Duration - (.Duration4 * 6)) / 60) / 8), 2) & " to " & Round((((.Duration + (.Duration4 * 6)) / 60) / 8), 2)
                
            MsgBox sMessageText, vbOKOnly, oOrganisationDetails.OrganisationName

        ElseIf oFirstTask.Duration4 <= 0 Then
            MsgBox "No PERT Values calculated for this project.", vbOKOnly, oOrganisationDetails.OrganisationName
            
        End If
    End With
    
    Set oFirstTask = Nothing
End Sub

Public Sub ShowPERTFields()

    On Error GoTo err_ShowPERTFields
    
    CustomFieldRename FieldID:=pjCustomTaskDuration1, NewName:="Optimistic Duration"
    CustomFieldRename FieldID:=pjCustomTaskDuration2, NewName:="Most Likely Duration"
    CustomFieldRename FieldID:=pjCustomTaskDuration3, NewName:="Pessimistic Duration"
    CustomFieldRename FieldID:=pjCustomTaskDuration4, NewName:="Proj. Variance / Std. Dev."
    CustomFieldRename FieldID:=pjCustomTaskNumber1, NewName:="Optimistic Weight"
    CustomFieldRename FieldID:=pjCustomTaskNumber2, NewName:="Most Likely Weight"
    CustomFieldRename FieldID:=pjCustomTaskNumber3, NewName:="Pessimistic Weight"
        
    ''Showing Optimistic, Most Likely & Pessimistic fields
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Duration1", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=3
    TableApply Name:="&Entry"
    
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Duration2", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=4
    TableApply Name:="&Entry"
    
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Duration3", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=5
    TableApply Name:="&Entry"
    
    ''Showing Optimistic, Most Likely & Pessimistic Weight fields
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Number1", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=6
    TableApply Name:="&Entry"
    
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Number2", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=7
    TableApply Name:="&Entry"
    
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Number3", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=8
    TableApply Name:="&Entry"
    
    ''Showing Standard Deviation Field
    TableEditEx Name:="&Entry", TaskTable:=True, NewName:="", NewFieldName:="Duration4", Width:=14, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, _
    RowHeight:=1, ColumnPosition:=14
    TableApply Name:="&Entry"

err_ShowPERTFields:
    'Do nothing
End Sub

Public Sub HidePERTFields()
    
    On Error GoTo err_HidePERTFields

    'Hiding Optimistic, Most Likely, & Pessimistic Fields
    SelectTaskColumn Column:="Duration1"
    ColumnDelete
    
    SelectTaskColumn Column:="Duration2"
    ColumnDelete
    
    SelectTaskColumn Column:="Duration3"
    ColumnDelete
    
    ''Hiding Optimistic, Most Likely, & Pessimistic Weight Fields
    SelectTaskColumn Column:="Number1"
    ColumnDelete
    
    SelectTaskColumn Column:="Number2"
    ColumnDelete
    
    SelectTaskColumn Column:="Number3"
    ColumnDelete
    
    'Hiding Standard Deviation Field
    SelectTaskColumn Column:="Duration4"
    ColumnDelete
    
err_HidePERTFields:
    'Do nothing
End Sub


