Attribute VB_Name = "modHolidayAndResourceCalendar"
Dim oOrgDetails As New clsOrganizationDetails
Dim oDBConnection As New clsDbConnectivity

Public Sub AddResourceWiseHolidaysAndSettingTheirBaseCalendar()

    On Error GoTo AddResourceWiseHolidaysAndSettingTheirBaseCalendar
    
    ''Excel variables
    Dim oWorkSheet As Excel.Worksheet
    
    ''Variables required for adding dates from Excel sheet to MS project
    Dim iCtr As Integer, iSheetCounter As Integer
    Dim sStartDate As String, sEndDate As String, sEventName As String
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Opening the excel sheet to read holiday list of the current year
    Dim sExcelFilePath As String
    
    sExcelFilePath = Trim$(InputBox("Enter the path of the excel file.", oOrgDetails.OrganisationName, "C:\Amar\Personal\Studies\Project Management\Bible\Leaves Teamplate.xlsx"))
    
    If sExcelFilePath = vbNullString Then
        MsgBox "Please enter a valid file name.", vbOKOnly, oOrgDetails.OrganisationName
        GoTo exit_procedure
    End If
    
    'Opening the holiday's excel sheet to read list of holidays; sheet wise
    For iSheetCounter = 1 To Excel.Application.Workbooks.Open(sExcelFilePath).Sheets.Count
        
        Set oWorkSheet = Excel.Application.Workbooks.Open(sExcelFilePath).Sheets.Item(iSheetCounter)
        
        'Creating a new calendar based on Standard calendar to add the resource leaves in their specific calendar.
        BaseCalendarCreate Name:=Trim$(oWorkSheet.Name), FromName:="Standard"
        
        ''Setting the counter as 1, cause the values in the excel sheet are starting from 2nd row
        iCtr = 2
        
        While Trim$(oWorkSheet.Range("A" & Trim$(Str$(iCtr))).Value) <> vbNullString
            
            sEventName = Trim$(oWorkSheet.Range("A" & Trim$(Str$(iCtr))).Value)
            sStartDate = Trim$(oWorkSheet.Range("B" & Trim$(Str$(iCtr))).Value)
            
            'Assigning the end date to sEndDate variable
            sEndDate = Trim$(oWorkSheet.Range("C" & Trim$(Str$(iCtr))).Value)
            
            If sEndDate = vbNullString Then     'Checking if the end date is a blank string.
                sEndDate = sStartDate           'If yes then assigning start date to end date as well. This means that it is one day holiday and not a range.
            
            'ElseIf sEndDate <> vbNullString Then
                'This check is not required since we have already assigned the end date to this variable before the If statement started.
            
            End If
            
            ActiveProject.BaseCalendars(oWorkSheet.Name).Exceptions.Add Type:=1, Start:=sStartDate, Finish:=sEndDate, _
                Name:=sEventName
            
            iCtr = iCtr + 1
            
        Wend
        
        'After all the holiday's are added in the calendar adding a record of the resource in the resource sheet in MS project and assigning their personal calendar to their record.
        If iSheetCounter = 1 Then
            SelectResourceField Row:=0, Column:="Name"
        ElseIf iSheetCounter > 1 Then
            SelectResourceField Row:=1, Column:="Name"
        End If
        
        SetResourceField Field:="Name", Value:=Trim$(oWorkSheet.Name)
        SelectResourceField Row:=0, Column:="Base Calendar"
        SetResourceField Field:="Base Calendar", Value:=Trim$(oWorkSheet.Name)
        
    Next iSheetCounter
    
    MsgBox "Holidays Added Successfully", vbOKOnly, oOrgDetails.OrganisationName
    
exit_procedure:
    Exit Sub

AddResourceWiseHolidaysAndSettingTheirBaseCalendar:
    MsgBox "Error Description: " & Err.Description & vbCrLf & _
        "Error Number: " & Err.Number, vbOKOnly, oOrgDetails.OrganisationName
End Sub

Public Sub AddHolidays()
    On Error GoTo err_AddHolidays
    
    Dim rec As New ADODB.Recordset
    Dim sSQLQuery As String, sStartDate As String, sEndDate As String
    
    'declaring constant variable as field positions
    Const HOLIDAY As Integer = 0
    Const START_DATE As Integer = 1
    Const END_DATE As Integer = 2

    rec.ActiveConnection = oDBConnection.GetConnection
    
    sSQLQuery = "SELECT Holiday, StartDate, EndDate " & _
        "FROM Holidays"
        
    rec.Open sSQLQuery
    
    While rec.EOF <> True
        sStartDate = Trim$(rec.Fields(START_DATE).Value)
        
        If IsNull(rec.Fields(END_DATE).Value) Then      'If end date is not available in the database, then end date is assigned the start date. This would mean that there is one day holiday.
            sEndDate = sStartDate
        
        Else                                            'if end date has a value then assign that value to the sEndDate variable.
            sEndDate = Trim$(rec.Fields(END_DATE).Value)
            
        End If
        
        ActiveProject.BaseCalendars("Standard").Exceptions.Add Type:=1, Start:=sStartDate, Finish:=sEndDate, Name:=Trim$(rec.Fields(HOLIDAY).Value)
        rec.MoveNext
    Wend
    
    MsgBox "Holidays Added Successfully", vbOKOnly, oOrgDetails.OrganisationName
    
    Set rec = Nothing
    
    Exit Sub
    
err_AddHolidays:
    MsgBox "Error Description: " & Err.Description & vbCrLf & _
        "Error Number: " & Err.Number, vbOKOnly, oOrgDetails.OrganisationName
End Sub
