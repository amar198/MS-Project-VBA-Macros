# MS-Project-VBA-Macros
Contains macros written for MS Project tasks.

# File Name: modPERTFunctionality.bas
In this file I've written macros for PERT where optimistic, most likely, and pessimistic durations can be entered in predefined fields Duration1, 2 and 3 respectively. The weights for these durations can also be set in Number1, 2 and 3 fields respectively. These weights will define how the Estimate (E) is calculated and placed in the Duration field. E.g. if you place weight as Optimistic - 1, Most likely - 1, and Pessimistic - 4; then (E) will be calculated as 
E = [(Optimistic * 1) + (Most likely * 1) + (Pessimistic * 4)] / 6
This macro will also calculate Standard Deviation in the Duration4 field, to calculate Sigma values to understand the over all project duration.

It also contains code for Adding, and Hiding below mentioned fields.
Field List and its description
  1. Duration1 - Optimistic Duration
  2. Duration2 - Most likely Duration
  3. Duration3 - Pessimistic Duration
  4. Duration4 - Standard Deviation
  5. Number1 - Optimistic Weight
  6. Number2 - Most likely Weight
  7. Number3 - Pessimistic Weight

Following procedures are written in this code.
  1. PERT() - Calculates the Estimate (E) for all the field at once.
  2. ShowPERTFields() - Shows all the above fields, and changes the caption as well.
  3. HidePERTFields() - Hides all the above fields.
  4. GetDurationAsPerSigmaValues() - Calculates the entire project duration and shows the duration as per all the Sigma level (1 to 6).
  
# File Name: Steps to add VBA macro.docx
This file contains the steps to add PERT related macros in your MS Project application.

# File Name: Steps to add Timeline macro.docx
This file contains the steps to add Timeline macro in your MS Project application.

# File name: modHolidayAndResourceCalendar.bas, ProjectManagement.accdb, and Leaves Teamplate.xlsx
The VBA module file modHolidayAndResourceCalendar.bas contains 2 procedures.
  1. AddHolidays() - Adds the organisation's holidays saved in an MS Access DB to the Base calendar. Please refer the database file "ProjectManagement.accdb".
  2. AddResourceWiseHolidaysAndSettingTheirBaseCalendar() - Adds the planned leaves of project resources in their respective calendars which are created on the Base calendar. The leave details are entered by team members in an excel file with the given template in "Leaves Template.xlsx" file. Each sheet name contains the name of the resource which is used for creating an entry of the resource in the MS-Project's resource sheet. Team members can add holiday's as single day or a range.
