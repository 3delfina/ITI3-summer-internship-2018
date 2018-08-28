Attribute VB_Name = "generateReport"
'@Folder("Report_Generation")
Public week_name As String
Public week_number As Integer

Public Sub generateReport()

    'Report template layout constants
    Dim WEEK_NAME_CELL As String
    Dim WEEK_RANGE_CELL  As String
    Dim TOTAL_CELL  As String
    Dim TOTAL_MEMBERSHIP_CELL  As String
    Dim TOTAL_EXTRA_CELL As String
    Dim CLASS_COLUMN As String
    Dim ROW_OFFSET As Integer
    Dim COLLECTED_COLUMN As String
    Dim EXPECTED_COLUMN As String
    Dim PROJECT_CODE_COLUMN As String
    
    WEEK_NAME_CELL = "A1"
    WEEK_RANGE_CELL = "B1"
    TOTAL_CELL = "B3"
    TOTAL_MEMBERSHIP_CELL = "B5"
    TOTAL_EXTRA_CELL = "B6"
    TOTAL_FEES_CELL = "B4"
    CLASS_COLUMN = "A"
    ROW_OFFSET = 8
    COLLECTED_COLUMN = "B"
    EXPECTED_COLUMN = "C"
    PROJECT_CODE_COLUMN = "D"
    
    'Set up application settings
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    '''Open workbooks
    'Open members workbook
    Dim members_workbook As Workbook, members As Worksheet
    On Error GoTo membersWorkbookFail
    Set members_workbook = globalLib.openAndGetWorkbook("members.xlsx", globalLib.getMembersPath)
    On Error GoTo membersSheetFail
    Set members = members_workbook.Worksheets("members")
    
    'Open classes workbook
    Dim classes_workbook As Workbook, classes As Worksheet
    On Error GoTo classesWorkbookFail
    Set classes_workbook = globalLib.openAndGetClasses
    On Error GoTo classesSheetFail
    Set classes = classes_workbook.Worksheets("Classes")
    
    'Open template register workbook
    Dim report_workbook As Workbook
    Dim report As Worksheet
    On Error GoTo reportWorkbookFail
    Set report_workbook = globalLib.openAndGetWorkbook("template.xlsx", globalLib.getReportsPath & "template\")
    On Error GoTo summarySheetFail
    Set report = report_workbook.Worksheets("summary")
    

    'Set up values aggregated from all registers
    Dim total_membership As Variant
    Dim total_extras As Variant
    Dim week_range As String
    total_membership = 0
    total_extras = 0
       
    'Set up register iteration variables
    Dim directory As String, current_file As String
    directory = ThisWorkbook.Path & globalLib.getRegistersPath
    current_file = Dir(directory & "*.xlsx")
    
    Dim current_reg_workbook As Workbook
    Dim current_reg As Worksheet
    Dim week_col As Integer
    Dim num_classes_report As Integer
    num_classes_report = 0

    'Open first and extract current week and week range information
    If current_file <> vbNullString Then
        
        'Open register
        On Error GoTo regWorkbookFail
        Set current_reg_workbook = globalLib.openAndGetWorkbook(current_file, globalLib.getRegistersPath)
        On Error GoTo classSheetFail '1
        Set current_reg = current_reg_workbook.Worksheets("Class")
        
        'Determine name
        On Error GoTo cannotGetWeekName
        week_name = getWeekName(current_reg)     'current_reg)
        
        
        'Show prompt to get which week to consider in report
        generateReportPrompt.Show
        
        'Determine range
        week_range = getWeekRange(current_reg)   'current_reg)
        
        'Create path for report for given name and range
        Dim report_filename As String
        Dim report_path As String
        report_filename = week_name & " - " & week_range & ".xlsx" 'Replace(week_name & "-" & week_range, " ", "-")
        report_path = ThisWorkbook.Path & globalLib.getReportsPath
        
        'Save copy of template
        On Error GoTo fileAlreadyExists
        report_workbook.SaveCopyAs (report_path & report_filename)

        'Close template and first register
        current_reg_workbook.Close
        report_workbook.Close
    
        'Open newly created report file
        On Error GoTo reportWorkbookFail
        Set report_workbook = globalLib.openAndGetWorkbook(report_filename, globalLib.getReportsPath)
        On Error GoTo summarySheetFail
        Set report = report_workbook.Worksheets("summary")

    End If
    
    'Iterate over registers
    Do While current_file <> vbNullString
    
        num_classes_report = num_classes_report + 1
    
        'Open register
        Workbooks.Open (directory & current_file)
        On Error GoTo reportWorkbookFail
        Set current_reg_workbook = Workbooks(current_file)
        On Error GoTo classSheetFail
        Set current_reg = current_reg_workbook.Worksheets("Class")
        
        'Update formulas
        On Error GoTo formulaUpdateFail
        globalLib.updateFormulasInRegisters current_reg_workbook
        
        week_col = getWeekColFromWeekName(current_reg, week_name)
        
        'Add aggregated values
        total_membership = total_membership + current_reg.Cells(7, week_col).value
        total_extras = total_extras + current_reg.Cells(8, week_col).value
        
        'Get last row
        Dim row As Integer
        If report.Range("A" & (ROW_OFFSET + 1)).value = "" Then
            row = ROW_OFFSET + 1
        Else
            row = globalLib.getLastRow(report) + 1
        End If
        
                
        'Get class code
        Dim class_code As String
        class_code = Left(current_file, Len(current_file) - 5)
        
        'Set up hyperlink name
        report.Range(CLASS_COLUMN & row).formula = "=HYPERLINK(""" & directory & current_file & """,""" & class_code & """)"
        
        'Set collected
        Dim collect As Variant
        collect = current_reg.Cells(6, week_col)
        report.Range(COLLECTED_COLUMN & row).value = collect
    
        'Set expected
        Dim expected As Variant
        expected = current_reg.Cells(5, week_col)
        report.Range(EXPECTED_COLUMN & row).value = expected
        
        
        'Set project_code
        Dim code As Variant
        code = getProjectCode(classes, class_code)
        report.Range(PROJECT_CODE_COLUMN & row).value = code
         
        'Close register and get next
        current_reg_workbook.Close
        current_file = Dir()
        
    Loop
    
    '''After going though register, set up common and aggregated values
    'Set week_name
    report.Range(WEEK_NAME_CELL).value = week_name
        
    'Set week_range
    report.Range(WEEK_RANGE_CELL).value = week_range
        
    'Set total_extra
    report.Range(TOTAL_EXTRA_CELL).value = total_extras
    
    'Set total_membership
    report.Range(TOTAL_MEMBERSHIP_CELL).value = total_membership
    
    '''Set formulas
    'Set total
    report.Range(TOTAL_CELL).formula = "=SUM(B4:B6)"
    'Set total_fees
    report.Range(TOTAL_FEES_CELL).formula = "=SUM(B9:B99)"
        
        
    'Check if all classes has been considered (offline sync issue)
    Dim num_classes As Integer
    num_classes = globalLib.getLastRow(classes) - 1
           
    If Not (num_classes_report = num_classes) Then
        report.Cells(1, 3).value = "INCOMPLETE"
        MsgBox "The report has been generated but it is incomplete." & vbNewLine & "Ensure all files are available online and has been converted." & vbNewLine & "(One or more registers has been downloaded offline by instructors.)"
    Else
        MsgBox "The report has been generated."
    End If
        
    'Close workbooks
    report_workbook.Close Savechanges:=True
    members_workbook.Close
    classes_workbook.Close

    Exit Sub
    
cannotGetWeekName:
    MsgBox "Cannot retrieve week name. " & vbNewLine & Err.Description
    Exit Sub
    
formulaUpdateFail:
    MsgBox "Formula updating failed. " & vbNewLine & Err.Description
    Exit Sub

classesWorkbookFail:
    MsgBox "Classes workbook cannot be opened. " & vbNewLine & Err.Description
    Exit Sub
    
classesSheetFail:
    MsgBox "Worksheet 'Classes' cannot be opened in Classes workbook." & vbNewLine & Err.Description
    Exit Sub
    
membersWorkbookFail:
    MsgBox "Members workbook cannot be opened." & vbNewLine & Err.Description
    Exit Sub
    
membersSheetFail:
    MsgBox "orksheet 'members' cannot be opened in Members workbook." & vbNewLine & Err.Description
    Exit Sub
    
reportWorkbookFail:
    MsgBox "Report workbook cannot be opened. " & vbNewLine & Err.Description
    Exit Sub
    
summarySheetFail:
    MsgBox "Worksheet 'summary' cannot be opened in a report workbook." & vbNewLine & Err.Description
    Exit Sub
    
regWorkbookFail:
    MsgBox "Cannot open register " & current_file & vbNewLine & Err.Description
    Exit Sub
    
classSheetFail:
    MsgBox "Worksheet 'Class' cannot be opened in register " & current_file & vbNewLine & Err.Description
    Exit Sub
    
fileAlreadyExists:
    MsgBox "Given report already exists. If you wish to generate it again, please delete or rename the old file first and try again."
    Exit Sub

End Sub

Private Function getWeekName(register As Worksheet) As String
    
    Dim current_date As String
    Dim date_row As Integer
    Dim date_column As Integer
    
    date_row = 2
    date_column = 6
    current_date = register.Cells(date_row, date_column).value
    
    On Error GoTo dateError
    Do While (DateDiff("d", Format(Date, "dd/mmm/yyyy"), Format(current_date, "dd/mmm/yyyy")) < 0)
        
        date_column = date_column + 3
        current_date = register.Cells(date_row, date_column).value
        
        If current_date = "" Then
            getWeekName = register.Cells(1, date_column - 3).value
            Exit Function
        End If
    Loop

    getWeekName = register.Cells(1, date_column).value
    
    Exit Function
    
dateError:
    Err.Raise vbObjectError + 513, "", _
              "The date in workbook is in wrong format. Please ensure English date is used in form 'dd/mmm/yyyy' where mmm is three first letters of a month." & vbNewLine & Err.Description
    getWeekName = "Date Error"
    Exit Function
    
End Function

Private Function getWeekColFromWeekName(register As Worksheet, week_str As String) As Integer
    
    Dim col As Integer
    Dim end_col As Integer
    Dim tmp_cell As Variant
    
    register.Activate
    Set tmp_cell = register.Cells(3, 6)
    
    end_col = register.Range(tmp_cell.Address).End(xlToRight).Column

    For col = 6 To end_col Step 3
    
        If week_str = register.Cells(1, col).value Then
            getWeekColFromWeekName = col
        End If
    
    Next col
    
    If getWeekColFromWeekName = 0 Then
        MsgBox "Cannot find given week in a register."
    End If

End Function

Private Function getWeekRange(ByRef register As Worksheet) As String
    
    Dim col As Integer
    col = getWeekColFromWeekName(register, week_name)
    
    Dim class_day As Variant
    class_day = register.Cells(2, col).value
    
    Dim range_start As Date
    Dim range_end As Date
    
    range_start = class_day - (Weekday(class_day, vbMonday) - 1) 'DateAdd("ww", -1, class_day - (Weekday(class_day, vbMonday) - 1))
    range_end = DateAdd("d", 6, range_start)
    
    Dim week_range As String
    week_range = Format(range_start, "dd mmm")
    week_range = week_range & " - "
    week_range = week_range & Format(range_end, "dd mmm yyyy")
    
    getWeekRange = week_range
End Function

Private Function getProjectCode(ByRef classes As Worksheet, _
                                ByVal class_code As String) As String
                                
    Dim code As String
    code = ""
     
    Dim row As Integer
    For row = 2 To globalLib.getLastRow(classes) + 1
     
        If classes.Range("C" & row).value = class_code Then
            code = classes.Range("N" & row).value
            
        End If
    Next row
     
    If code = "" Then
        MsgBox "Could not find project code for class: " & class_code
    End If
    
    getProjectCode = code
                                
End Function


