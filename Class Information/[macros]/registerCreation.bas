Attribute VB_Name = "registerCreation"
'@Folder Register_Management
Option Explicit
Public warnMessage As String
Public warnNeeded As Boolean

'Registers creation
' * Add membership fee information
' * Create register for each class
' * Populate register for each class

Public Sub createRegisters()
    
    warnMessage = ""
    warnNeeded = False
    
    'Warn user
    Dim msg, style, title, response
    msg = "The action will overwrite current register data. " & vbCrLf & "Ensure the old registers have been put into archive." & vbCrLf & "Do you want to continue ?"
    style = vbYesNo + vbExclamation + vbDefaultButton2
    title = "Warning"
    
    response = MsgBox(msg, style, title)
    
    If response = vbNo Then
        Exit Sub
    End If
    
    'Call the form
    getTermDates.Show
    
    If getTermDates.running Then
        Exit Sub
    End If
    
    If Not getTermDates.startDate = "" Then
    
        'Set up application setting
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
 
        'Open register list sheet in master
        Dim master As Workbook
        Dim registers_sheet As Worksheet
        Set master = Workbooks("master.xlsm")
        On Error GoTo registersSheetFail
        Set registers_sheet = master.Worksheets("Registers")
        
        'Clear up previous shortcuts
        registers_sheet.Range("A2", "B" & 50).value = vbNullString
 
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
        Dim base_reg_workbook As Workbook
        On Error GoTo baseWorkbookFail
        Set base_reg_workbook = globalLib.openAndGetWorkbook("register-template.xlsx", globalLib.getRegistersPath & "template\")

        Dim code_column As String, code_start_row As Integer, code_end_row As Integer
        Dim price_column As String
        code_column = "C"
        code_start_row = 2
        price_column = "O"
        code_end_row = globalLib.getLastRow(classes)
            
        Dim i As Integer, class_code As String
        Dim price As Double
        For i = code_start_row To code_end_row
        
            class_code = classes.Range(code_column & i).value
            price = classes.Range(price_column & i).value
            On Error GoTo creationFail
            createRegister registers_sheet, class_code, members, base_reg_workbook, classes, price
    
        Next i
 
    
        'Close workbooks
        members_workbook.Close
        classes_workbook.Close
        base_reg_workbook.Close
    
        'Unload the form (it was hidden all the time)
        Unload getTermDates
    
    End If
    
    If warnNeeded = True Then
        'Warn user
               
        title = "Wheelchair limit warning"
        response = MsgBox(warnMessage, vbOKOnly, title)
        
    End If
    
    master.Worksheets("Control Centre").Select
    master.save
    

    MsgBox "Registers were created."
    
    Exit Sub
    
creationFail:
    MsgBox "Cannot create register." & vbNewLine & vbNewLine & Err.Description
    members_workbook.Close
    classes_workbook.Close
    base_reg_workbook.Close
    Exit Sub

membersWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
              "Members workbook cannot be opened. " & vbNewLine & Err.Description
    Exit Sub

classesWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
              "Classes workbook cannot be opened. " & vbNewLine & Err.Description
    Exit Sub

baseWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
              "Register template cannot be opened. " & vbNewLine & Err.Description
    Exit Sub
    
registersSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Cannot open 'Registers' sheet in master spreadsheet." & vbNewLine & Err.Description
    Exit Sub
membersSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    Exit Sub

classesSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Classes sheet cannot be opened in Classes workbook" & vbNewLine & Err.Description

End Sub

Private Sub createRegister(ByRef registers_sheet As Worksheet, ByVal class_name As String, members As Worksheet, base_workbook As Workbook, classes As Worksheet, price As Double)
    
    Dim register_name As String
    register_name = class_name & ".xlsx"
    
    On Error GoTo cannotSaveCopy
    base_workbook.SaveCopyAs ThisWorkbook.Path & globalLib.getRegistersPath & register_name
    
    
    'Open newly created register workbook
    Dim register_workbook As Workbook
    On Error GoTo regWorkbookFail
    Set register_workbook = globalLib.openAndGetWorkbook(register_name, globalLib.getRegistersPath)
   
    'Put the price in Term totals
    Dim reg_totals As Worksheet
    On Error GoTo totalsSheetFail
    Set reg_totals = register_workbook.Worksheets("Term Totals")
    reg_totals.Range("B2").value = price
    

    
   
    ' Call all necessary subroutines
    On Error GoTo populateFail
    populateClassSheet class_name, register_workbook, members, classes
        
    On Error GoTo populateFail
    populateNotesSheet class_name, register_workbook, members
    
    On Error GoTo populateFail
    addTotalsFormula class_name, register_workbook
    
    On Error GoTo hyperlinkFail
    createHyperlink registers_sheet, class_name
    
    
    Dim reg_class As Worksheet
    On Error GoTo classSheetFail
    Set reg_class = register_workbook.Worksheets("Class")
    reg_class.Activate
    
    
    'Close register workbook
    register_workbook.Close Savechanges:=True
    

 
    Exit Sub
    
populateFail:
    Err.Raise vbObjectError + 513, "", _
              "Cannot populate register " & register_name & " with data." & vbNewLine & "Please generate registers again after solving issue descibed below to ensure registers are not corrupted." & vbNewLine & vbNewLine & Err.Description
    Exit Sub
    
hyperlinkFail:
    Err.Raise vbObjectError + 513, "", _
              "Cannot create hyperlink to register " & register_name & " with data." & vbNewLine & "Please generate registers again after solving issue descibed below to ensure registers are not corrupted." & vbNewLine & vbNewLine & Err.Description
    Exit Sub
        
cannotSaveCopy:
    Err.Raise vbObjectError + 513, "", _
              "Cannot save the file." & vbNewLine & Err.Description
    Exit Sub
    
regWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
              register_name & " cannot be opened. " & vbNewLine & Err.Description
                                
totalsSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Totals sheet cannot be opened in register " & register_name & vbNewLine & Err.Description


classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Class sheet cannot be opened in register " & register_name & vbNewLine & Err.Description


End Sub

Private Sub createHyperlink(ByRef registers_sheet As Worksheet, ByVal class_name As String)
    

    
    Dim new_row As Integer
    new_row = globalLib.getLastRow(registers_sheet) + 1
    
    registers_sheet.Range("A" & new_row).value = "=HYPERLINK(""" & ThisWorkbook.Path & globalLib.getRegistersPath & class_name & ".xlsx" & """,""" & class_name & """)"
    registers_sheet.Range("B" & new_row).value = "Online"
    
    
End Sub

Private Sub FormatPainterCopy(ByVal threeRange As String)

    ActiveSheet.Columns("F:H").Select
    Selection.Copy
    ActiveSheet.Columns(threeRange).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                           SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Private Sub populateClassSheet(ByVal class_name As String, _
                               ByRef register As Workbook, _
                               ByRef members As Worksheet, _
                               ByRef classes As Worksheet)
    Dim reg_class As Worksheet
    On Error GoTo classSheetFail
    Set reg_class = register.Worksheets("Class")
    
    'Select to activate this sheet!
    reg_class.Select
    
    '-------------------------------
    'Setting wrap text for entire doc with some little exceptions
    reg_class.Cells.WrapText = True
    reg_class.Range("E5:E10").WrapText = False
    reg_class.Range("A2:E2").WrapText = False
    reg_class.Rows(1).WrapText = False
    reg_class.Rows(2).WrapText = False
    '-------------------------------
    'Make dates bigger
    With reg_class.Rows(2)
        .Font.Size = 12
    End With
    
    reg_class.Range("A2").value = "Class: " & class_name
    'Put filters for Marco to search
    'reg_class.Range("B3:C3").Select
    'Selection.AutoFilter
    
    
    Dim lRow As Integer                          'last row (entry in the members table)
    Dim memRow As Integer                        'row in the members table
    Dim regRow As Integer                        'row in the register table
    Dim clasRow As Integer                       'row in the classes table
    Dim inClasRow As Integer                     'second row counter in classes table to go back up
    Dim dayString As String                      'day of the week the class is happening
    Dim found As Boolean                         'stop the loop if a day of the class is found
    Dim lCol As Long                             'lCol is the column where we type attended, lCol+1 is for paid
    
    'PUT THE DATES ON THE REGISTER
    
    'Find the lowest entry in classes
    classes.Activate
    lRow = classes.Cells.Find(What:="*", _
                              After:=ActiveSheet.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
    
    'Find a row with a specific class name
    found = False
    clasRow = 2
    Do While clasRow <= lRow
        If ActiveSheet.Range("C" & clasRow).value = class_name Then
            'If found a class name, look for a day of the week it is happening, go up the rows
            inClasRow = clasRow
            Do While inClasRow > 0
                If Not IsEmpty(ActiveSheet.Range("A" & inClasRow).value) Then
                    dayString = ActiveSheet.Range("A" & inClasRow).value
                    found = True
                    Exit Do
                End If
                inClasRow = inClasRow - 1
            Loop
            Exit Do
        End If
        clasRow = clasRow + 1
    Loop
    
    ' Class not on the class list or day not found
    If found = False Then
        Debug.Print "could not find this class in the list: "; class_name
    End If
    
    'Go to register class
    reg_class.Activate
     
    Dim sDate As Date                            'start date of term
    Dim eDate As Date                            'end date of term
    Dim loopDate As Date
    
    'Creating a dictionary with days as keys and ints as value
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("Monday") = 1
    dict("Tuesday") = 2
    dict("Wednesday") = 3
    dict("Thursday") = 4
    dict("Friday") = 5
    dict("Saturday") = 6
    dict("Sunday") = 7
    
    'startDate = #3/7/2018#
    'endDate = #8/25/2018#
    
    'TODO: Sara please check on your pc
    sDate = DateValue(getTermDates.startDate)
    eDate = DateValue(getTermDates.endDate)
    loopDate = sDate
    
    
    'Find the date of the first lesson (as it happens on a specific day of the week)
    'Weekday(loopDate, 2) returns ints with Monday = 1, ... Sunday = 7

    Do While Not Weekday(loopDate, 2) = dict(dayString)
        loopDate = loopDate + 1
    Loop
            
    ActiveSheet.Range("F2").Select
    'Remember a cell where we are working to activate it after calling a different function
    Dim dateCell As Range
    Set dateCell = ActiveCell
    Dim reg_col As Integer
    Dim weekNo As Integer
    weekNo = 0
    
   
    'Iterate over all lesson dates till end of a term
    Do While DateDiff("d", Format(eDate, "dd/mmm/yyyy"), Format(loopDate, "dd/mmm/yyyy")) <= 0
    
        dateCell.value = Format(loopDate, "dd/mmm/yyyy") 'Add lesson date to the table
        
        'Adding week number
        Range(globalLib.colNumToLetter(dateCell.Column) & dateCell.row - 1).NumberFormat = "General"
        weekNo = weekNo + 1
        dateCell.OFFSET(-1, 0).value = "Week " & weekNo
        
        loopDate = loopDate + 7                  'Calculate next possible lesson date
        lCol = ActiveCell.Column                 'Remember current column
        Dim threeRange As String                 'Range of two columns which should be coloured same as F:H
        Let threeRange = globalLib.colNumToLetter(lCol) & ":" & globalLib.colNumToLetter(lCol + 2)
        If Not globalLib.colNumToLetter(lCol) = "F" Then
            Call FormatPainterCopy(threeRange)   'Copy colouring and formatting
            ActiveSheet.Range(globalLib.colNumToLetter(lCol) & "3").value = "ATTEND" 'Type attend and pay to two column headers
            ActiveSheet.Range(globalLib.colNumToLetter(lCol + 1) & "3").value = "PAY"
            ActiveSheet.Range(globalLib.colNumToLetter(lCol + 2) & "3").value = "COMMENT"
        End If
        
        'Go back to the active cell in this sheet and move to the next date
        dateCell.Select
        ActiveCell.OFFSET(0, 3).Select
        Set dateCell = ActiveCell
        reg_col = dateCell.Column
    Loop
    
    'Add formulas
    Call globalLib.updateFormulasInRegisters(register)
    
    'Apply colour coding
    reg_class.Activate
    On Error GoTo colourFail
    colourCoding.past_lessons_colour

    
    'Find the last entry in the members table
    members.Activate
    lRow = members.Cells.Find(What:="*", _
                              After:=ActiveSheet.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
    
    memRow = 2                                   ' row in the members table with the first entry after headers
    regRow = 11                                  'row in the attendance register to start with
    
    Dim wheelchairCount As Integer
    wheelchairCount = 0
    
    Do While memRow <= lRow
        'If names of classes match, then copy the information from members table
        If ActiveSheet.Range("C" & memRow).value = class_name Then
            'Set that row height to 40
            reg_class.Rows(regRow).RowHeight = 40
            reg_class.Rows(regRow).VerticalAlignment = xlVAlignCenter
            
            'Copy name, surname
            reg_class.Range("B" & regRow, "C" & regRow).value = members.Range("A" & memRow, "B" & memRow).value
            reg_class.Range("B" & regRow).value = UCase(reg_class.Range("B" & regRow).value)
            reg_class.Range("C" & regRow).value = UCase(reg_class.Range("C" & regRow).value)
             
            'Copy wheelchair info
            reg_class.Range("D" & regRow).value = members.Range("H" & memRow).value
            
            If members.Range("H" & memRow).value = "y" Then
                wheelchairCount = wheelchairCount + 1
            End If
            
            'Copy no. of carers info
            reg_class.Range("A" & regRow).value = members.Range("G" & memRow).value
            
            'Populate the line with false
            reg_class.Range("F" & regRow & ":" & globalLib.colNumToLetter(lCol + 2) & regRow).value = False
            Dim ind As Integer
            
            For ind = globalLib.colLetterToNum("H") To lCol + 2 Step 3
                reg_class.Cells(regRow, ind).value = " "
            Next ind
            
            
            'Membership fee - yes/no in table members, true/false in the register
            reg_class.Range("E" & regRow).value = (members.Range("D" & memRow).value = "yes")
            
            reg_class.Activate
            'Do alignment!
            Call globalLib.alignLine(regRow)
            members.Activate
            'Proceed to the next entry in the register
            regRow = regRow + 1
            
        End If
        
        'Done with one member, go to the next one
        memRow = memRow + 1
    Loop
    register.Activate
    
    If Not globalLib.emptyCells("B11:C11") Then
        Call globalLib.sortSurnames(reg_class, "C", 11, globalLib.colNumToLetter(reg_col - 1), regRow - 1)
    End If
    
    'TODO Call setBlockPayment while looping through members
    On Error GoTo classSheetFail:
    ActiveWorkbook.Worksheets("Class").PageSetup.CenterHorizontally = True
    
    reg_class.UsedRange.AutoFilter
    
    Call globalLib.centerAcrossSelection(register)
    'Make a class name visible
    Range("A2:E2").HorizontalAlignment = xlHAlignCenterAcrossSelection
    
    If wheelchairCount > 5 Then
        warnNeeded = True
        warnMessage = warnMessage & "There are " & wheelchairCount & " wheelchair users in " & class_name & " class" & vbCrLf
               
    End If

    Exit Sub
   
colourFail:
    Err.Raise vbObjectError + 513, "", _
              "Formatting failed in " & class_name & " . Ensure right version of Excel is set in globalLib." & vbNewLine & Err.Description
    Exit Sub
classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Class sheet cannot be opened in register " & class_name & vbNewLine & Err.Description
End Sub



Private Sub populateNotesSheet(ByVal class_name As String, _
                               ByRef register As Workbook, _
                               ByRef members As Worksheet)
    
    Dim reg_notes As Worksheet
    On Error GoTo notesSheetFail
    Set reg_notes = register.Worksheets("Notes")
    
    'Put filters for Marco to search
    'reg_notes.Activate
    'reg_notes.Range("A1:B1").Select
    'Selection.AutoFilter
    
    'Wrap text
    reg_notes.Cells.WrapText = True
    
    Dim lRow As Integer                          'last row (entry in the members table)
    Dim memRow As Integer                        'row in the members table
    Dim notesRow As Integer                      'row in contact details table
    
    'Find the last entry in the members table
    members.Activate
    lRow = members.Cells.Find(What:="*", _
                              After:=ActiveSheet.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
    
    memRow = 2                                   'row in the members table with the first entry
    notesRow = 2                                 'row in the notes table to start with
    
    Do While memRow <= lRow
        'If names of classes match, then copy the information from members table
        If ActiveSheet.Range("C" & memRow).value = class_name Then
        
            reg_notes.Rows(notesRow).RowHeight = 40
            reg_notes.Rows(notesRow).VerticalAlignment = xlVAlignCenter
            
            'Copy name, surname
            reg_notes.Range("A" & notesRow, "B" & notesRow).value = members.Range("A" & memRow, "B" & memRow).value
            reg_notes.Range("A" & notesRow).value = UCase(reg_notes.Range("A" & notesRow).value)
            reg_notes.Range("B" & notesRow).value = UCase(reg_notes.Range("B" & notesRow).value)
            
            'Set up notes from members
            reg_notes.Range("C" & notesRow).value = members.Range("O" & memRow).value
            
            'after filling in one row in notes sheet, move down by one row
            notesRow = notesRow + 1
            
        End If
        'move down by one member in the members table
        memRow = memRow + 1
    Loop
    
    reg_notes.Activate
    reg_notes.UsedRange.AutoFilter
    
    If Not globalLib.emptyCells("A2:B2") Then
        Call globalLib.sortSurnames(reg_notes, "B", 2, "Z", notesRow - 1)
    End If
    
    Exit Sub
    
notesSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Notes sheet cannot be opened in register " & class_name & vbNewLine & Err.Description
    
End Sub

Public Sub addTotalsFormula(ByVal class_name As String, ByRef register As Workbook)
    
    'Add formulas to Term Totals
    'Average number of people and sum of collected fees
        
    Dim reg_class As Worksheet
    On Error GoTo classSheetFail
    Set reg_class = register.Worksheets("Class")
    
    reg_class.Activate
    
    Dim lCol As Long
    Dim formulaStrAverage As String
    Dim formulaSum As String
    Dim OFFSET As Integer
    Dim offsetSum As Integer
    
    'Find the last lesson column in the spreadsheet
    lCol = ActiveSheet.Cells.Find(What:="*", _
                                  After:=ActiveSheet.Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False).Column
    
    ActiveSheet.Range("I9").Select               'Select the cell with the date of the first calculation of no. of people
    OFFSET = 4
  
    
    formulaStrAverage = "=AVERAGE(Class!R[5]C[4]"
    formulaSum = "= SUM(Class!R[5]C[4], Class!R[3]C[4]"
    
    'Iterate till the date of the last lesson
    Do While ActiveCell.Column <= lCol
        OFFSET = OFFSET + 3
        
        formulaStrAverage = formulaStrAverage & ",Class!R[5]C[" & OFFSET & "]"
        formulaSum = formulaSum & ",Class!R[5]C[" & OFFSET & "]" & ",Class!R[3]C[" & OFFSET & "]"
        ActiveCell.OFFSET(0, 3).Activate
    Loop
    formulaStrAverage = formulaStrAverage & ")"
    formulaSum = formulaSum & ")"
    
    ActiveWorkbook.Sheets("Term Totals").Select
    Range("B4").Select
    ActiveCell.FormulaR1C1 = formulaStrAverage
    Range("B3").Select
    ActiveCell.FormulaR1C1 = formulaSum
    
    Exit Sub
    
classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Class sheet cannot be opened in register " & class_name & vbNewLine & Err.Description
End Sub


