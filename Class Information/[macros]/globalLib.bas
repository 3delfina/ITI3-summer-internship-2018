Attribute VB_Name = "globalLib"
'@Folder("Other")
'Deleted path comment, added this one
Option Explicit

Public Function getMembersPath() As String
    getMembersPath = "\Members\"
End Function

Public Function getRegistersPath() As String
    getRegistersPath = "\Registers\"
End Function

Public Function getClassesPath() As String
    getClassesPath = "\Classes\"
End Function

Public Function getContactPath() As String
    getContactPath = "\Members\Contact\"
End Function

Public Function getReportsPath() As String
    getReportsPath = "\Weekly Reports\"
End Function

Public Function getContactTemplatePath() As String
    getContactTemplatePath = getContactPath & "Template\contact-lists-template.xlsx"
End Function

Public Function isExcel2010() As Boolean
    isExcel2010 = True
End Function

Public Function openAndGetWorkbook(ByVal workbook_name As String, ByVal workbook_rel_path As String) As Workbook

    'Set up sheet path
    Dim full_path As String
    full_path = ThisWorkbook.Path & workbook_rel_path & workbook_name
    
    'Open workbook
    On Error GoTo cannotOpen
    Workbooks.Open (full_path)
    
    On Error GoTo cannotAssign
    Set openAndGetWorkbook = Workbooks(workbook_name)
    
    Exit Function
    
cannotOpen:
    Err.Raise vbObjectError + 513, "", _
              "Cannot open a workbook." & vbNewLine & Err.Description
    openAndGetWorkbook = Null
    Exit Function
    
cannotAssign:
    Err.Raise vbObjectError + 513, "", _
              "Cannot open a workbook." & vbNewLine & Err.Description
    openAndGetWorkbook = Null
    Exit Function
    
End Function

Public Function openAndGetMembers() As Workbook
    On Error GoTo worksheetFail
    Set openAndGetMembers = openAndGetWorkbook("members.xlsx", getMembersPath)
    Exit Function
    
worksheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Cannot open Members workbook." & vbNewLine & Err.Description
    

End Function

'@Ignore FunctionReturnValueNotUsed, FunctionReturnValueNotUsed
Public Function openAndGetClasses() As Workbook
    On Error GoTo worksheetFail
    Set openAndGetClasses = openAndGetWorkbook("venue-sheet.xlsx", getClassesPath)
    Exit Function
    
worksheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Cannot open Classes workbook." & vbNewLine & Err.Description
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''
'Return letter equivalent of number for a column '

'@Ignore FunctionReturnValueNotUsed''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function colNumToLetter(ByVal col As Long) As String
    Dim temp As Variant
    temp = Split(ActiveSheet.Cells(1, col).Address(True, False), "$")
    colNumToLetter = temp(0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''
'Return number equivalent of letter for a column '
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function colLetterToNum(ByVal letter As String) As Long
    colLetterToNum = ActiveSheet.Range(letter & 1).Column
End Function

Public Function getLastRow(ByRef sheet As Worksheet) As Integer
                                                     
    sheet.Activate
        
    On Error GoTo templateError
    getLastRow = sheet.Cells.Find(What:="*", _
                              After:=ActiveSheet.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row

    Exit Function
    
templateError:
    Err.Raise vbObjectError + 513, "", _
              "There is a problem with the layout of a worksheet." & vbNewLine & Err.Description
    getLastRow = 1000
                       
End Function

Public Sub sortSurnames(ByRef ws As Worksheet, _
                        ByVal sortCol As String, _
                        ByVal topRow As Integer, _
                        ByVal rCol As String, _
                        ByVal bRow As Integer)
    'sortCol - column with surnames
    'topRow  - top row of the selection (no merged cells in the selected area!)
    'rCol    - rightmost column in your range
    'bRow    - bottom row in your range
    
   
    ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.clear
    'Sort fields is of format:
    'ActiveWorkbook.Worksheets(ws).Sort.SortFields.Add Key:=Range("C10:C14")
    ActiveWorkbook.Worksheets(ws.Name).Sort.SortFields.Add Key:=ActiveSheet.Range(sortCol & topRow & ":" & sortCol & bRow) _
                                                                 , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ws.Name).Sort
        .SetRange ActiveSheet.Range("A" & topRow & ":" & rCol & bRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
End Sub

Public Function emptyCells(ByVal rng As String) As Boolean
    If WorksheetFunction.CountA(ActiveSheet.Range(rng)) = 0 Then
        emptyCells = True
    Else
        emptyCells = False
    End If
End Function

Public Sub alignLine(ByVal rowNo As Integer)

    ActiveSheet.Rows(rowNo & ":" & rowNo).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
End Sub

Public Sub updateFormulasInRegisters(ByVal register_wb As Workbook)

    Dim register As Worksheet
    On Error GoTo classSheetFail
    Set register = register_wb.Worksheets("Class")
    
    register_wb.Activate
    register.Select
    register.Range("F5").Select                  'Select first cell with expected fees formula
    
    Dim lCol As Integer
    Dim currentCell As Range
    
    'Find last column
    lCol = Cells.Find(What:="*", _
                      After:=Range("A1"), _
                      LookAt:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Column
        
    Set currentCell = ActiveCell
    
    'Get the class fee from term totals
    Dim fee As Integer
    On Error GoTo termSheetFail
    fee = register_wb.Worksheets("Term Totals").Range("B2").value
    
    'Iterate over columns and insert formulas
    Do While currentCell.Column <= lCol
    
        'Expected fees formula
        register.Range(currentCell.Address).FormulaR1C1 = _
                                                        "=(COUNTIF(R[6]C[1]:R[145]C[1], TRUE)-R[5]C )*" & fee
        
        'Put a formula for number of people who attended
        register.Range(currentCell.OFFSET(4, 0).Address).FormulaR1C1 = _
                                                                     "=COUNTIF(R[2]C:R[143]C, TRUE)"
        
        ActiveCell.OFFSET(0, 3).Select
        Set currentCell = ActiveCell
    Loop
    
    'Save formulas
    ActiveWorkbook.save
        
    Exit Sub
    
classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Worksheet 'Class' does not exists in a workbook " & register_wb.Name & vbNewLine & Err.Description
    Exit Sub
termSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Worksheet 'Term Totals' does not exists in a workbook " & register_wb.Name & vbNewLine & Err.Description
    Exit Sub
    

End Sub

Sub centerAcrossSelection(ByVal register As Workbook)

    Dim reg_class As Worksheet
    On Error GoTo classSheetFail
    Set reg_class = register.Worksheets("Class")
    reg_class.Activate
    reg_class.Range("F1").Select
    
    Dim lCol As Integer
    'Find last column
    lCol = Cells.Find(What:="*", _
                      After:=Range("A1"), _
                      LookAt:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Column
    
    Do While ActiveCell.Column <= lCol
        Range(colNumToLetter(ActiveCell.Column) & "2:" & colNumToLetter(ActiveCell.Column + 2) & "2").HorizontalAlignment = xlHAlignCenterAcrossSelection
        ActiveCell.OFFSET(0, 3).Select
    Loop
    
    Exit Sub
        
classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Worksheet 'Class' does not exists in a workbook " & register.Name & vbNewLine & Err.Description
    Exit Sub

End Sub

Public Function workbookExists(ByVal workbook_name As String, ByVal workbook_rel_path As String) As Boolean

    'Set up sheet path
    Dim full_path As String
    full_path = ThisWorkbook.Path & workbook_rel_path & workbook_name
    
    If Dir(full_path) = "" Then
        workbookExists = False
    Else
        workbookExists = True
    End If
    
End Function

Public Sub addHyperlink()
    
End Sub
