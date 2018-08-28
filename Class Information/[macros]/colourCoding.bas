Attribute VB_Name = "colourCoding"
'@Folder Register_Management

Public Sub block_color(ByVal rightmostCol As String)
    'Call it once during register creation
    Dim Copyrange1 As String                     'ranges to select and work with
    Let Copyrange1 = "$A$1:$" & rightmostCol & "$150"
    
    'Color entire row if see a "block" word
    ActiveSheet.Range(Copyrange1).Select
    
    If globalLib.isExcel2010 Then
        'Excel 2010
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=COUNTIF(1:1; ""block"")"
    Else
        'Excel 365
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=COUNTIF(1:1, ""block"")"
    End If
    
           
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 39423
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Item(1).Priority = 3
    
    
End Sub

Public Sub Colour(ByVal att As String, ByVal paid As String)

    Dim Copyrange1 As String                     'ranges to select and work with
    Dim formula As String                        'formula
    Dim endRange As String
    endRange = globalLib.colNumToLetter(globalLib.colLetterToNum(paid) + 1)
    Let Copyrange1 = "$A$11:$" & endRange & "$150"
    ActiveSheet.Range(Copyrange1).Select         'Range("A11:I150").Select for example
    
    If globalLib.isExcel2010 Then
        'Excel 2010
        Let formula = "=AND(NOT($" & paid & "11); $" & att & "11 = TRUE; $" & paid & "11<>"""")"
    Else
        'Excel 365
        Let formula = "=AND(NOT($" & paid & "11), $" & att & "11 = TRUE, $" & paid & "11<>"""")"
    End If
    
    
    ' Formula is of this format "=AND(NOT($H11); $attend = TRUE; $H11<>"""")"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                   formula
     
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(248, 244, 94)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Item(1).Priority = 2
    

    
    
End Sub

Public Sub Membership()

    'Colour cells which are FALSE and not empty in the membership column
    ActiveSheet.Range("E11:E150").Select
   
    If globalLib.isExcel2010 Then
        'Excel 2010
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=AND(NOT($E11);$E11<>"""")"
    Else
        'Excel 365
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=AND(NOT($E11),$E11<>"""")"
    End If
    
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(135, 206, 250)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Item(1).Priority = 1
    
    
End Sub

'Remove format conditions (old macros)
Public Sub clear()
    With ThisWorkbook.ActiveSheet.Range("A1:ZZ150")
        .FormatConditions.Delete
    End With
End Sub

Public Sub past_lessons_colour()
    'Find lessons which are in the past and colour those who did not pay
    'Remove previous format conditions first
    Call clear
    
    'Called first to have priority one in google sheets
    'Colour those who did not pay for the membership
    Call Membership
    
    Dim lCol As Long                             'rightmost column with the date
    Dim attCol As Long                           'column for attended checkboxes
    Dim paidCol As Long                          'column for paid checkboxes
    Dim dateCell As Range                        'cell containing a date
    
    
    'Find the last lesson column in the spreadsheet
    lCol = ActiveSheet.Cells.Find(What:="*", _
                                  After:=ActiveSheet.Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False).Column
                    

    ActiveSheet.Range("F2").Select               'Select the cell with the date of the first lesson
    Set dateCell = ActiveCell                    'Remember the position of first lessons date
  
    
    'Iterate till the date of the last lesson
    Do While ActiveCell.Column <= lCol
                
        attCol = ActiveCell.Column               'set the column of attended checkboxes
        paidCol = ActiveCell.Column + 1          'set the column of paid checkboxes
            
        'Call the colouring function to check for non-payers
        Call Colour(globalLib.colNumToLetter(attCol), globalLib.colNumToLetter(paidCol))
        dateCell.Select                          'Active cell is again the date of the lesson
               
        
        ActiveSheet.Range(globalLib.colNumToLetter(ActiveCell.Column) & "11:" & globalLib.colNumToLetter(ActiveCell.Column + 2) & 150).Borders(xlEdgeRight).LineStyle = xlContinuous
        'Move to the next date of the lesson
        ActiveCell.OFFSET(0, 3).Select
        Set dateCell = ActiveCell
    Loop
    
    'Color cells in orange if there are block payments
    Call block_color(globalLib.colNumToLetter(lCol))

    'Called again to be visible in excel
    'Colour those who did not pay for the membership
    Call Membership
    
End Sub
