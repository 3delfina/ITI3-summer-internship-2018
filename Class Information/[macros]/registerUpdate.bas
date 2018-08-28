Attribute VB_Name = "registerUpdate"
'@Folder Register_Management
Option Explicit

Public Sub updateAll()
    Call updateRegisters("all")
End Sub

Public Sub updateBlock()
    Call updateRegisters("block")
End Sub

Public Sub updateMembership()
    Call updateRegisters("membership")
End Sub

Public Sub updateMemberNotes()
    Call updateRegisters("notes")
End Sub

Public Sub updateAllRegisterFormulas()
    Call updateRegisters("formulas")
End Sub

''''''''''''''''''''''''''''''''''''''''''''
'Iterate through registers and update each '
''''''''''''''''''''''''''''''''''''''''''''
Private Sub updateRegisters(goal As String)

    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Open members workbook
    Dim members_workbook As Workbook, members As Worksheet
    On Error GoTo membersWorkbookFail
    Set members_workbook = globalLib.openAndGetWorkbook("members.xlsx", globalLib.getMembersPath)
    On Error GoTo membersSheetFail
    Set members = members_workbook.Worksheets("members")
    
    Dim online_registers As Variant
    ReDim online_registers(0 To 0)
    
    'Iterate through register sheets
    Dim directory As String, current_file As String
    directory = ThisWorkbook.Path & globalLib.getRegistersPath
    current_file = Dir(directory & "*.xlsx")
    
    Do While current_file <> vbNullString
    
        Workbooks.Open (directory & current_file)
        
        online_registers(UBound(online_registers)) = current_file
        ReDim Preserve online_registers(0 To UBound(online_registers) + 1)
        
        On Error GoTo updateFail
        updateRegister members_workbook, current_file, members, goal
    
        current_file = Dir()
    
    Loop
    
    'Update statuses in Register sheet on master
    updateStatus online_registers
    
    'Close members sheet
    members_workbook.Close Savechanges:=True
    
    
    'Open register list sheet in master
    Dim master As Workbook
    Set master = Workbooks("master.xlsm")
    master.Worksheets("Control Centre").Select
    
    MsgBox "       Finished updating"

    Exit Sub
    
updateFail:
    MsgBox "Cannot update register " & current_file & vbNewLine & vbNewLine & Err.Description
    Exit Sub
    
membersWorkbookFail:
    MsgBox "Members workbook cannot be opened. " & vbNewLine & Err.Description
    Exit Sub
    
membersSheetFail:
    MsgBox "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    Exit Sub


End Sub

'''''''''''''''''''''''''''
' Update single register  '
'''''''''''''''''''''''''''
Private Sub updateRegister(ByVal members_workbook As Workbook, ByVal register_name As String, ByRef members As Worksheet, goal As String)

    Debug.Print "Starting " & register_name & "..."

    'Open register workbook
    Dim register_workbook As Workbook, register As Worksheet, register_notes As Worksheet
    On Error GoTo regWorkbookFail:
    Set register_workbook = Workbooks(register_name)
    On Error GoTo classSheetFail:
    Set register = register_workbook.Worksheets("Class")
    On Error GoTo notesSheetFail:
    Set register_notes = register_workbook.Worksheets("Notes")
    
    If goal = "block" Or goal = "membership" Or goal = "all" Or goal = "notes" Then
    
        'Set up register related values
        Dim reg_start_column As String, reg_start_row As Integer, reg_end_row As Integer
        'HARDCODED
        reg_start_column = "B"
        reg_start_row = 11
        reg_end_row = globalLib.getLastRow(register)
    
        'Set up members related values
        Dim mem_start_column As String, mem_start_row As Integer, mem_end_row As Integer
        'HARDCODED
        mem_start_column = "A"
        mem_start_row = 2
        mem_end_row = globalLib.getLastRow(members)
    
        'Iterate over members in members sheet and register sheet to find a match
        Dim name_in_reg As String, name_in_mem As String, expected_class As String, i As Integer, j As Integer

        For i = reg_start_row To reg_end_row
    
            name_in_reg = register.Range("B" & i).value & register.Range("C" & i).value
        
            For j = mem_start_row To mem_end_row
        
                name_in_mem = UCase(members.Range("A" & j).value & members.Range("B" & j).value)
                expected_class = members.Range("C" & j).value & ".xlsx"
            
                If (name_in_reg = name_in_mem) And (register_name = expected_class) Then
                    'do stuff
                    If goal = "all" Or goal = "membership" Then
                        On Error GoTo updateMembershipFail
                        updateMembershipFee i, j, members, register
                    End If
                    
                    If goal = "all" Or goal = "block" Then
                        On Error GoTo updateBlockFail
                        updateBlockPayment i, j, members, register
                    End If
                    
                    If goal = "all" Or goal = "notes" Then
                        On Error GoTo notesUpdateFail
                        updateNotes i, j, members, register_notes
                    End If
                    
                    Exit For
            
                ElseIf j = mem_end_row Then
   
                    'Error; Person in register is not in members sheet
                End If
        
            Next j
    
        Next i
    End If
    
    'Call update formulas for this register
    '------------------------------------------------
    If goal = "all" Or goal = "formulas" Then
        On Error GoTo updateFormulaFail
        Call globalLib.updateFormulasInRegisters(register_workbook)
        
        'Update the formulas in totals
        On Error GoTo updateFormulaFail
        registerCreation.addTotalsFormula register_name, register_workbook
    End If
    '------------------------------------------------

    register.Activate
    'Close register sheet
    register_workbook.Close Savechanges:=True
    
    
    Debug.Print "Finished " & register_name & "."
    
    Exit Sub

notesUpdateFail:
    Err.Raise vbObjectError + 513, "", _
              "Notes update failed. " & vbNewLine & Err.Description
    Exit Sub

updateFormulaFail:
    Err.Raise vbObjectError + 513, "", _
              "Formula update failed. " & vbNewLine & Err.Description
    Exit Sub

updateBlockFail:
    Err.Raise vbObjectError + 513, "", _
              "Block payment update failed. " & vbNewLine & Err.Description
    Exit Sub

updateMembershipFail:
    Err.Raise vbObjectError + 513, "", _
              "Membership fee update failed. " & vbNewLine & Err.Description
    Exit Sub

classSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Class sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Exit Sub
    
notesSheetFail:
    Err.Raise vbObjectError + 513, "", _
              "Notes sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Exit Sub
      
regWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
              register_name & " cannot be opened. " & vbNewLine & Err.Description
       
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sync membership fee information between a register sheet and the members sheet '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub updateMembershipFee(ByVal reg_row As Integer, _
                                ByVal mem_row As Integer, _
                                ByRef members As Worksheet, _
                                ByRef register As Worksheet)
                                
    Debug.Print "Syncing membership fee..."
                                
    'Constants based on templates:
    'HARDCODED
    Dim mem_col As String, reg_col As String
    mem_col = "D"
    reg_col = "E"
    
    'Get values of membership fee
    Dim reg_fee As Range, mem_fee As Range
    
    register.Activate
    Set reg_fee = register.Range(reg_col & reg_row)
    
    members.Activate
    Set mem_fee = members.Range(mem_col & mem_row)
    
    'Sync fee values
    If reg_fee.value <> mem_fee.value Then
        If reg_fee.value Then
            mem_fee.value = "yes"
        Else
            mem_fee.value = "no"
        End If
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sync notes information between a register sheet and the members sheet '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub updateNotes(ByVal reg_row As Integer, _
                        ByVal mem_row As Integer, _
                        ByRef members As Worksheet, _
                        ByRef register_notes As Worksheet)
                                
    Debug.Print "Syncing notes..."
                                
    'Constants based on templates:
    'HARDCODED
    Dim mem_col As String, reg_col As String
    mem_col = "O"
    reg_col = "C"
    
    'Members in notes sheet are in the same order but are moved up by an offset
    Dim OFFSET As Integer
    OFFSET = 9
    reg_row = reg_row - OFFSET
    
    'Get values of membership fee
    Dim reg_note As Range, mem_note As Range
    
    register_notes.Activate
    Set reg_note = register_notes.Range(reg_col & reg_row)
    
    members.Activate
    Set mem_note = members.Range(mem_col & mem_row)
    
    'Sync fee values
    If reg_note.value <> mem_note.value Then
        mem_note.value = reg_note.value
    End If

End Sub

Private Sub updateBlockPayment(ByVal reg_row As Integer, _
                               ByVal mem_row As Integer, _
                               ByRef members As Worksheet, _
                               ByRef register As Worksheet)
    Debug.Print "Updating block payment..."
                                
    Dim mem_block_column As String
    mem_block_column = "E"
    
    Dim dates_row As Integer, date_start_col As String
    dates_row = 2
    date_start_col = "F"
    
    Dim block_start_date As Variant
    block_start_date = members.Range(mem_block_column & mem_row).value

    'Check if block payment is present
    If IsDate(block_start_date) Then
        
        'If present, iterate over lesson dates in register
        Dim current_col As String
        current_col = date_start_col
        
        Dim current_date As Variant
        current_date = register.Range(current_col & dates_row).value
        
        'Set up boolean to control whether date is already in range so date comparison can be skipped
        Dim toSet As Boolean
        toSet = False
        
        Do While current_date <> ""
            
            If toSet Then
                register.Range(globalLib.colNumToLetter(globalLib.colLetterToNum(current_col) + 1) & reg_row).value = "BLOCK"
            Else
                If DateDiff("d", Format(current_date, "dd/mm/yyyy"), Format(block_start_date, "dd/mm/yyyy")) <= 0 Then
                    'Set "BLOCK" as value of payment for given dates
                    toSet = True
                    register.Range(globalLib.colNumToLetter(globalLib.colLetterToNum(current_col) + 1) & reg_row).value = "BLOCK"
                End If
            End If
            
            current_col = globalLib.colNumToLetter(globalLib.colLetterToNum(current_col) + 3)
            current_date = register.Range(current_col & dates_row).value
            
        Loop
    End If
End Sub

Private Sub updateStatus(online As Variant)

    Dim master As Workbook
    Dim registers_sheet As Worksheet
    Set master = Workbooks("master.xlsm")
    Set registers_sheet = master.Worksheets("Registers")
    
    Dim last_row As Integer
    last_row = globalLib.getLastRow(registers_sheet)
    
    Dim row As Integer
    Dim index As Integer
    
    registers_sheet.Range("B2", "B" & last_row).value = "Offline"
    
    For row = 2 To last_row
        For index = 0 To UBound(online)
            If registers_sheet.Range("A" & row).value & ".xlsx" = online(index) Then
                registers_sheet.Range("B" & row).value = "Online"
                Exit For
            End If
        Next index
    Next row

    master.save
    

End Sub
