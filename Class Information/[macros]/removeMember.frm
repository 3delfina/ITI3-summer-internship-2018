VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} removeMember 
   Caption         =   "Remove a member from the database"
   ClientHeight    =   7530
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11904
   OleObjectBlob   =   "removeMember.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "removeMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder Members_Management
Public exists As Boolean                         'if member exists
Public foundRow As Integer                       'row in members table where the information is stored

Private Sub search_Click()
    
    foundListBox.clear
    
    'Open members workbook
    Dim members_workbook As Workbook, members As Worksheet
    On Error GoTo workbookFail
    Set members_workbook = globalLib.openAndGetWorkbook("members.xlsx", globalLib.getMembersPath)
    On Error GoTo worksheetFail
    Set members = members_workbook.Worksheets("members")
    
    If (members.AutoFilterMode And members.FilterMode) Or (members.FilterMode) Then
        members.ShowAllData
    End If
    
    'HARDCODED
    Dim code_column As String
    Dim start_row As Integer
    Dim end_row As Integer
    code_column = "B"
    start_row = 2
    end_row = globalLib.getLastRow(members)
        
    'Set up loop variables
    Dim row As Integer
    Dim name1 As String
    Dim surname1 As String
    Dim class1 As String
    
    'Iterate over classes and check if the venue and day suit
    For row = start_row To end_row
        'HARDCODED
        name1 = members.Range("A" & row).value
        surname1 = members.Range("B" & row).value
        class1 = members.Range("C" & row).value
        If ((LCase(nameBox.value) = LCase(name1)) Or (nameBox.value = "")) And _
           ((LCase(surnameBox.value) = LCase(surname1)) Or (surnameBox.value = "")) And _
           ((class1 = classBox.value) Or (classBox.value = "")) Then
            foundListBox.AddItem (row & ": " & name1 & " " & surname1 & ", " & class1 & " " & members.Range("P" & row).value)
        End If
    Next row
    
    If foundListBox.ListCount = 0 Then
        MsgBox "Sorry, no matching information found, please check the information entered"
    Else
        MsgBox "Matches were found, please select a member and press Remove selected member"
    End If
    
    members_workbook.Close
    
    Exit Sub
    
workbookFail:
    MsgBox "Members workbook cannot be opened." & vbNewLine & Err.Description
    Unload Me
    Exit Sub
worksheetFail:
    MsgBox "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    members_workbook.Close Savechanges:=True
    Unload Me
    Exit Sub
 
End Sub

Private Sub cancelEditing_Click()
    Unload Me
End Sub

Private Sub members_delete_person(blockBool, ByRef members As Worksheet)

    'Go to members table and edit person's details (foundRow is global)
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim lRow As Integer
    members.Activate
    lRow = members.Cells.Find(What:="*", _
                              After:=Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
    
    members.Rows(foundRow).EntireRow.Delete
        
    'Sort by surnames and close the workbook
    Call globalLib.sortSurnames(members, "B", 2, "AZ", lRow)
    
    
    
End Sub

Private Sub registers_delete_person(blockBool, ByRef reg_class As Worksheet, ByRef notes_class As Worksheet)
    'Go to register in the class and notes tab, find person and edit his details
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim lRow As Integer                          'last row
    
    'Find the lowest entry in classes
    reg_class.Activate
    lRow = reg_class.Cells.Find(What:="*", _
                                After:=Range("A1"), _
                                LookAt:=xlPart, _
                                LookIn:=xlFormulas, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False).row
    
   
    Dim regRow As Integer
    Dim found_in_table As Boolean
    found_in_table = False
    
    'Go from the first entry in the register and look for a member
    'In registers all names are uppercase
    For regRow = 11 To lRow
        If UCase(surnameBox.Text) = reg_class.Range("C" & regRow).value And _
           UCase(nameBox.Text) = reg_class.Range("B" & regRow).value Then
            found_in_table = True
            Exit For
        End If
    Next regRow
    
    If found_in_table = False Then
        Debug.Print "Could not find the person in register Class"
    Else
        reg_class.Rows(regRow).EntireRow.Delete
           
    End If
    

    '-----------------------------------------------------------------------------------
    'Updating notes tab in register

    
    notes_class.Activate
    lRow = notes_class.Cells.Find(What:="*", _
                                  After:=Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False).row
                    
    
    found_in_table = False
    
    For regRow = 2 To lRow
        'Look for the person's name and surname, all uppercase
        If UCase(surnameBox.Text) = notes_class.Range("B" & regRow).value And _
           UCase(nameBox.Text) = notes_class.Range("A" & regRow).value Then
            found_in_table = True
            Exit For
        End If
    Next regRow
    
    If found_in_table = False Then
        Debug.Print "Could not find the person in register Notes"
    Else
        notes_class.Rows(regRow).EntireRow.Delete
    End If
    
    reg_class.Activate
    
End Sub

Private Function found_member_import_to_form() As Boolean

    found_member_import_to_form = True

    Dim LArray() As String
    LArray = Split(foundListBox.value, ":")
    
    'Get the row of the member in members spreadsheet
    foundRow = LArray(0)
    
    'As member was found, fill out the form cells and let user modify them
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Open members workbook
    Dim members_workbook As Workbook, members As Worksheet
    On Error GoTo workbookFail
    Set members_workbook = globalLib.openAndGetWorkbook("members.xlsx", globalLib.getMembersPath)
    On Error GoTo worksheetFail
    Set members = members_workbook.Worksheets("members")
    
    'First put name, surname and class to boxes to avoid confusion (they might have been empty)
    nameBox.value = members.Range("A" & foundRow).value
    surnameBox.value = members.Range("B" & foundRow).value
    classBox.value = members.Range("C" & foundRow).value
    
    'Lock the name, surname, class to ensure it does not change
    nameBox.Enabled = False
    surnameBox.Enabled = False
    classBox.Enabled = False
    
    If Not classBox.value = "no class" Then
        Dim register_name1 As String
        Dim register_name2 As String
        register_name1 = classBox.Text & ".xlsx"
        register_name2 = classBox.Text & ".gsheet"
        
        If globalLib.workbookExists(register_name1, globalLib.getRegistersPath) Then
            'The file exists in excel format
        ElseIf globalLib.workbookExists(register_name2, globalLib.getRegistersPath) Then
            found_member_import_to_form = False
            MsgBox classBox.Text & " is in google sheets format now so cannot be modified." _
                & vbCrLf & "Please press cancel and go to conversion center to make it an excel file." & vbCrLf & _
                "Note: the file might be used by instructors (they use google sheets format) and will have to be modified later."
    
        Else
            found_member_import_to_form = False
            MsgBox classBox.Text & " could not be found in registers folder, please make sure registers are created"
        End If
    End If
    
    
    'Close members' workbook
    members_workbook.Close
    
    Exit Function
    
workbookFail:
    Err.Raise vbObjectError + 513, "", _
                "Members workbook cannot be opened." & vbNewLine & Err.Description
    Unload Me
    Exit Function
worksheetFail:
    Err.Raise vbObjectError + 513, "", _
                "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    members_workbook.Close
    Unload Me
    Exit Function

End Function

Private Sub removeButton_Click()
    
    If Me.foundListBox.ListIndex = -1 Then
        If foundListBox.ListCount = 0 Then
            MsgBox "Please search for members first, nothing to delete from the database"
        Else
            MsgBox "No member was selected to be deleted from the database, please choose one"
        End If
    Else
        On Error GoTo importError
        If found_member_import_to_form = False Then
            Exit Sub
        End If
        'Confirm?
        'Calling 2 functions to save changes in members and registers
        'Warn user
        Dim msg, style, title, response
        msg = "The action will remove all the data about the selected person. " & vbCrLf & "Do you want to continue ?"
        style = vbYesNo + vbExclamation + vbDefaultButton2
        title = "Warning"
    
        response = MsgBox(msg, style, title)
    
        If response = vbNo Then
            Exit Sub
        End If
        
        If Not classBox.value = "no class" Then
            'Set up register to update
            Dim register_name As String
            register_name = classBox.Text & ".xlsx"
         
            'Workbook
            Dim register_workbook As Workbook
            On Error GoTo registerWorkbookFail
            Set register_workbook = globalLib.openAndGetWorkbook(register_name, globalLib.getRegistersPath)
        
            'Class worksheet
            Dim reg_class As Worksheet
            On Error GoTo classSheetFail
            Set reg_class = register_workbook.Worksheets("Class")
        
        
            'Notes sheet
            Dim notes_class As Worksheet
            On Error GoTo notesSheetFail
            Set notes_class = register_workbook.Worksheets("Notes")
        End If
        
        'Set up members workbook
        Dim members_workbook As Workbook, members As Worksheet
        
        On Error GoTo membersWorkbookFail
        Set members_workbook = globalLib.openAndGetWorkbook("members.xlsx", globalLib.getMembersPath)
        
        On Error GoTo membersSheetFail
        Set members = members_workbook.Worksheets("members")
   
        On Error GoTo membersDeleteFail
        members_delete_person blockBool, members
        
        If Not classBox.value = "no class" Then
            On Error GoTo registersDeleteFail
            registers_delete_person blockBool, reg_class, notes_class
            register_workbook.Close Savechanges:=True
        End If
        
    
        ' Close and save workbooks
        members_workbook.Close Savechanges:=True
        
        Unload Me
        MsgBox "Deleted successfully"
    End If
    
    Exit Sub
    
importError:
    MsgBox "Cannot load members information to removal form." & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
registerWorkbookFail:
    MsgBox register_name & " cannot be opened. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersWorkbookFail:
    MsgBox "Members workbook cannot be opened. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
                                
classSheetFail:
    MsgBox "Class sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersSheetFail:
    MsgBox "Members sheet cannot be opened" & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
notesSheetFail:
    MsgBox "Notes sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersDeleteFail:
    MsgBox "Cannot delete member from Members workbook. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
registersDeleteFail:
    MsgBox "Cannot delete member from the register workbook. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
End Sub

Private Sub UserForm_Initialize()
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Default: member found in member table is false
    exists = False
    Dim index As Integer
    
   
    'Set up classes sheet
    Dim classes_workbook As Workbook
    Dim classes As Worksheet
    On Error GoTo workbookFail
    Set classes_workbook = globalLib.openAndGetClasses
    On Error GoTo worksheetFail
    Set classes = classes_workbook.Worksheets("Classes")
    classes.Activate
    
    'HARDCODED
    Dim code_column As String
    Dim start_row As Integer
    Dim end_row As Integer
    Dim i As Integer
    code_column = "C"
    start_row = 2
    end_row = globalLib.getLastRow(classes)
        
    With classBox
        .AddItem "no class"
        For row = start_row To end_row
            .AddItem classes.Range(code_column & row).value
        Next row
    End With
    
    classes_workbook.Close
    
    Exit Sub
    
workbookFail:
    MsgBox "Classes workbook cannot be opened." & vbNewLine & Err.Description
    Unload Me
    Exit Sub
worksheetFail:
    MsgBox "Classes sheet cannot be opened in Classes workbook" & vbNewLine & Err.Description
    classes_workbook.Close
    Unload Me
    Exit Sub

End Sub


