VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} modifyMember 
   Caption         =   "Change details of a current member"
   ClientHeight    =   7530
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11832
   OleObjectBlob   =   "modifyMember.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "modifyMember"
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
        MsgBox "Matches were found, please select a member and press Modify selected member"
    End If
    
    members_workbook.Close Savechanges:=True
    
    Exit Sub
    
workbookFail:
    MsgBox "Members workbook cannot be opened." & vbNewLine & Err.Description
    Unload Me
    Exit Sub
worksheetFail:
    MsgBox "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    members_workbook.Close
    Unload Me
    Exit Sub
    
End Sub

Private Sub cancelEditing_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'For combo boxes, set Match entry complete and match required to false
'The code below checks if the typed value in comboboxes matched with options available

Private Sub yearBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If yearBox.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Payments: " & yearBox.value & " is an invalid year, please select from the list"
        cancel = True
    End If
End Sub

Private Sub monthBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If monthBox.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Payments: " & monthBox.value & " is an invalid month, please select from the list"
        cancel = True
    End If
End Sub

Private Sub dayBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If dayBox.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Payments: " & dayBox.value & " is an invalid day, please select from the list"
        cancel = True
    End If
End Sub

Private Sub carersNo_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If carersNo.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Support: " & carersNo.value & " is an invalid number, please select number of carers"
        cancel = True
    End If
End Sub

Private Sub DOBYear_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If DOBYear.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Basic Info, D.O.B.: " & DOBYear.value & " is an invalid year, please select from the list"
        cancel = True
    End If
End Sub

Private Sub DOBMonth_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If DOBMonth.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Basic Info, D.O.B.: " & DOBMonth.value & " is an invalid month, please select from the list"
        cancel = True
    End If
End Sub

Private Sub DOBDay_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If DOBDay.MatchFound = True Then
        Exit Sub
    Else
        MsgBox "Basic Info, D.O.B.:  " & DOBDay.value & " is an invalid day, please select from the list"
        cancel = True
    End If
End Sub

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Private Sub members_edit_person(blockBool, ByRef members As Worksheet)

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
    
    If membershipYes = True Then
        members.Range("D" & foundRow).value = "yes"
    Else
        members.Range("D" & foundRow).value = "no"
    End If
    
    If blockBool = False Then
        members.Range("E" & foundRow).value = "-"
    Else
        members.Range("E" & foundRow).value = Format(yearBox.value & "/" & monthBox.value & "/" & dayBox.value, "yyyy/mm/dd")
    End If
    
    members.Range("F" & foundRow).value = supportName
    members.Range("G" & foundRow).value = carersNo.value
    
    If wheelchairYes = True Then
        members.Range("H" & foundRow).value = "y"
    Else
        members.Range("H" & foundRow).value = "n"
    End If
    
    members.Range("I" & foundRow).value = requirementsText.value
    
    If photoYes = True Then
        members.Range("J" & foundRow).value = "yes"
    Else
        members.Range("J" & foundRow).value = "no"
    End If
    
    
    If emailContact = True Then
        members.Range("K" & foundRow).value = "email"
    ElseIf smsContact = True Then
        members.Range("K" & foundRow).value = "text"
    Else
        members.Range("K" & foundRow).value = "telephone"
    End If
    
    If Not Trim(phoneNo.value & vbNullString) = vbNullString Then
        If Not Trim(homePhoneNo.value & vbNullString) = vbNullString Then
            members.Range("L" & foundRow).value = phoneNo.value & ";" & homePhoneNo.value
        Else
            members.Range("L" & foundRow).value = phoneNo.value
        End If
    Else
        members.Range("L" & foundRow).value = homePhoneNo.value
    End If
    
    
    members.Range("M" & foundRow).value = email.value
    members.Range("N" & foundRow).value = organization.value
    members.Range("P" & foundRow).value = Format(DOBYear.value & "/" & DOBMonth.value & "/" & DOBDay.value, "yyyy/mm/dd")
    members.Range("Q" & foundRow).value = addressBox.value
    members.Range("R" & foundRow).value = postcodeBox.value
    members.Range("S" & foundRow).value = designatedContact.value
    members.Range("T" & foundRow).value = extraInfoText.value
    
    If friends1 = True Then
        members.Range("U" & foundRow).value = 1
    ElseIf friends2 = True Then
        members.Range("U" & foundRow).value = 2
    ElseIf friends3 = True Then
        members.Range("U" & foundRow).value = 3
    ElseIf friends4 = True Then
        members.Range("U" & foundRow).value = 4
    ElseIf friends5 = True Then
        members.Range("U" & foundRow).value = 5
    End If

    If fit1 = True Then
        members.Range("V" & foundRow).value = 1
    ElseIf fit2 = True Then
        members.Range("V" & foundRow).value = 2
    ElseIf fit3 = True Then
        members.Range("V" & foundRow).value = 3
    ElseIf fit4 = True Then
        members.Range("V" & foundRow).value = 4
    ElseIf fit5 = True Then
        members.Range("V" & foundRow).value = 5
    End If
    
    If confident1 = True Then
        members.Range("W" & foundRow).value = 1
    ElseIf confident2 = True Then
        members.Range("W" & foundRow).value = 2
    ElseIf confident3 = True Then
        members.Range("W" & foundRow).value = 3
    ElseIf confident4 = True Then
        members.Range("W" & foundRow).value = 4
    ElseIf confident5 = True Then
        members.Range("W" & foundRow).value = 5
    End If
    
    If travelPublicTransport = True Then
        members.Range("X" & foundRow).value = "Public transport"
    ElseIf travelTaxi = True Then
        members.Range("X" & foundRow).value = "Taxi"
    ElseIf travelCar = True Then
        members.Range("X" & foundRow).value = "Personal car"
    ElseIf travelWalkScooter = True Then
        members.Range("X" & foundRow).value = "Walking/mobility scooter"
    Else
        members.Range("X" & foundRow).value = "Other"
    End If
    
    If sdsYes = True Then
        members.Range("Y" & foundRow).value = "yes"
    ElseIf sdsNo = True Then
        members.Range("Y" & foundRow).value = "no"
    End If
    
    If cheque.value = True Then
        members.Range("Z" & foundRow).value = "Cheque"
    ElseIf cash.value = True Then
        members.Range("Z" & foundRow).value = "Cash"
    ElseIf directTransfer.value = True Then
        members.Range("Z" & foundRow).value = "Direct transfer"
    End If
    
    If adultMemb = True Then
        members.Range("AA" & foundRow).value = "Adult"
    ElseIf youthMemb = True Then
        members.Range("AA" & foundRow).value = "Youth"
    ElseIf noneMemb = True Then
        members.Range("AA" & foundRow).value = "None"
    End If
    
        
    'Sort by surnames and close the workbook
    Call globalLib.sortSurnames(members, "B", 2, "AZ", lRow)

    Exit Sub
    
End Sub

Private Sub registers_edit_person(blockBool, ByRef reg_class As Worksheet, ByRef notes_class As Worksheet)
    'Go to register in the class tab, find person and edit his details
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
    
        reg_class.Range("A" & regRow).value = carersNo.Text
        'reg_class.Range("B" & regRow).value = UCase(nameBox.Text)
        'reg_class.Range("C" & regRow).value = UCase(surnameBox.Text)
    
        If wheelchairYes = True Then
            reg_class.Range("D" & regRow).value = "y"
        Else
            reg_class.Range("D" & regRow).value = "n"
        End If
    
        If membershipYes = True Then
            reg_class.Range("E" & regRow).value = True
        Else
            reg_class.Range("E" & regRow).value = False
        End If
    
        'Make the row height to be 40
        reg_class.Rows(regRow).RowHeight = 40
        reg_class.Rows(regRow).VerticalAlignment = xlVAlignCenter
    
        Dim lCol As Integer
        
        lCol = Cells.Find(What:="*", _
                          After:=Range("A1"), _
                          LookAt:=xlPart, _
                          LookIn:=xlFormulas, _
                          SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious, _
                          MatchCase:=False).Column
    
          
    End If
    
    '-----------------------------------------------------------------------------------
    
    reg_class.Activate

End Sub

Private Sub found_member_import_to_form()

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
            members_workbook.Close
            MsgBox classBox.Text & " is in google sheets format now so it cannot be modified." _
                & vbCrLf & "Please press cancel and go to conversion center to make it an excel file." & vbCrLf & _
                "Note: the file might be used by instructors (they use google sheets format) and will have to be modified later."
            Exit Sub
        Else
            members_workbook.Close
            MsgBox classBox.Text & " could not be found in registers folder, please make sure registers are created"
            Exit Sub
        End If
    End If
        
    MsgBox "Please check other tabs for information to update"
    
    
    If members.Range("D" & foundRow).value = "yes" Then
        membershipYes.value = True
    Else
        membershipNo.value = True
    End If

    
    supportName.value = members.Range("F" & foundRow).value
    carersNo.value = members.Range("G" & foundRow).value
    
    
    If members.Range("H" & foundRow).value = "y" Then
        wheelchairYes.value = True
    Else
        wheelchairNo.value = True
    End If

    requirementsText.value = members.Range("I" & foundRow).value
    
    If Not IsDate(members.Range("E" & foundRow).value) Then
        yearBox.Text = "-"
        monthBox.Text = "-"
        dayBox.Text = "-"
    Else
        yearBox.Text = year(members.Range("E" & foundRow).value)
        monthBox.Text = Month(members.Range("E" & foundRow).value)
        dayBox.Text = day(members.Range("E" & foundRow).value)
    End If
    
   
    If members.Range("J" & foundRow).value = "yes" Then
        photoYes.value = True
    Else
        photoNo.value = True
    End If
    
    
    If members.Range("K" & foundRow).value = "email" Then
        emailContact.value = True
    ElseIf members.Range("K" & foundRow).value = "text" Then
        smsContact.value = True
    ElseIf members.Range("K" & foundRow).value = "telephone" Then
        callContact.value = True
    End If
    
    
    If members.Range("L" & foundRow).value Like "*;*" Then
        LArray = Split(members.Range("L" & foundRow).value, ";")
        phoneNo.value = LArray(0)
        homePhoneNo.value = LArray(1)
    Else
        phoneNo.value = members.Range("L" & foundRow).value
    End If
    
    email.value = members.Range("M" & foundRow).value
    organization.value = members.Range("N" & foundRow).value
    
    If Not IsDate(members.Range("P" & foundRow).value) Then
        DOBYear.Text = "-"
        DOBMonth.Text = "-"
        DOBDay.Text = "-"
    Else
        DOBYear.Text = year(members.Range("P" & foundRow).value)
        DOBMonth.Text = Month(members.Range("P" & foundRow).value)
        DOBDay.Text = day(members.Range("P" & foundRow).value)
    End If
    
    addressBox.Text = members.Range("Q" & foundRow).value
    postcodeBox.Text = members.Range("R" & foundRow).value
    designatedContact.Text = members.Range("S" & foundRow).value
    extraInfoText.Text = members.Range("T" & foundRow).value
    
    If members.Range("U" & foundRow).value = 1 Then
        friends1.value = True
    ElseIf members.Range("U" & foundRow).value = 2 Then
        friends2.value = True
    ElseIf members.Range("U" & foundRow).value = 3 Then
        friends3.value = True
    ElseIf members.Range("U" & foundRow).value = 4 Then
        friends4.value = True
    ElseIf members.Range("U" & foundRow).value = 5 Then
        friends5.value = True
    End If

    If members.Range("V" & foundRow).value = 1 Then
        fit1.value = True
    ElseIf members.Range("V" & foundRow).value = 2 Then
        fit2.value = True
    ElseIf members.Range("V" & foundRow).value = 3 Then
        fit3.value = True
    ElseIf members.Range("V" & foundRow).value = 4 Then
        fit4.value = True
    ElseIf members.Range("V" & foundRow).value = 5 Then
        fit5.value = True
    End If
    
    If members.Range("W" & foundRow).value = 1 Then
        confident1.value = True
    ElseIf members.Range("W" & foundRow).value = 2 Then
        confident2.value = True
    ElseIf members.Range("W" & foundRow).value = 3 Then
        confident3.value = True
    ElseIf members.Range("W" & foundRow).value = 4 Then
        confident4.value = True
    ElseIf members.Range("W" & foundRow).value = 5 Then
        confident5.value = True
    End If
    
        
    If members.Range("Y" & foundRow).value = "yes" Then
        sdsYes.value = True
    ElseIf members.Range("Y" & foundRow).value = "no" Then
        sdsNo.value = True
    End If
    
    If members.Range("Z" & foundRow).value = "Cheque" Then
        cheque.value = True
    ElseIf members.Range("Z" & foundRow).value = "Cash" Then
        cash.value = True
    ElseIf members.Range("Z" & foundRow).value = "Direct transfer" Then
        directTransfer.value = True
    End If
    
    If members.Range("AA" & foundRow).value = "Adult" Then
        adultMemb.value = True
    ElseIf members.Range("AA" & foundRow).value = "Youth" Then
        youthMemb.value = True
    ElseIf members.Range("AA" & foundRow).value = "None" Then
        noneMemb.value = True
    End If
    
    'Close members' workbook
    members_workbook.Close
    
    'Enable all other tabs, previously only name and class were enabled
    MultiPage1.Pages(1).Enabled = True
    MultiPage1.Pages(2).Enabled = True
    MultiPage1.Pages(3).Enabled = True
    MultiPage1.Pages(4).Enabled = True
    MultiPage1.Pages(5).Enabled = True
    
    Exit Sub
    
workbookFail:
    MsgBox "Members workbook cannot be opened." & vbNewLine & Err.Description
    Unload Me
    Exit Sub
worksheetFail:
    MsgBox "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    members_workbook.Close
    Unload Me
    Exit Sub
    
End Sub

Private Sub chooseModifyButton_Click()
    
    If Me.foundListBox.ListIndex = -1 Then
        If foundListBox.ListCount = 0 Then
            MsgBox "Please search for members first, nothing to modify"
        Else
            MsgBox "No member was selected to modify, please choose one"
        End If
    Else
        Call found_member_import_to_form
    End If
End Sub

Private Sub save_Click()
    'If the member was changed and a save button was clicked,
    'first validity of data is checked,
    'then 2 functions to save changes in members and registers are called
    
    Dim blockBool As Boolean                     'shows if the person paid/not paid the block payment
    Dim blockDateString As String
    blockDateString = ""
    Dim correctDate As Boolean
    correctDate = False
    Dim correctDOB As Boolean
    correctDOB = False
    Dim DOBString As String
    DOBString = ""
    ''''''''''''''''''''''''''''''''''''''''''''
    'Check all the compulsory data is there
    ''''''''''''''''''''''''''''''''''''''''''''
    
    'Contact details: Date of birth check
    DOBString = DOBYear.value & "/" & DOBMonth.value & "/" & DOBDay.value
    If IsDate(DOBString) Then
        correctDOB = True
    End If
    
    
    'Payments: block payment check
    If yearBox = "-" And monthBox = "-" And dayBox = "-" Then
        blockBool = False
        correctDate = True
    ElseIf yearBox = "-" Or monthBox = "-" Or dayBox = "-" Then
        MsgBox "Payments tab: block payment date is wrong"
    Else
        blockBool = True
        blockDateString = yearBox.value & "/" & monthBox.value & "/" & dayBox.value
        If Not IsDate(blockDateString) Then
            MsgBox "Payments tab: block payment date does not exist"
        Else
            correctDate = True
        End If
    End If
    
    
    'Go through compulsory data
    
    If correctDate = True Then
        
        'Contact details check
        If correctDOB = False Then
            MsgBox "Please enter the date of birth in Contact details Tab or check if the date is valid"
        
        ElseIf emailContact = False And smsContact = False And callContact = False Then
            MsgBox "Please specify preferred communication in Contact details Tab"
         
        ElseIf phoneNo = vbNullString And homePhoneNo = vbNullString And (smsContact = True Or callContact = True) Then
            MsgBox "Contact details Tab: please type a phone number (member's preferred communication)"
    
        ElseIf email = vbNullString And emailContact = True Then
            MsgBox "Contact details Tab: please type an email (member's preferred communication)"
            
        ElseIf Trim(addressBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the address in Contact details Tab"
    
        ElseIf Trim(postcodeBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the postcode in Contact details Tab"
            
            'Payments check
        ElseIf membershipYes = False And membershipNo = False Then
            MsgBox "Please specify if membership was paid in Payments Tab"
                    
        ElseIf adultMemb = False And youthMemb = False And noneMemb = False Then
            MsgBox "Payments Tab: please choose type of memberhsip"
    
            'Support check
        ElseIf carersNo = vbNullString Then
            MsgBox "Please enter the number of carers in Support Tab"
                    
            'Requirements check
        
        ElseIf wheelchairYes = False And wheelchairNo = False Then
            MsgBox "Please specify the wheelchair info in Requirements Tab"
        
        
        Else
            'Calling 2 functions to save changes in members and registers
            
            If Not classBox.value = "no class" Then
                'Set up register to update
                Dim register_name As String
                register_name = classBox.Text & ".xlsx"
             
                'Workbook
                Dim register_workbook As Workbook
                On Error GoTo registerWorkbookFail   'register_name
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
            
            On Error GoTo membersEditFail
            members_edit_person blockBool, members
            
            If Not classBox.value = "no class" Then
                On Error GoTo registersEditFail
                registers_edit_person blockBool, reg_class, notes_class
                register_workbook.Close Savechanges:=True
            End If
            
            ' Close and save workbooks
            members_workbook.Close Savechanges:=True
            
            
            Unload Me
            MsgBox "Modified successfully"
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    'End of check
    ''''''''''''''''''''''''''''''''''''''''''''
    Exit Sub
    
registerWorkbookFail:
    MsgBox register_name & " cannot be opened. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
                                
classSheetFail:
    MsgBox "Class sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Unload Me
    Exit Sub
        
notesSheetFail:
    MsgBox "Notes sheet cannot be opened in register " & register_name & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersWorkbookFail:
    MsgBox "Members workbook cannot be opened. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersSheetFail:
    MsgBox "Members sheet cannot be opened in Members workbook" & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
membersEditFail:
    MsgBox "Cannot modify member in Members workbook. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
registersEditFail:
    MsgBox "Cannot modify member in the register workbook. " & vbNewLine & Err.Description
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
        
    'Go to the front page
    Me.MultiPage1.value = 0
    
    'Only enable the first page where user is looking for a member, no other tabs are available
    MultiPage1.Pages(1).Enabled = False
    MultiPage1.Pages(2).Enabled = False
    MultiPage1.Pages(3).Enabled = False
    MultiPage1.Pages(4).Enabled = False
    MultiPage1.Pages(5).Enabled = False
    
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
    
    With carersNo
        For i = 0 To 5
            .AddItem i
        Next i
    End With
    
    yearBox.Text = "-"
    monthBox.Text = "-"
    dayBox.Text = "-"
    
    With yearBox
        .AddItem "-"
        For i = 2018 To 2100
            .AddItem i
        Next
    End With
    
    With monthBox
        .AddItem "-"
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    With dayBox
        .AddItem "-"
        For i = 1 To 31
            .AddItem i
        Next
    End With
    
    With DOBYear
        For i = 1920 To 2070
            .AddItem i
        Next
    End With
    
    With DOBMonth
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    With DOBDay
        For i = 1 To 31
            .AddItem i
        Next
    End With
    
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


