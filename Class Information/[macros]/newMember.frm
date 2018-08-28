VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} newMember 
   Caption         =   "Fill out the details of a new member"
   ClientHeight    =   7530
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   11664
   OleObjectBlob   =   "newMember.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "newMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder Members_Management

Private Sub cancelRegistration_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
'For combo boxes, set Match entry complete and match required to false
'The code below checks if the typed value in comboboxes matched with options available
Private Sub classBox_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If classBox.MatchFound = True And Not classBox.value = "no class" Then
        Dim register_name1 As String
        Dim register_name2 As String
        register_name1 = classBox.Text & ".xlsx"
        register_name2 = classBox.Text & ".gsheet"
        
        If globalLib.workbookExists(register_name1, globalLib.getRegistersPath) Then
            'The file exists in excel format
        ElseIf globalLib.workbookExists(register_name2, globalLib.getRegistersPath) Then
            MsgBox classBox.Text & " is in google sheets format now so it cannot be modified." _
                 & vbCrLf & "Please press cancel and go to conversion center to make it an excel file." & vbCrLf & _
                   "Note: the file might be used by instructors (they use google sheets format) and will have to be modified later."
            
        Else
            MsgBox classBox.Text & " could not be found in registers folder, please make sure registers are created"
        End If
        Exit Sub
    Else
        If Not classBox.value = "no class" Then
            MsgBox "Basic Info: " & classBox.value & " is an invalid class code, please select from the list"
            cancel = True
        End If
    End If
End Sub

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

Private Sub members_new_person(blockBool, ByRef members As Worksheet)

    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    
    Dim lRow As Integer
    members.Activate
    If (members.AutoFilterMode And members.FilterMode) Or (members.FilterMode) Then
        members.ShowAllData
    End If
    
    lRow = members.Cells.Find(What:="*", _
                              After:=Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).row
    'lRow is the last row where the info about a new member is copied to
    lRow = lRow + 1
    members.Range("A" & lRow).value = nameBox.Text
    members.Range("B" & lRow).value = surnameBox.Text
    members.Range("C" & lRow).value = classBox.Text
    If membershipYes = True Then
        members.Range("D" & lRow).value = "yes"
    Else
        members.Range("D" & lRow).value = "no"
    End If
    
    If blockBool = False Then
        members.Range("E" & lRow).value = "-"
    Else
        members.Range("E" & lRow).value = Format(yearBox.value & "/" & monthBox.value & "/" & dayBox.value, "yyyy/mm/dd")
    End If
    
    members.Range("F" & lRow).value = supportName
    members.Range("G" & lRow).value = carersNo.value
    
    If wheelchairYes = True Then
        members.Range("H" & lRow).value = "y"
    Else
        members.Range("H" & lRow).value = "n"
    End If
    
    members.Range("I" & lRow).value = requirementsText.value
    
    If photoYes = True Then
        members.Range("J" & lRow).value = "yes"
    Else
        members.Range("J" & lRow).value = "no"
    End If
    
    
    If emailContact = True Then
        members.Range("K" & lRow).value = "email"
    ElseIf smsContact = True Then
        members.Range("K" & lRow).value = "text"
    Else
        members.Range("K" & lRow).value = "telephone"
    End If
    
    If Not Trim(phoneNo.value & vbNullString) = vbNullString Then
        If Not Trim(homePhoneNo.value & vbNullString) = vbNullString Then
            members.Range("L" & lRow).value = phoneNo.value & ";" & homePhoneNo.value
        Else
            members.Range("L" & lRow).value = phoneNo.value
        End If
    Else
        members.Range("L" & lRow).value = homePhoneNo.value
    End If
    
    
    members.Range("M" & lRow).value = email.value
    members.Range("N" & lRow).value = organization.value
    members.Range("P" & lRow).value = Format(DOBYear.value & "/" & DOBMonth.value & "/" & DOBDay.value, "yyyy/mm/dd")
    members.Range("Q" & lRow).value = addressBox.value
    members.Range("R" & lRow).value = postcodeBox.value
    members.Range("S" & lRow).value = designatedContact.value
    members.Range("T" & lRow).value = extraInfoText.value
    
    If friends1 = True Then
        members.Range("U" & lRow).value = 1
    ElseIf friends2 = True Then
        members.Range("U" & lRow).value = 2
    ElseIf friends3 = True Then
        members.Range("U" & lRow).value = 3
    ElseIf friends4 = True Then
        members.Range("U" & lRow).value = 4
    ElseIf friends5 = True Then
        members.Range("U" & lRow).value = 5
    End If

    If fit1 = True Then
        members.Range("V" & lRow).value = 1
    ElseIf fit2 = True Then
        members.Range("V" & lRow).value = 2
    ElseIf fit3 = True Then
        members.Range("V" & lRow).value = 3
    ElseIf fit4 = True Then
        members.Range("V" & lRow).value = 4
    ElseIf fit5 = True Then
        members.Range("V" & lRow).value = 5
    End If
    
    If confident1 = True Then
        members.Range("W" & lRow).value = 1
    ElseIf confident2 = True Then
        members.Range("W" & lRow).value = 2
    ElseIf confident3 = True Then
        members.Range("W" & lRow).value = 3
    ElseIf confident4 = True Then
        members.Range("W" & lRow).value = 4
    ElseIf confident5 = True Then
        members.Range("W" & lRow).value = 5
    End If
    
    If travelPublicTransport = True Then
        members.Range("X" & lRow).value = "Public transport"
    ElseIf travelTaxi = True Then
        members.Range("X" & lRow).value = "Taxi"
    ElseIf travelCar = True Then
        members.Range("X" & lRow).value = "Personal car"
    ElseIf travelWalkScooter = True Then
        members.Range("X" & lRow).value = "Walking/mobility scooter"
    Else
        members.Range("X" & lRow).value = "Other"
    End If
    
    If sdsYes = True Then
        members.Range("Y" & lRow).value = "yes"
    ElseIf sdsNo = True Then
        members.Range("Y" & lRow).value = "no"
    End If
    
    If cheque.value = True Then
        members.Range("Z" & lRow).value = "Cheque"
    ElseIf cash.value = True Then
        members.Range("Z" & lRow).value = "Cash"
    ElseIf directTransfer.value = True Then
        members.Range("Z" & lRow).value = "Direct transfer"
    End If
    
    If adultMemb = True Then
        members.Range("AA" & lRow).value = "Adult"
    ElseIf youthMemb = True Then
        members.Range("AA" & lRow).value = "Youth"
    ElseIf noneMemb = True Then
        members.Range("AA" & lRow).value = "None"
    End If
    
    'members.Rows(lRow).RowHeight = 15
        
    'Sort by surnames and close the workbook
    Call globalLib.sortSurnames(members, "B", 2, "AZ", lRow)
    
    
End Sub

Private Sub registers_new_person(blockBool, ByRef reg_class As Worksheet, ByRef notes_class As Worksheet)
    

    
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
    'Lowest row goes down by one for the new person, info from the form is copied
    lRow = lRow + 1
    reg_class.Range("A" & lRow).value = carersNo.Text
    reg_class.Range("B" & lRow).value = UCase(nameBox.Text)
    reg_class.Range("C" & lRow).value = UCase(surnameBox.Text)
    
    If wheelchairYes = True Then
        reg_class.Range("D" & lRow).value = "y"
    Else
        reg_class.Range("D" & lRow).value = "n"
    End If
    
    If membershipYes = True Then
        reg_class.Range("E" & lRow).value = True
    Else
        reg_class.Range("E" & lRow).value = False
    End If
    
    'Make the row height to be 40
    reg_class.Rows(lRow).RowHeight = 40
    reg_class.Rows(lRow).VerticalAlignment = xlVAlignCenter
    
    'Aligning prettily
    Call globalLib.alignLine(lRow)
    
    Dim lCol As Integer
    'Populate member's line with with false
    lCol = Cells.Find(What:="*", _
                      After:=Range("A1"), _
                      LookAt:=xlPart, _
                      LookIn:=xlFormulas, _
                      SearchOrder:=xlByColumns, _
                      SearchDirection:=xlPrevious, _
                      MatchCase:=False).Column
                    
    'Populate the line with false
    reg_class.Range("F" & lRow & ":" & globalLib.colNumToLetter(lCol) & lRow).value = False
    Dim ind As Integer
            
    For ind = globalLib.colLetterToNum("H") To lCol Step 3
        reg_class.Cells(lRow, ind).value = ""
    Next ind
    
    
    'Call colouring, membership
    Call colourCoding.past_lessons_colour
    'Alphabetical sort called last!!!
    Call globalLib.sortSurnames(reg_class, "C", 11, globalLib.colNumToLetter(lCol), lRow + 1)
    
    Dim wheelchairCount As Integer
    wheelchairCount = 0
    Dim currentRow
    
    If wheelchairYes.value = True Then
        For currentRow = 11 To lRow
            If reg_class.Range("D" & currentRow).value = "y" Then
                wheelchairCount = wheelchairCount + 1
            End If
        Next currentRow
    End If
    

    '-----------------------------------------------------------------------------

    
    notes_class.Activate
    lRow = notes_class.Cells.Find(What:="*", _
                                  After:=Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False).row
                    
    lRow = lRow + 1
    
    notes_class.Rows(lRow).RowHeight = 40
    notes_class.Rows(lRow).VerticalAlignment = xlVAlignCenter
    
    notes_class.Range("A" & lRow).value = UCase(nameBox.Text)
    notes_class.Range("B" & lRow).value = UCase(surnameBox.Text)
    
    
    'Sort by surnames
    Call globalLib.sortSurnames(notes_class, "B", 2, "Z", lRow + 1)
    
    
    '--------------------------------------------------------------------------
    
    reg_class.Activate
    
    
    Unload Me
    
    If wheelchairCount > 5 Then
        'Warn user
        Dim msg, title, response
        msg = "There are now " & wheelchairCount & " wheelchair users in " & classBox.Text & " class"
        title = "Wheelchair limit warning"
        response = MsgBox(msg, vbOKOnly, title)
    End If
    
    MsgBox "A new member was added successfully"
    
    Exit Sub
    
End Sub

Private Sub save_Click()
    
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
    
    'Basic Info: Date of birth check
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
    'Basic Info check
    If correctDate = True Then
    
        If classBox = vbNullString Then
            MsgBox "Please select the class in Basic Info Tab"
        
        ElseIf Trim(nameBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the first name in Basic Info Tab"
        
        ElseIf Trim(surnameBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the surname in Basic Info Tab"
        
        ElseIf correctDOB = False Then
            MsgBox "Please enter the date of birth in Basic Info Tab or check if the date is valid"
        
        ElseIf Trim(addressBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the address in Basic Info Tab"
    
        ElseIf Trim(postcodeBox.value & vbNullString) = vbNullString Then
            MsgBox "Please enter the postcode in Basic Info Tab"
            
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
        
            'Contact details check
        ElseIf emailContact = False And smsContact = False And callContact = False Then
            MsgBox "Please specify preferred communication in Contact details Tab"
         
        ElseIf phoneNo = vbNullString And homePhoneNo = vbNullString And callContact = True Then
            MsgBox "Contact details Tab: please type a phone number (member's preferred communication)"
        
        ElseIf phoneNo = vbNullString And smsContact = True Then
            MsgBox "Contact details Tab: please type a mobile phone number (member's preferred communication)"
            
        ElseIf email = vbNullString And emailContact = True Then
            MsgBox "Contact details Tab: please type an email (member's preferred communication)"
        Else
        
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
                    
            'Update register and members with new member info
            On Error GoTo addMembersFail
            members_new_person blockBool, members
            
            If Not classBox.value = "no class" Then
                On Error GoTo addRegistersFail
                registers_new_person blockBool, reg_class, notes_class
                register_workbook.Close Savechanges:=True
            Else
                Unload Me
            End If
            
            ' Close and save workbooks
            members_workbook.Close Savechanges:=True
            
            'Unload Me
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
    
addMembersFail:
    MsgBox "Cannot add member to Members workbook. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
addRegistersFail:
    MsgBox "Cannot add member to the register workbook. " & vbNewLine & Err.Description
    Unload Me
    Exit Sub
    
End Sub

Private Sub UserForm_Initialize()
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim index As Integer
    
    'Go to the front page
    Me.MultiPage1.value = 0
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
    
    Me.classBox.ListIndex = 0
    
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


