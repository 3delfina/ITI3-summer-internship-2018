Attribute VB_Name = "contactMembers"
'@Folder Contacting_Members

Public Sub contactMembers(ByRef found_classes() As String, _
                          ByVal chosen_day As String, _
                          ByVal chosen_venue As String)
    
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Create contact file from template
    Dim contact_workbook As Workbook
    
    On Error GoTo templateFail
    Set contact_workbook = Workbooks.Add(ThisWorkbook.Path & globalLib.getContactTemplatePath)
    
    'Set filename to date created, include searched day and/or venue if present
    Dim filename As String
    filename = vbNullString
    
    Dim current_date As String
    current_date = Format$(Now, "dd-mmm-yyyy-hh-nn-ss")
    
    If chosen_day <> "" Then
        filename = chosen_day & "-"
    End If
    
    If chosen_venue <> "" Then
        filename = chosen_venue & "-"
    End If
    
    filename = filename & current_date & ".xlsx"

    
    'Set up worksheets
    Dim mail_sheet As Worksheet
    
    On Error GoTo mailSheetFail
    Set mail_sheet = contact_workbook.Worksheets("Mail")
    
    Dim mail_last_row As Integer
    mail_last_row = 2
    
    Dim phone_sheet As Worksheet
    
    On Error GoTo phoneSheetFail
    Set phone_sheet = contact_workbook.Worksheets("Phone")
    
    Dim phone_last_row As Integer
    phone_last_row = 2
    
    Dim text_sheet As Worksheet
    
    On Error GoTo textSheetFail
    Set text_sheet = contact_workbook.Worksheets("Text")
    
    Dim text_last_row As Integer
    text_last_row = 2
    
    Dim other_sheet As Worksheet
    
    On Error GoTo otherSheetFail
    Set other_sheet = contact_workbook.Worksheets("Other")
    
    Dim other_last_row As Integer
    other_last_row = 2
    
    'Open members sheet
    Dim members_workbook As Workbook
    Dim members As Worksheet
    
    On Error GoTo membersWorkbookFail
    Set members_workbook = globalLib.openAndGetMembers
    
    On Error GoTo membersSheetFail
    Set members = members_workbook.Worksheets("members")
    
    'HARDCODE Set up template data
    Dim code_column As String
    Dim start_row As Integer
    Dim end_row As Integer
    code_column = "C"
    start_row = 2
    end_row = globalLib.getLastRow(members)
    
    'Iterate through members and check their classes
    Dim class_code As Variant
    Dim row As Integer
    Dim class_found As Boolean
    
    Dim mailList As String
    mailList = ""
    
    For row = start_row To end_row
    
        class_found = False
        
        For Each class_code In found_classes
        
            If class_code = members.Range(code_column & row).value Then
            
                class_found = True
                
                'Check preferred way of contact
                Dim pref_contact As String
                pref_contact = members.Range("K" & row).value
                
                'Add to proper sheet
                If LCase(pref_contact) = "email" Then
                    On Error GoTo populateFail
                    mailList = mailList & members.Range("M" & row).value & ";"
                    populateContactSheet mail_sheet, members, mail_last_row, row, "M"
                    
            
                ElseIf LCase(pref_contact) = "telephone" Then
                    On Error GoTo populateFail
                    populateContactSheet phone_sheet, members, phone_last_row, row, "L"
                    
                ElseIf LCase(pref_contact) = "text" Then
                    On Error GoTo populateFail
                    populateContactSheet text_sheet, members, text_last_row, row, "L"
                    
                Else
                    On Error GoTo populateFail
                    populateContactSheet other_sheet, members, other_last_row, row, "K"
                    
                End If
            
            End If
            
            If class_found Then Exit For
            
        Next class_code
    Next row
    
    
    If Not mailList = "" Then
        mailList = Left(mailList, Len(mailList) - 1)
        Debug.Print mailList
        'mail_sheet.Activate
        mail_sheet.Range("E2").value = mailList
    End If
    
    members_workbook.Close
    
    MsgBox "Contact list was created. It can be found in Members/Contact folder."
    
    'Close workbooks and open newly created
    Dim newly_created_abs_path As String
    newly_created_abs_path = ThisWorkbook.Path & globalLib.getContactPath & filename
    
    
    
    
    On Error GoTo cannotSaveNewContact
    contact_workbook.SaveAs filename:=newly_created_abs_path
    contact_workbook.Close Savechanges:=False
    
    On Error GoTo cannotOpenNewContact
    Workbooks.Open (newly_created_abs_path)
    
    Exit Sub
    
templateFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot create new contact list. Ensure that template file is in right folder. " & vbNewLine & Err.Description
    Exit Sub
    
cannotSaveNewContact:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot save newly created contact list." & vbNewLine & Err.Description
    Exit Sub
    
cannotOpenNewContact:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open newly created contact list." & vbNewLine & Err.Description
    Exit Sub
    
mailSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'mail' sheet in contact list template." & vbNewLine & Err.Description
    Exit Sub
    
phoneSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'phone' sheet in contact list template." & vbNewLine & Err.Description
    Exit Sub
    
textSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'text' sheet in contact list template." & vbNewLine & Err.Description
    Exit Sub
    
otherSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'other' sheet in contact list template." & vbNewLine & Err.Description
    Exit Sub
    
membersWorkbookFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open Members database." & vbNewLine & Err.Description
    Exit Sub
    
membersSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'members' worksheet in Members database." & vbNewLine & Err.Description
    Exit Sub
    
populateFail:
    Err.Raise vbObjectError + 513, "", _
                                "Failed to create contact list." & vbNewLine & Err.Description
    Exit Sub
    
    
End Sub

Private Sub populateContactSheet(ByRef contact As Worksheet, _
                                 ByRef members As Worksheet, _
                                 ByRef contact_row As Integer, _
                                 ByVal mem_row As Integer, _
                                 ByVal contact_column As String)
    'HARDCODED
    'members:  A Name, B Surname, C Class,
    '          L Telephone, M Email, K Prefered Communication
    '
    'contact: A Name, B Surname, C Class, D Contact
                                  
    contact.Range("A" & contact_row).value = members.Range("A" & mem_row).value
    contact.Range("B" & contact_row).value = members.Range("B" & mem_row).value
    contact.Range("C" & contact_row).value = members.Range("C" & mem_row).value
    contact.Range("D" & contact_row).value = members.Range(contact_column & mem_row).value
    
    contact_row = contact_row + 1

End Sub


