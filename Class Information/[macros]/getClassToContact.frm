VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getClassToContact 
   Caption         =   "Choose classes to contact"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8256
   OleObjectBlob   =   "getClassToContact.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getClassToContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder Contacting_Members

Private Sub cancel_Click()
    Unload Me
End Sub

'@Folder("VBAProject")
Private Sub UserForm_Initialize()
    
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Set list boxes to multi selection
    foundListBox.MultiSelect = fmMultiSelectExtended
    chosenListBox.MultiSelect = fmMultiSelectExtended
    
    'Set up found classes to all classes
    Dim classes() As String
    On Error GoTo cannotGetClasses
    classes = getClasses("", "")
    updateFoundClasses classes
    
    'Set up day selection
    With dayBox
        .AddItem ""
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
    End With
    
    'Set up venue selection
    Dim index As Integer
    Dim venues() As String
    venues = getVenues()
    
    With venueBox
        .AddItem ""
        For index = 1 To UBound(venues) - LBound(venues) + 1
            .AddItem venues(index)
        Next index
    End With
    
    Exit Sub
    
cannotGetClasses:
    MsgBox "Cannot create list of classes." & vbNewLine & Err.Description
    Unload Me


End Sub

'Update found classes based on criteria
Private Sub findClassesButton_Click()
    
    Dim classes() As String
    
    On Error GoTo cannotGetClasses
    classes = getClasses(dayBox.value, venueBox.value)
    updateFoundClasses classes
   
    Exit Sub
   
cannotGetClasses:
    MsgBox "Cannot create list of classes." & vbNewLine & Err.Description

End Sub

'TODO Restrict to unique only
'Add all selected to chosen box
Private Sub addButton_Click()

    Dim index As Long
    
    For index = 0 To foundListBox.ListCount - 1
        If foundListBox.Selected(index) Then
            chosenListBox.AddItem foundListBox.List(index)
        End If
        
    Next index
End Sub

'Remove all selected from chosen box
Private Sub removeButton_Click()

    Dim index As Long
    
    For index = 0 To chosenListBox.ListCount - 1
        If index >= chosenListBox.ListCount Then
            Exit Sub
        End If
    
        If chosenListBox.Selected(index) Then
            chosenListBox.RemoveItem index
            index = index - 1
        End If
    Next index
    
End Sub

' Get chosen classes
Private Sub confirmButton_Click()

    Dim index As Long
    Dim bound As Integer
    bound = chosenListBox.ListCount
    
    Dim classes() As String
    ReDim classes(0 To bound)
    
    For index = 0 To bound - 1
        classes(index) = chosenListBox.List(index)
    Next index

    ' Call contactMembers to handle file creation
    If Len(Join(classes)) > 0 Then
        On Error GoTo contactMembersFail
        contactMembers.contactMembers classes, dayBox.value, venueBox.value
        'Close form
        Unload Me
    Else
        MsgBox "Please select a class first"
    End If
    
   Exit Sub
    
contactMembersFail:
   MsgBox "Cannot create contact list. " & vbNewLine & Err.Description
    
End Sub

'Update found classes box list
Private Sub updateFoundClasses(ByRef classes() As String)

    foundListBox.clear
    
    Dim index As Integer
    
    With foundListBox
        For index = 0 To UBound(classes) - LBound(classes)
            .AddItem classes(index)
        Next index
    End With
End Sub

'TODO Simplify classes workbook open and close
'Get all unique venues from classes sheet
Private Function getVenues() As String()

    'Set up classes sheet
    Dim classes_workbook As Workbook
    On Error GoTo classesWorkbookFail
    Set classes_workbook = globalLib.openAndGetClasses
    Dim classes As Worksheet
    On Error GoTo classesSheetFail
    Set classes = classes_workbook.Worksheets("Classes")
    
    Dim venues() As String
    
    'Use collection to get only unique values
    Dim tmp As Collection
    Set tmp = New Collection
    
    'HARDCODED
    Dim code_column As String
    Dim start_row As Integer
    Dim end_row As Integer
    code_column = "G"
    start_row = 2
    end_row = globalLib.getLastRow(classes)
    
    Dim row As Integer
    Dim value As String
    
    On Error Resume Next
    For row = start_row To end_row
        value = classes.Range(code_column & row).value
        tmp.Add value, CStr(value)
    Next row
    
    ReDim venues(1 To tmp.Count)
    
    Dim index As Integer
    For index = 1 To tmp.Count
        venues(index) = tmp.Item(index)
    Next index
    
    classes_workbook.Close
    
    getVenues = venues
    
    Exit Function
    
classesWorkbookFail:
        Err.Raise vbObjectError + 513, "", _
                                "Cannot open Classes workbook. " & vbNewLine & Err.Description
        getVenues = Null
Exit Function

classesSheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open 'Classes' worksheet in Classes workbook. " & vbNewLine & Err.Description
    getVenues = Null
Exit Function

End Function

'TODO Refactor loop to exit when found
'Return array of classes based on given day and venue criteria
Public Function getClasses(ByVal chosen_day As String, _
                           ByVal chosen_venue As String) As String()
    
    'Set up array
    Dim found() As String
    Dim found_len As Variant
    ReDim found(0 To 0)
    found_len = 0

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
    code_column = "C"
    start_row = 2
    end_row = globalLib.getLastRow(classes)

    
    'Set up loop variables
    Dim row As Integer
    Dim day As String
    Dim venue As String
    
    'Iterate over classes and check if the venue and day suit
    For row = start_row To end_row
        'HARDCODED
        day = classes.Range("A" & row).value
        venue = classes.Range("G" & row).value
        If ((day = chosen_day) Or (chosen_day = "")) And ((venue = chosen_venue) Or (chosen_venue = "")) Then
            ReDim Preserve found(0 To found_len)
            found(found_len) = classes.Range(code_column & row).value
            found_len = found_len + 1
            
        End If
    Next row

    classes_workbook.Close
    
    getClasses = found
    
    Exit Function
    
worksheetFail:
    Err.Raise vbObjectError + 513, "", _
                                "'classes' sheet does not exists in a Classes workbook." & vbNewLine & Err.Description
    getClasses = Null
    Exit Function
    
workbookFail:
    Err.Raise vbObjectError + 513, "", _
                                "Cannot open Classes workbook." & vbNewLine & Err.Description
    getClasses = Null
    Exit Function
    
End Function


