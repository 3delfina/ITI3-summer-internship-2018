VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} getTermDates 
   Caption         =   "Choose term start and end date"
   ClientHeight    =   4530
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4884
   OleObjectBlob   =   "getTermDates.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "getTermDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder Register_Management

Public startDate As String
Public endDate As String
Public running As Boolean



Private Sub CommandButton1_Click()
    
    'IMPORTANT: in all dates, year goes first, hashes don't work
    startDate = yearBox1.value & "/" & monthBox1.value & "/" & dayBox1.value
    endDate = yearBox2.value & "/" & monthBox2.value & "/" & dayBox2.value
    
    If Not IsDate(startDate) Then
        MsgBox "Start date does not exist"
    
    ElseIf Not IsDate(endDate) Then
        MsgBox "End date does not exist"
        
    ElseIf DateDiff("d", Format(endDate, "dd/mmm/yyyy"), Format(startDate, "dd/mmm/yyyy")) > 0 Then
        MsgBox "Start date is after the end date"
        
    ElseIf DateDiff("d", Format(endDate, "dd/mmm/yyyy"), Format(startDate, "dd/mmm/yyyy")) > -7 Then
        MsgBox "There is less than a week between start and end date, please check dates entered"
    Else
        'If you close the form, even public variables will lose their value
        running = False
        Hide
    End If
    
    
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    running = True
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Making sure that someone could just close the form without modifying anything
    
    startDate = ""
    endDate = ""
    
    '''This is temporarily here for convenience
    yearBox1.Text = 2018
    yearBox2.Text = 2018
    monthBox1.Text = 1
    monthBox2.Text = 1
    dayBox1.Text = 1
    dayBox2.Text = 31
    '''
    
    With yearBox1
        For i = 2018 To 2100
            .AddItem i
        Next
    End With
    
    With yearBox2
        For i = 2018 To 2100
            .AddItem i
        Next
    End With
    
    With monthBox1
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    
    With monthBox2
        For i = 1 To 12
            .AddItem i
        Next
    End With
    
    
    With dayBox1
        For i = 1 To 31
            .AddItem i
        Next
    End With
    
    With dayBox2
        For i = 1 To 31
            .AddItem i
        Next
    End With
End Sub



Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
    'Dim master As Workbook
    'Set master = Workbooks("master.xlsm")
    
    'master.Worksheets("Control Centre").Select
    'master.save
    running = True
End Sub
