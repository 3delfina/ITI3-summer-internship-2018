VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} generateReportPrompt 
   Caption         =   "Generate a report"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "generateReportPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "generateReportPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Report_Generation")

Private Sub generate_button_Click()
    
    generateReport.week_name = week_box.value
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    
    'Set up application setting
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim week_num As Integer
    
    On Error GoTo templateFail
    week_num = CInt(Right(generateReport.week_name, Len(generateReport.week_name) - 5))
    
    For week = week_num To 1 Step -1
        week_box.AddItem "Week " & week
    Next week

    week_box.Text = week_box.List(0)
    
    Exit Sub
    
templateFail:
    MsgBox "Cannot retrieve week number. Please ensure the template has not been changed." & vbNewLine & Err.Description

End Sub


