Attribute VB_Name = "shortcuts"
'@Folder("Other")
Option Explicit

'Shortcuts module

Public Sub openMembers()

    Dim members As Workbook
    On Error GoTo cannotOpen
    Set members = globalLib.openAndGetMembers
    members.Activate
    
    Exit Sub
    
cannotOpen:
    MsgBox "Cannot open the workbook." & vbNewLine & Err.Description
    
End Sub

Public Sub openRegistersFolder()
    ThisWorkbook.FollowHyperlink ThisWorkbook.Path & "\Registers\"
End Sub

Public Sub openReportsFolder()
    ThisWorkbook.FollowHyperlink ThisWorkbook.Path & "\Weekly Reports\"
End Sub

Public Sub openClasses()
    On Error GoTo missingClasses
    globalLib.openAndGetClasses.Activate
    
    Exit Sub
    
missingClasses:
    MsgBox "Classes workbook cannot be found" & vbNewLine & Err.Description
End Sub

Public Sub openContactForm()
    On Error GoTo formError
    getClassToContact.Show
    
    Exit Sub
    
formError:
    MsgBox "Contact list form cannot be opened." & vbNewLine & Err.Description
    
End Sub

