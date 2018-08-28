Attribute VB_Name = "importModules"
'@Folder("Other")
Option Explicit

' Works!
'test
Public Sub deleteAndImport()

    '>>> CAREFUL <<<
    'The code below first deletes all existing files.
    'DO NOT RUN WITHOUT EXPORTING PROJECT FILES FIRST!
    'DATA, SWEAT AND BLOOD >>> CAN <<< BE LOST!
    'Enjoy.
    
    'Delete existing modules
    On Error Resume Next
    Dim element As Object
    For Each element In ActiveWorkbook.VBProject.VBComponents
        If element.Name <> "importModules" Then
            Debug.Print "Deleted: " & element.Name
            ActiveWorkbook.VBProject.VBComponents.Remove element
        End If
    Next
    
    'Import all *.bas files
    Dim directory As String
    Dim current_file As String

    directory = ThisWorkbook.Path & "\[macros]\"
    current_file = Dir(directory & "*")
    
    Do While current_file <> vbNullString
        If current_file <> "importModules.bas" Then
            Debug.Print "Added: " & current_file
            Application.VBE.ActiveVBProject.VBComponents.Import (directory & current_file)
        End If
        current_file = Dir()
    Loop
    
End Sub

