Attribute VB_Name = "mGitHub"
Sub ExportModulesAndClasses()

'Requirements
' 1 - The code must export all the modules and class modules to the same folder of the spreadsheet running the code.

'Acceptance Criteria
' 1 - All modules and class modules must be exported.
' 2 - Modules and classes in the folder but not in the spreadsheet must not be deleted.

Dim fso As Object
Dim file As Object

Set fso = CreateObject("Scripting.FileSystemObject")

For Each file In ThisWorkbook.VBProject.VBComponents
    
    If file.Type = 1 Then '1 for standard modules
        If fso.FileExists(ThisWorkbook.Path & "\" & file.Name & ".bas") Then
            fso.DeleteFile (ThisWorkbook.Path & "\" & file.Name & ".bas")
        End If
        
        file.Export ThisWorkbook.Path & "\" & file.Name & ".bas"
        
    ElseIf file.Type = 2 Then '2 for classes
        If fso.FileExists(ThisWorkbook.Path & "\" & file.Name & ".cls") Then
            fso.DeleteFile (ThisWorkbook.Path & "\" & file.Name & ".cls")
        End If
        
        file.Export ThisWorkbook.Path & "\" & file.Name & ".cls"
        
    End If
Next file

End Sub

