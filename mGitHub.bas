Attribute VB_Name = "mGitHub"
Function CheckIfModuleExists(fileName As String) As Long
    Dim FSO As Object
    Dim file As Object
    Dim moduleName As String
    Dim componentIndex As Long
    
    componentIndex = 1

    For Each file In ThisWorkbook.VBProject.VBComponents
        ' Remove the file extension
        moduleName = Left(fileName, Len(fileName) - 4)

        If file.Name = moduleName Then
            CheckIfModuleExists = componentIndex
            Exit Function
        End If
        
        componentIndex = componentIndex + 1
    Next file
    
    CheckIfModuleExists = 0
End Function

Sub DeleteIfModuleExists(moduleName As String)
    Dim componentIndex As Long

    componentIndex = CheckIfModuleExists(moduleName)

    If componentIndex > 0 Then
        ThisWorkbook.VBProject.VBComponents.Remove _
            ThisWorkbook.VBProject.VBComponents(componentIndex)
        'MsgBox "Module " & moduleName & " has been deleted."
    Else
        'MsgBox "Module " & moduleName & " does not exist."
    End If
End Sub

Sub ExportModulesAndClasses()

'Requirements:
' 1. The code must export all the modules and class modules to the same folder of the spreadsheet running the code.

'Acceptance Criteria:
' 1. All modules and class modules must be exported.
' 2. Modules and classes in the folder but not in the spreadsheet must not be deleted.

'---

' Declare variables
Dim FSO As Object
Dim file As Object

Set FSO = CreateObject("Scripting.FileSystemObject")

For Each file In ThisWorkbook.VBProject.VBComponents
    
    If file.Type = 1 Then '1 for standard modules
        If FSO.FileExists(ThisWorkbook.Path & "\" & file.Name & ".bas") Then
            FSO.DeleteFile (ThisWorkbook.Path & "\" & file.Name & ".bas")
        End If
        
        file.Export ThisWorkbook.Path & "\" & file.Name & ".bas"
        
    ElseIf file.Type = 2 Then '2 for classes
        If FSO.FileExists(ThisWorkbook.Path & "\" & file.Name & ".cls") Then
            FSO.DeleteFile (ThisWorkbook.Path & "\" & file.Name & ".cls")
        End If
        
        file.Export ThisWorkbook.Path & "\" & file.Name & ".cls"
        
    End If
Next file

End Sub

Sub ImportModulesAndClasses()

'Requirements:
'1.  The code must import all the modules and class modules in the same folder of the spreadsheet that is running the code.

'Acceptance Criteria:
'1.  If the module or class to be imported already exists in the spreadsheet, the new feature must consider the imported file as the updated version and replace it in the spreadsheet.
'2.  If the module or class to be imported doesn't exist in the spreadsheet, the new feature must consider the imported file as the updated version and replace it in the spreadsheet.
'3.  If a module or class in the spreadsheet doesn't exist in the folder to be imported, the version already in the spreadsheet must be kept.
'4.  The code must automatically discover the folder name of the spreadsheet running the code, so it can be used in other projects without the developer updating it directly in the code.
'5.  The code must automatically discover the folder name of the spreadsheet running the code, so it can be used in other projects without the developer updating it directly in the code.

'---

' Declare variables
Dim FSO As Object, folder As Object, file As Object
Dim fileName As String, filePath As String, moduleName As String

' Get the path of the current workbook
filePath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))

' Create a File System Object
Set FSO = CreateObject("Scripting.FileSystemObject")

' Get the folder object
Set folder = FSO.GetFolder(filePath)

' Loop through each file in the folder
For Each file In folder.Files
    ' Get the file name
    fileName = file.Name
    
    ' Check if the file is a module or class module
    If Right(fileName, 3) = "bas" Or Right(fileName, 3) = "cls" Then
    
        ' Delete the module or class inside the project if there is a file to be imported with the same name
        DeleteIfModuleExists (fileName)
   
        ' Import the new file
        ThisWorkbook.VBProject.VBComponents.Import filePath & fileName

    End If
Next file

' Clean up
Set file = Nothing
Set folder = Nothing
Set FSO = Nothing

End Sub
