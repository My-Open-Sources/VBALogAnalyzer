Attribute VB_Name = "mRun"
Sub Run()

'Declare variables
Dim FSO As Object
Dim folder As Object
Dim file As Object
Dim ws As Worksheet
Dim lastModifiedDate As Date
Dim counter As Integer
Dim fileLimit As Integer

Application.DisplayAlerts = False

'Call the mConfig module to read the Config worksheet and set the debugMode variable
If mConfig.GetConfig = False Then Exit Sub

'Initialize counter
counter = 0

'Delete "Results" worksheet if it already exists
On Error Resume Next
ThisWorkbook.Worksheets("Results").Delete
On Error GoTo 0

'Create new "Results" worksheet
Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
ws.Name = "Results"

'Add headers
ws.Cells(1, 1) = "File Name"
ws.Cells(1, 2) = "User"
ws.Cells(1, 3) = "DB"

'Create FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")

'Get the folder where the workbook is located
Set folder = FSO.GetFolder(ThisWorkbook.Path)

'Loop through the files in the folder
For Each file In folder.Files

    'If debugMode is true, print the file name
    If debugMode Then Debug.Print " "
    If debugMode Then Debug.Print "File Name: " & file.Name

    'If debugMode is true, just consider some initial files for testing
    If debugMode Then
        fileLimit = 1000
    Else
        fileLimit = 10000
    End If
    
    'Verify log files
    If counter < fileLimit Then
    
        'If debugMode is true
        If debugMode Then Debug.Print "  - LCase(Right(file.Name, 4)): " & LCase(Right(file.Name, 4))
        If debugMode Then Debug.Print "  - Left(file.Name, 14): " & Left(file.Name, 14)
            
        'Check if the file has a ".txt" extension and the file name starts with "cm_docsupdate"
        If LCase(Right(file.Name, 4)) = ".txt" And Left(file.Name, 14) = "cm_docsupdate_" Then
            'Check if the last modified date of the file is after 30/November/2022
            lastModifiedDate = file.DateLastModified
            
            If lastModifiedDate > DateSerial(2022, 11, 30) Then
            
                'Open the file and read the content
                Dim fileContent As String
                fileContent = FSO.OpenTextFile(file.Path).ReadAll
                
                'check if the content contains the string "USER        : "
                If InStr(fileContent, "USER        : ") > 0 Then
                
                    'Check if the content of the line after "USER        : " ends with "TBatista" or "PMurphy"
                    Dim lines() As String
                    lines = Split(fileContent, vbCrLf)
                    Dim i As Long
                    Dim user As String
                    
                    For i = 0 To UBound(lines)
                        If InStr(lines(i), "USER        : ") > 0 Then
                            user = Replace(lines(i), "USER        : ", "")
                            Exit For
                        End If
                    Next i
                    
                    For i = 0 To UBound(lines)
                        If InStr(lines(i), ";DATABASE=") > 0 Then
                            Dim db As String
                            db = Replace(lines(i), ";DATABASE=", "")
                            Exit For
                        End If
                    Next i
                    
                    'Ignore if the user name is A
                    If user <> "TBatista" And user <> "PMurphy" Then
                        'Insert the file name and user in the "Results" worksheet
                        ws.Cells(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1, 1) = file.Name
                        ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1, 2) = user
                        ws.Cells(ws.Cells(ws.Rows.Count, "C").End(xlUp).Row + 1, 3) = db
                    End If
                    
                End If
            End If
        End If
    'Increment the counter
    counter = counter + 1
    Else
        'Exit the loop if the counter reaches 15
        Exit For
    End If
Next file

Application.DisplayAlerts = True

End Sub

