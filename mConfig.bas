Attribute VB_Name = "mConfig"
'Module for reading the Config worksheet
Public debugMode As Boolean

Public Function GetConfig() As Boolean
    On Error GoTo ErrorHandler

    'Read the value from the Config worksheet cell C7
    Dim configValue As String
    configValue = ThisWorkbook.Worksheets("Config").Cells(3, 3).Value

    'Check if the value is "ON" or "OFF"
    If configValue = "ON" Then
        debugMode = True
        mAppTools.ClearImmediateWindow
    ElseIf configValue = "OFF" Then
        debugMode = False
    Else
        'If the value is not "ON" or "OFF", show an error message and exit the function
        MsgBox "Debug Mode value is not right. Please fix it."
        Exit Function
    End If

    GetConfig = True
    Exit Function

ErrorHandler:
    GetConfig = False
End Function
