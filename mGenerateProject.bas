Attribute VB_Name = "mGenerateProject"
Sub GenerateProject()
Attribute GenerateProject.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.DisplayAlerts = False

    Sheets.Add(After:=Sheets("Sheet1")).Name = "Main"

    Range("B2").Select
    ActiveSheet.Buttons.Add(48, 14.5, 192, 58).Select
    Selection.OnAction = "Run"
    Selection.Characters.Text = "Run"
    With Selection.Characters(Start:=1, Length:=3).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    
    Range("E12").Select
    Sheets.Add(After:=Sheets("Main")).Name = "Config"

    Range("C2").Select
    ActiveCell.FormulaR1C1 = "ON / OFF"
    
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Debug Mode"
    
    Columns("B:B").EntireColumn.AutoFit
    
    Range("C3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="ON, OFF"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    ActiveCell.SpecialCells(xlCellTypeSameValidation).Select

    Range("C3").Select
    ActiveCell.FormulaR1C1 = "ON"
    
    Sheets("Sheet1").Select
    ActiveWindow.SelectedSheets.Delete
    
Application.DisplayAlerts = True
    
End Sub
