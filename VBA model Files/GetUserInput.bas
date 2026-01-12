Attribute VB_Name = "GetUserInput"

Function GetStartingRow() As Integer
    '-store user input in 'sUserInput' variable
    Dim sUserInput As String
    Dim message As String
    Dim row As Integer
    message = "Change starting row" & vbCrLf & "To continue press Enter"
    sUserInput = InputBox(message, "Model Configurator", CStr(DEFAULT_CONFIG_STARTING_ROW))

    'test input before continuing to validate the input
    If sUserInput = "" Then
        Call RestoreStateAndEnd
    ElseIf Not (Len(sUserInput) > 0 And IsNumeric(sUserInput)) Then
        row = DEFAULT_CONFIG_STARTING_ROW
    Else
        row = sUserInput
        If row < DEFAULT_CONFIG_STARTING_ROW Or row > configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).row Then
            row = DEFAULT_CONFIG_STARTING_ROW
        End If
    End If
    GetStartingRow = row
End Function

Function InformCopy(lastSheetName As String, firstSheetName As String, createdSheets As Boolean, Optional askToContinue As Boolean = False) As Boolean
    Dim toContinue As VbMsgBoxResult
    Dim message As String
    If Not createdSheets Then
        If askToContinue Then
            message = "Created tables in '" & lastSheetName & "'." & vbCrLf & vbCrLf & "Continue?"
            toContinue = MsgBox(message, vbYesNo)
        Else
            MsgBox ("Created tables in '" & lastSheetName & "'.")
        End If
    Else
        firstSheetName = IIf(InStr(firstSheetName, CStr(firstYear)) = Len(firstSheetName) - 3, firstSheetName, firstSheetName & " " & firstYear)
        If askToContinue Then
            message = "Created tables in '" & firstSheetName & "' to '" & lastSheetName & "'." & vbCrLf & vbCrLf & "Continue?"
            toContinue = MsgBox(message, vbYesNo)
        Else
            MsgBox ("Created tables in '" & firstSheetName & "' to '" & lastSheetName & "'.")
        End If
    End If
    InformCopy = (toContinue = vbYes)
End Function

Function MsgBoxAskToClearData(sheetName As String) As Boolean
    Dim keepData As VbMsgBoxResult
    Dim message As String
    message = "Keep '" & sheetName & "' sheet data?" & vbCrLf
    keepData = MsgBox(message, vbYesNoCancel)
    If keepData = vbCancel Then
        Call RestoreStateAndEnd
    End If
    MsgBoxAskToClearData = (keepData = vbNo)
End Function

' Retrieves the table set data from the configuration sheet for the specified row.
Function GetTableSetData(row As Long) As TableSetData
    Dim obj As New TableSetData
    Dim firstCell As range
    
    On Error Resume Next
    Set obj.sheet = ThisWorkbook.Sheets(configSheet.Cells(row, InstructionSetting.sheet).Value)
    If obj.sheet Is Nothing And firstYear <> -1 Then
        Set obj.sheet = ThisWorkbook.Sheets(configSheet.Cells(row, InstructionSetting.sheet).Value & " " & firstYear) ' This could lead to unwanted behavior
    End If
    If obj.sheet Is Nothing Then
        MsgBox "Sheet '" & configSheet.Cells(row, InstructionSetting.sheet).Value & "' doesn't exist or it's misspelled." & vbCrLf & "Check " & Col_Letter(3) & row & " in " & configSheet.Name & vbCrLf & "No more tables will be created."
        Call RestoreStateAndEnd
    End If
    
    Set firstCell = obj.sheet.range(IIf(Len(configSheet.Cells(row, InstructionSetting.firstCell).Value) = 0, DEFAULT_FIRST_CELL, configSheet.Cells(row, InstructionSetting.firstCell).Value))
    If firstCell Is Nothing Then
        MsgBox "FirstCell '" & configSheet.Cells(row, InstructionSetting.firstCell).Value & "' is invalid." & vbCrLf & "Check " & Col_Letter(InstructionSetting.firstCell) & row & " in " & configSheet.Name
        Call RestoreStateAndEnd
    End If
    
    obj.iType = GetType(configSheet.Cells(row, InstructionSetting.InstructionType).Value)
    obj.firstRow = firstCell.row
    obj.firstColumn = firstCell.Column
    obj.rowShift = GetInt(configSheet.Cells(row, InstructionSetting.rowShift), 0)
    obj.columnShift = GetInt(configSheet.Cells(row, InstructionSetting.columnShift), 0)
    obj.createSheets = IIf(Len(configSheet.Cells(row, InstructionSetting.createSheets).Value) = 0, False, True)
    If Len(configSheet.Cells(row, InstructionSetting.CopyOutputHeader).Value) <> 0 Then
        Set obj.sheetHeader = New InputData
        Set obj.sheetHeader.range = obj.sheet.range(obj.sheet.Cells(1, 1), obj.sheet.Cells(6, obj.sheet.Cells(1, 15000).End(xlToRight).Column)) ' Just to copy the header format when creating new sheets
        obj.sheetHeader.iType = InstructionType.Header
        obj.sheetHeader.fixedFormulas = True
    End If
    'Println vbCrLf & "    TableSetData" & vbCrLf & "Type: " & EnumName(obj.iType) & vbCrLf & "Start Collumn: " & obj.firstColumn & " Row: " & obj.firstRow & vbCrLf & "DistanceToLast: " & obj.rowShift
    obj.ClearData = GetClearData(configSheet.Cells(row, InstructionSetting.ClearData).Value, obj.sheet.Name, row)
    Set GetTableSetData = obj
End Function

' Retrieves the input cells and their properties from the configuration sheet for the specified row.
Function GetInputData(row As Long) As InputData
    Dim obj As New InputData
    Dim sheet As Worksheet
    Dim firstCell As range
    Dim width As Long
    Dim maxSize As Long
    
    On Error Resume Next: Set sheet = ThisWorkbook.Sheets(configSheet.Cells(row, 3).Value)
    If sheet Is Nothing Then
        MsgBox "Sheet '" & configSheet.Cells(row, InstructionSetting.sheet).Value & "' doesn't exist or it's misspelled." & vbCrLf & "Check " & Col_Letter(InstructionSetting.sheet) & row & " in " & configSheet.Name
        Call RestoreStateAndEnd
    End If
    
    Set firstCell = sheet.range(IIf(Len(configSheet.Cells(row, InstructionSetting.firstCell).Value) = 0, DEFAULT_FIRST_CELL, configSheet.Cells(row, InstructionSetting.firstCell).Value))
    If firstCell Is Nothing Then
        MsgBox "FirstCell '" & configSheet.Cells(row, InstructionSetting.firstCell).Value & "' is invalid." & vbCrLf & "Check " & Col_Letter(InstructionSetting.firstCell) & row & " in " & configSheet.Name
        Call RestoreStateAndEnd
    End If

    width = GetInt(configSheet.Cells(row, InstructionSetting.width), 1)
    maxSize = GetInt(configSheet.Cells(row, InstructionSetting.MaxLength), 15000)
    
    obj.iType = GetType(configSheet.Cells(row, InstructionSetting.InstructionType).Value)
    If obj.iType <> InstructionType.Title Or obj.iType <> InstructionType.Column Then
        obj.rowShift = GetInt(configSheet.Cells(row, InstructionSetting.rowShift), 0)
    End If
    If obj.iType <> InstructionType.Header Then
        obj.columnShift = GetInt(configSheet.Cells(row, InstructionSetting.columnShift), 0)
    End If
    obj.titleIsHeader = IIf(Len(configSheet.Cells(row, InstructionSetting.titleIsHeader).Value) = 0, False, True)
    obj.lastIsTotal = IIf(Len(configSheet.Cells(row, InstructionSetting.lastIsTotal).Value) = 0, False, True)
    obj.countInTotal = IIf(Len(configSheet.Cells(row, InstructionSetting.countInTotal).Value) = 0, False, True)
    obj.fixedFormulas = IIf(Len(configSheet.Cells(row, InstructionSetting.FixedReferences).Value) = 0, False, True)
    obj.hasFormatOnly = IIf(Len(configSheet.Cells(row, InstructionSetting.hasFormatOnly).Value) = 0, False, True)
    Set obj.range = GetInput(sheet, firstCell, width, maxSize, obj.iType)
    'Println vbCrLf & "    InputData" & vbCrLf & "Type: " & EnumName(obj.iType) & vbCrLf & "Columns: " & obj.range.Columns.Count & " Rows: " & obj.range.Rows.Count & vbCrLf & "rowShift: " & obj.rowShift & vbCrLf & "LastIsTotal: " & obj.lastIsTotal & vbCrLf & "CountInTotal: " & obj.countInTotal & vbCrLf & "FixedFormulas: " & obj.fixedFormulas
    Set GetInputData = obj
End Function

' Retrieves the input cells and their properties from the configuration sheet for the specified row.
Function GetStudyYears(row As Long) As InputData
    Dim obj As New InputData
    Dim sheet As Worksheet
    Dim firstCell As range
    Dim maxSize As Long

    On Error Resume Next: Set sheet = ThisWorkbook.Sheets(configSheet.Cells(row, InstructionSetting.sheet).Value)
    If sheet Is Nothing Then
        MsgBox "Sheet '" & configSheet.Cells(row, InstructionSetting.sheet).Value & "' doesn't exist or it's misspelled." & vbCrLf & "Check " & Col_Letter(3) & row & " in " & configSheet.Name
        Call RestoreStateAndEnd
    End If
    
    Set firstCell = sheet.range(IIf(Len(configSheet.Cells(row, InstructionSetting.firstCell).Value) = 0, DEFAULT_FIRST_CELL, configSheet.Cells(row, InstructionSetting.firstCell).Value))
    If firstCell Is Nothing Then
        MsgBox "FirstCell '" & configSheet.Cells(row, InstructionSetting.firstCell).Value & "' is invalid." & vbCrLf & "Check " & Col_Letter(InstructionSetting.firstCell) & row & " in " & configSheet.Name
        Call RestoreStateAndEnd
    End If
    
    maxSize = GetInt(configSheet.Cells(row, InstructionSetting.MaxLength), 15000)
    
    obj.iType = GetType(configSheet.Cells(row, InstructionSetting.InstructionType).Value)
    obj.rowShift = GetInt(configSheet.Cells(row, InstructionSetting.rowShift), 0)
    obj.columnShift = GetInt(configSheet.Cells(row, InstructionSetting.columnShift), 0)
    obj.fixedFormulas = True
    obj.hasFormatOnly = IIf(Len(configSheet.Cells(row, InstructionSetting.hasFormatOnly).Value) = 0, False, True)
    Set obj.range = GetInput(sheet, firstCell, 1, maxSize, obj.iType)
    'Println vbCrLf & "    StudyYears" & vbCrLf & "Type: " & EnumName(obj.iType) & vbCrLf & "N of years: " & obj.Range.Cells.Count & vbCrLf & "RowShift: " & obj.rowShift & vbCrLf & "FixedFormulas: " & obj.fixedFormulas
    firstYear = obj.range.Cells(1).Value
    Set GetStudyYears = obj
End Function

Function GetClearData(str As String, sheetName As String, row As Long) As ClearDataOptions
    str = LCase(str)
    If (str = "ask" Or str = "") Then
        GetClearData = MsgBoxAskToClearData(sheetName)
    ElseIf (str = "clear") Then
        GetClearData = True
    ElseIf (str = "keep") Then
        GetClearData = False
    Else
        MsgBox "Clear Data option '" & str & "' is not valid. Defaulting to 'Ask'." & vbCrLf & "Check " & Col_Letter(InstructionSetting.ClearData) & row & " in " & configSheet.Name
        GetClearData = MsgBoxAskToClearData(sheetName)
    End If
End Function



