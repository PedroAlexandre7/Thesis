' Start of EnumesModule.bas

Attribute VB_Name = "EnumsModule"
Public configSheet As Worksheet
Public currentYear As Integer
Public firstYear As Integer 'this is kinda fishy, shhhh!
Public sheetsData As New Scripting.Dictionary
Public DEFAULT_CLEAR_STARTING_ROW As Integer
Public DEFAULT_CONFIG_STARTING_ROW As Integer
Public DEFAULT_FIRST_CELL As String

Public Enum InstructionType
    na = -1
    Output = 0
    Header = 1
    Column = 2
    Title = 3
    years = 4
End Enum

Public Enum InstructionSetting
    InstructionType = 2
    sheet = 3
    firstCell = 4
    rowShift = 5
    columnShift = 6
    width = 7
    MaxLength = 8
    titleIsHeader = 9
    createSheets = 10
    CopyOutputHeader = 11
    hasFormatOnly = 12
    ClearData = 13
    FixedReferences = 14
    lastIsTotal = 15
    countInTotal = 16
End Enum

Public Enum SpAddresses
    SheetsListStartR = 3
    SheetsListStartC = 9
End Enum

Public Enum ClearDataOptions
    Ask = 0
    ClearDataOption = 1
    KeepDataOption = 2
End Enum

Function SetGlobalVariables()
    DEFAULT_CLEAR_STARTING_ROW = 7
    DEFAULT_CONFIG_STARTING_ROW = 9
    DEFAULT_FIRST_CELL = "B8"
    firstYear = -1
    Call SetConfigSheet
    Set sheetsData = New Scripting.Dictionary
End Function

Function SetConfigSheet()
    On Error Resume Next
    Set configSheet = ThisWorkbook.Sheets("Model Configurator")
    On Error GoTo 0
    If configSheet Is Nothing Then
        Set configSheet = ModelConfigurator
    End If
End Function

' End of EnumesModule.bas


' Start of GetUserInput.bas

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
        Set obj.sheetHeader = New InputCells
        Set obj.sheetHeader.range = obj.sheet.range(obj.sheet.Cells(1, 1), obj.sheet.Cells(6, obj.sheet.Cells(1, 15000).End(xlToRight).Column)) ' Just to copy the header format when creating new sheets
        obj.sheetHeader.iType = InstructionType.Header
        obj.sheetHeader.fixedFormulas = True
    End If
    'Println vbCrLf & "    TableSetData" & vbCrLf & "Type: " & EnumName(obj.iType) & vbCrLf & "Start Collumn: " & obj.firstColumn & " Row: " & obj.firstRow & vbCrLf & "DistanceToLast: " & obj.rowShift
    obj.ClearData = GetClearData(configSheet.Cells(row, InstructionSetting.ClearData).Value, obj.sheet.Name, row)
    Set GetTableSetData = obj
End Function

' Retrieves the input cells and their properties from the configuration sheet for the specified row.
Function GetInputCells(row As Long) As InputCells
    Dim obj As New InputCells
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
    'Println vbCrLf & "    InputCells" & vbCrLf & "Type: " & EnumName(obj.iType) & vbCrLf & "Columns: " & obj.range.Columns.Count & " Rows: " & obj.range.Rows.Count & vbCrLf & "rowShift: " & obj.rowShift & vbCrLf & "LastIsTotal: " & obj.lastIsTotal & vbCrLf & "CountInTotal: " & obj.countInTotal & vbCrLf & "FixedFormulas: " & obj.fixedFormulas
    Set GetInputCells = obj
End Function

' Retrieves the input cells and their properties from the configuration sheet for the specified row.
Function GetStudyYears(row As Long) As InputCells
    Dim obj As New InputCells
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

' End of GetUserInput.bas


' Start of Macro.bas

Attribute VB_Name = "Macros"
Sub ShowAllSheets()
If ActiveWorkbook.Sheets.Count > 15 Then
    Application.CommandBars("Workbook tabs").Controls("More Sheets...").Execute
Else
  Application.CommandBars("Workbook tabs").ShowPopup
End If
End Sub

Sub AllCaps()
    For Each myCell In Selection.Cells: myCell.Value = UCase(myCell.Value): Next
End Sub

Sub UpdateSheetsList()
    
    Dim mCSup As Worksheet
    Dim i As Integer
    Dim FIRST_SHEET_INDEX As Integer
    Dim tableSheetsCount As Integer
    Dim sheetsCount As Integer
    Dim tableRow As Integer
    Dim tableColumn As Integer
    
    On Error Resume Next
    Set mCSup = ThisWorkbook.Sheets("Model Configurator Sup")
    On Error GoTo 0
    If mCSup Is Nothing Then
        Set mCSup = ModelConfiguratorSup
    End If

    FIRST_SHEET_INDEX = mCSup.Cells(SpAddresses.SheetsListStartR, SpAddresses.SheetsListStartC).Value
    
    tableRow = mCSup.ListObjects("Table_SheetList").range.row
    tableColumn = mCSup.ListObjects("Table_SheetList").range.Column
    tableSheetsCount = mCSup.ListObjects("Table_SheetList").range.Rows.Count - 1 ' to not count with header
    sheetsCount = ThisWorkbook.Sheets.Count - FIRST_SHEET_INDEX ' to count starting from Inputs >>

    ' Delete excess rows
    If sheetsCount < tableSheetsCount Then
        Dim j
        Dim cellsToDelete As range:
        Dim cellsCount As Integer
        
        Set cellsToDelete = mCSup.range(mCSup.Cells(tableRow + sheetsCount, tableColumn), mCSup.Cells(tableRow, tableColumn).End(xlDown))
        cellsCount = cellsToDelete.Rows.Count
        
        For j = 1 To cellsCount
            cellsToDelete.Cells(cellsCount + 1 - j, 1).Delete
        Next j
        
    End If
    
    ' Write numbers
    For i = 1 To sheetsCount
        mCSup.Cells(tableRow + i, tableColumn) = ThisWorkbook.Sheets(FIRST_SHEET_INDEX + i).Name
    Next i
        
End Sub

Sub FillRandom()
    For Each cell In Selection
    cell.Value = Int((30 * Rnd) + 5)
    Next
End Sub


Sub DivideCellsByTwo()
    Dim cell As range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.Value = cell.Value / 2
        End If
    Next cell
End Sub

' End of Macro.bas


' Start of Main.bas

Attribute VB_Name = "Main"
' Main module that calls the creation of tables based on user input and configuration.
Sub CreateTablesFromInput()
Attribute CreateTablesFromInput.VB_ProcData.VB_Invoke_Func = "q\n14"
    
    ' ------ Start optimization ------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    ' --------------------------------
    
    Println "Creating Tables"
    
    Dim headersAndColumns As New Collection
    Dim titleInput As New InputCells
    Dim studyYears As New InputCells
    Dim tableSet As New TableSetData
    'Dim foundFirstCol As Boolean : foundFirstCol = False
    Dim row As Long
    Dim lastRow As Long
    Dim firstOSheetName As String ' if used again may need to reset it in next loop
    Dim toContinue As Boolean

    ' Update Sheets List
    Call UpdateSheetsList
    
    ' Set Global Variables
    Call SetGlobalVariables
    
    ' Set Local Variables
    row = GetStartingRow
    lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).row
    
    ' Set default values
    titleInput.iType = na
    studyYears.iType = na

    For row = row To lastRow
        Select Case GetType(configSheet.Cells(row, InstructionSetting.InstructionType).Value)
            Case InstructionType.Header
                headersAndColumns.Add GetInputCells(row)
            Case Column
                headersAndColumns.Add GetInputCells(row)
            Case Title
                Set titleInput = GetInputCells(row)
            Case years
                Set studyYears = GetStudyYears(row)
            Case InstructionType.Output
                Set tableSet = GetTableSetData(row)
                firstOSheetName = tableSet.sheet.Name
                Call addToSheetsData(tableSet.sheet)
                Call CreateTables(tableSet, studyYears, titleInput, headersAndColumns)

                If row <> lastRow Then
                    toContinue = InformCopy(tableSet.sheet.Name, firstOSheetName, tableSet.createSheets, True)
                    If Not toContinue Then
                        Exit For
                    End If
                Else
                    Call InformCopy(tableSet.sheet.Name, firstOSheetName, tableSet.createSheets)
                End If
                Set headersAndColumns = New Collection
                titleInput.iType = na
                studyYears.iType = na
                firstYear = -1
            Case Else
                MsgBox "Type '" & configSheet.Cells(row, InstructionSetting.InstructionType).Value & "' is not valid. This input will be ignored." & vbCrLf & "Check " & Col_Letter(InstructionSetting.sheet) & row & " in " & configSheet.Name
        End Select
    Next

    ' --- Restore Excel's normal state ---
    Call RestoreState
    ' ------------------------------------
End Sub

' Restore Excel's normal state
Function RestoreState()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Function

Function RestoreStateAndEnd()
    Call RestoreState
    End
End Function

' Creates the tables in the specified output sheet based on the provided inputs and configuration.
Sub CreateTables(tableSet As TableSetData, studyYears As InputCells, ByVal titleInput As InputCells, headersAndColumns As Collection)
    ' Support Variables
    Dim hiddenSheet As Worksheet
    Dim currentRowOffset As Long
    Dim currentColumnOffset As Long: currentColumnOffset = 0
    Dim farthestWrittenColumn As Long
    Dim baseOutputSheetName As String
    Dim i As Long

    ' In case an input is in the outputsheet it copies to that input to the temporary sheet in order to avoid conflicts
    Call PutInputInATempSheet(tableSet, headersAndColumns, hiddenSheet)
    
    ' Prepare Variables
    If studyYears.iType <> na And tableSet.createSheets Then
        If InStr(Len(tableSet.sheet.Name) - Len(studyYears.range.Cells(1).Value), tableSet.sheet.Name, studyYears.range.Cells(1).Value) Then
            baseOutputSheetName = Replace(tableSet.sheet.Name, studyYears.range.Cells(1).Value, "")
        Else
            baseOutputSheetName = tableSet.sheet.Name + " "
            Dim originalSheetNewName As String: originalSheetNewName = baseOutputSheetName & studyYears.range.Cells(1).Value
            If WorksheetExists(originalSheetNewName) Then
                Set tableSet.sheet = ThisWorkbook.Worksheets(originalSheetNewName)
            Else
                tableSet.sheet.Name = originalSheetNewName
            End If
        End If
    Else
        Call ClearData(tableSet.ClearData, tableSet.sheet)
    End If

    ' Create Table(s)
    If studyYears.iType = na Then
        Call CreateTablesForEachTitle(tableSet, titleInput, headersAndColumns, tableSet.firstRow)
    Else
        For i = 1 To studyYears.range.Rows.Count
            currentYear = studyYears.range.Cells(i).Value
            currentRowOffset = tableSet.firstRow
            If tableSet.createSheets Then
                Call AddOrChangeSheetWithYear(tableSet, baseOutputSheetName & studyYears.range.Cells(i).Value)
                Call ClearData(tableSet.ClearData, tableSet.sheet)
                Call CreateTablesForEachTitle(tableSet, titleInput, headersAndColumns, currentRowOffset)
            Else
                farthestWrittenColumn = CreateTablesForEachTitle(tableSet, titleInput, headersAndColumns, currentRowOffset, currentColumnOffset, studyYears, i)
                currentColumnOffset = currentColumnOffset + farthestWrittenColumn + tableSet.columnShift
            End If
        Next i
    End If

    Call DeleteTemporarySheet(hiddenSheet)
    
End Sub

' Adds a new sheet for the given year or renames the existing sheet.
Sub AddOrChangeSheetWithYear(tableSet As TableSetData, newSheetName As String)
    If tableSet.sheet.Name <> newSheetName Then
        If WorksheetExists(newSheetName) Then
            Set tableSet.sheet = ThisWorkbook.Sheets(newSheetName)
        Else
            Set tableSet.sheet = ThisWorkbook.Sheets.Add(After:=tableSet.sheet)
            If Not tableSet.sheetHeader Is Nothing Then
                'Call addToSheetsData(tableSet.sheet) ' Fishy, but easy fix, just make it check it exist, if it doen
                Call CopyCells(tableSet.sheet, tableSet.sheetHeader, 1, 1, , False, False)
                'sheetsData.Remove tableSet.sheet
            End If
            tableSet.sheet.Name = newSheetName
        End If
    End If
    Call addToSheetsData(tableSet.sheet)
End Sub

Sub addToSheetsData(sheet As Worksheet)
    If Not sheetsData.Exists(sheet) Then
        sheetsData.Add sheet, New SheetData
    End If
End Sub

Sub ClearData(hasToClearData As Boolean, sheet As Worksheet)
    If hasToClearData Then
        sheet.Rows(DEFAULT_CLEAR_STARTING_ROW & ":" & Rows.Count).Clear
        With sheet
            .Rows(DEFAULT_CLEAR_STARTING_ROW & ":" & Rows.Count).RowHeight = 15
            .Cells.ColumnWidth = 8.43
        End With
    End If
End Sub

' Creates tables for each year based on the provided inputs and configuration.
Function CreateTablesForEachTitle(tableSet As TableSetData, ByVal titleInput As InputCells, headersAndColumns As Collection, currentRowOffset As Long, Optional currentColumnOffset As Long = 0, Optional studyYears As InputCells = Nothing, Optional row As Long = 1) As Long
    Dim k As Long
    Dim returnedOffsetData As OffsetReturnData
    Dim maxColumnCount As Long: maxColumnCount = 0
    
    If titleInput.iType = na Then
        ' Fill Years
        If Not studyYears Is Nothing Then
            Call CopyCells(tableSet.sheet, studyYears, currentRowOffset, tableSet.firstColumn + studyYears.columnShift + currentColumnOffset, row)
            currentRowOffset = currentRowOffset + 1 + studyYears.rowShift
        End If
        Set returnedOffsetData = CreateHeadersAndColumns(tableSet, headersAndColumns, currentRowOffset, currentColumnOffset)
        maxColumnCount = WorksheetFunction.Max(maxColumnCount, returnedOffsetData.maxColumnCount)
    Else
        For k = 1 To titleInput.range.Rows.Count
            ' Fill Years
            If Not studyYears Is Nothing Then
                Call CopyCells(tableSet.sheet, studyYears, currentRowOffset, tableSet.firstColumn + studyYears.columnShift + currentColumnOffset, row)
                currentRowOffset = currentRowOffset + 1 + studyYears.rowShift
            End If
            ' Fill Title
            Call CopyCells(tableSet.sheet, titleInput, currentRowOffset, tableSet.firstColumn + titleInput.columnShift + currentColumnOffset, k, Not titleInput.titleIsHeader)
            currentRowOffset = currentRowOffset + IIf(titleInput.titleIsHeader, 0, 1) + titleInput.rowShift
            Set returnedOffsetData = CreateHeadersAndColumns(tableSet, headersAndColumns, currentRowOffset, currentColumnOffset, IIf(titleInput.titleIsHeader, titleInput.range.Columns.Count + titleInput.columnShift, 0))
            currentRowOffset = returnedOffsetData.currentRowOffset
            maxColumnCount = WorksheetFunction.Max(maxColumnCount, returnedOffsetData.maxColumnCount)
        Next k
    End If
    CreateTablesForEachTitle = maxColumnCount
End Function

' Copies inputs that are in the output sheet to a temporary hidden sheet to avoid conflicts
Sub PutInputInATempSheet(tableSet As TableSetData, headersAndColumns As Collection, hiddenSheet As Worksheet)
    Dim rowOffset As Long: rowOffset = 500000
    Dim columnOffset As Long: columnOffset = 8000
    Dim inputO As InputCells
    
    Set changedRows = New Scripting.Dictionary
    Set changedColumns = New Scripting.Dictionary
    
    For i = 1 To headersAndColumns.Count
        If headersAndColumns(i).range.Worksheet Is tableSet.sheet Then
            If hiddenSheet Is Nothing Then
                Set hiddenSheet = ThisWorkbook.Sheets.Add
                sheetsData.Add hiddenSheet, New SheetData
            End If
            Set inputO = headersAndColumns(i)
            Set inputO.range = CopyCells(hiddenSheet, inputO, rowOffset, columnOffset)
            rowOffset = rowOffset + inputO.range.Rows.Count + 1
            columnOffset = columnOffset + inputO.range.Columns.Count + 1
        End If
    Next
End Sub

' Deletes the temporary sheet if it exists
Sub DeleteTemporarySheet(hiddenSheet As Worksheet)
    If Not hiddenSheet Is Nothing Then
        Application.DisplayAlerts = False
        hiddenSheet.Delete
        Application.DisplayAlerts = True
    End If
End Sub

' Creates headers and columns in the output sheet based on the provided inputs and configuration.
Function CreateHeadersAndColumns(tableSet As TableSetData, headersAndColumns As Collection, currentRowOffset As Long, currentColumnOffset As Long, Optional spaceForTitle As Integer = 0) As OffsetReturnData
    Dim columnHeightToAdd As Long: columnHeightToAdd = 0  ' Makes sure when writing next Headers, they start after the columns.
    Dim offsetData As OffsetReturnData: Set offsetData = New OffsetReturnData
    Dim nextColumn As InputCells
    Dim firstCell As range
    Dim lastHeader As range
    Dim rangeToFill As range
    
    For i = 1 To headersAndColumns.Count
        Dim obj As InputCells: Set obj = headersAndColumns(i)
        If obj.iType = Header Then
            ' Fill in Headers
            offsetData.maxColumnCount = WorksheetFunction.Max(offsetData.maxColumnCount, obj.range.Columns.Count + obj.columnShift + spaceForTitle)
            currentRowOffset = currentRowOffset + IIf(spaceForTitle <> 0 And i = 1, 0, obj.rowShift) + columnHeightToAdd
            Call CopyCells(tableSet.sheet, obj, currentRowOffset, tableSet.firstColumn + currentColumnOffset + spaceForTitle)
            currentRowOffset = currentRowOffset + obj.range.Rows.Count
            columnHeightToAdd = 0
            'Fill Totals
            If (obj.lastIsTotal And i < headersAndColumns.Count) Then

                For j = i + 1 To headersAndColumns.Count 'TESTAR ?
                    If (headersAndColumns(j).iType = InstructionType.Column And headersAndColumns(j).countInTotal = True) Then
                        Set nextColumn = headersAndColumns(j)
                        Exit For
                    End If
                Next
                If (Not nextColumn Is Nothing) Then
                    Set firstCell = tableSet.sheet.Cells(currentRowOffset, tableSet.firstColumn + obj.range.Columns.Count - 1)
                    Set rangeToFill = tableSet.sheet.range(firstCell, tableSet.sheet.Cells(firstCell.row + nextColumn.range.Rows.Count - 1, firstCell.Column))
                    Call FillTotal(obj.iType, rangeToFill, obj.range.Columns.Count - nextColumn.range.Columns.Count - nextColumn.columnShift - 1)
                End If
            End If
            Set lastHeader = obj.range
        Else
            ' Fill in Column
            offsetData.maxColumnCount = WorksheetFunction.Max(offsetData.maxColumnCount, obj.range.Columns.Count + obj.columnShift)
            Call CopyCells(tableSet.sheet, obj, currentRowOffset, tableSet.firstColumn + obj.columnShift + currentColumnOffset)
            columnHeightToAdd = WorksheetFunction.Max(columnHeightToAdd, obj.range.Rows.Count)
            'Fill Totals
            If obj.lastIsTotal And Not lastHeader Is Nothing Then
                Set firstCell = tableSet.sheet.Cells(currentRowOffset + obj.range.Rows.Count - 1, tableSet.firstColumn + obj.columnShift + obj.range.Columns.Count)
                Dim headersLenght As Long
                If (i < headersAndColumns.Count) Then
                    Set nextColumn = headersAndColumns(i + 1)
                    headersLenght = IIf(nextColumn.iType = InstructionType.Column And (nextColumn.countInTotal = False Or nextColumn.range.Rows.Count = obj.range.Rows.Count), nextColumn.columnShift, lastHeader.Columns.Count)
                Else
                    headersLenght = lastHeader.Columns.Count
                End If
                Set rangeToFill = tableSet.sheet.range(firstCell, tableSet.sheet.Cells(firstCell.row, firstCell.Column + headersLenght - obj.columnShift - obj.range.Columns.Count - 1))
                Call FillTotal(obj.iType, rangeToFill, obj.range.Rows.Count - 1)
            End If
        End If
        '        Set nextColumn = Nothing                 'May bug something!!?
    Next
    offsetData.currentRowOffset = currentRowOffset + (columnHeightToAdd + tableSet.rowShift)
    Set CreateHeadersAndColumns = offsetData
End Function

' Fills the specified range with SUM formulas based on the instruction type and sum range.
Sub FillTotal(objtype As InstructionType, rangeToFill As range, sumRange As Long)
    For i = 1 To rangeToFill.Rows.Count
        For j = 1 To rangeToFill.Columns.Count
            If objtype = Column Then
                rangeToFill.Cells(i, j).formulaR1C1 = "=sum(r[-" & sumRange & "]c[0]:r[-1]c[0])"
            Else
                rangeToFill.Cells(i, j).formulaR1C1 = "=sum(r[0]c[-" & sumRange & "]:r[0]c[-1])"
            End If
        Next j
    Next i
End Sub

' Copies the specified input cells to the output sheet at the given offsets, adjusting formulas and dimensions as needed.
Function CopyCells(outputSheet As Worksheet, obj As InputCells, currentRowOffset As Long, columnOffset As Long, Optional row As Long = -1, Optional saveRowHeight As Boolean = True, Optional saveColumnWidth As Boolean = True) As range
    Dim objRange As range
    If row = -1 Then
        Set objRange = obj.range
    Else
        Set objRange = obj.range.Rows(row)
    End If

    objRange.Copy
    If obj.hasFormatOnly Then
        outputSheet.Cells(currentRowOffset, columnOffset).PasteSpecial Paste:=xlPasteFormats
    Else
        outputSheet.Cells(currentRowOffset, columnOffset).PasteSpecial Paste:=xlPasteAll
    End If
    

    ' --- Processing ---
    ' For maximum performance, calculate the row/column shift only once.
    Dim destRange As range: Set destRange = outputSheet.range(outputSheet.Cells(currentRowOffset, columnOffset), outputSheet.Cells(currentRowOffset + objRange.Rows.Count - 1, columnOffset + objRange.Columns.Count - 1))
    Dim deltaRow As Long, deltaCol As Long
    deltaRow = destRange.Cells(1, 1).row - objRange.Cells(1, 1).row
    deltaCol = destRange.Cells(1, 1).Column - objRange.Cells(1, 1).Column
    If Not saveColumnWidth And obj.iType = Header Then
        Set objRange = objRange.Resize(objRange.Rows.Count, 1) ' Just to copy the row height
        Exit Function
    End If

    For i = 0 To objRange.Rows.Count - 1
        For j = 0 To objRange.Columns.Count - 1
            Call SetFormulasAndHW(obj.iType, obj.fixedFormulas, outputSheet.Cells(currentRowOffset + i, columnOffset + j), objRange.Cells(i + 1, j + 1), deltaRow, deltaCol, IIf(saveRowHeight, sheetsData.Item(outputSheet).changedRows, New Scripting.Dictionary), IIf(saveColumnWidth, sheetsData.Item(outputSheet).changedColumns, New Scripting.Dictionary))
        Next j
    Next i
    Set CopyCells = destRange
End Function

' Sets formulas and adjusts row height and column width for the specified cell.
Sub SetFormulasAndHW(objtype As InstructionType, objFixedFormulas As Boolean, cellToModify As range, cellWithData As range, ByVal deltaRow As Long, ByVal deltaCol As Long, changedRows As Scripting.Dictionary, changedColumns As Scripting.Dictionary)
    If cellToModify.HasFormula Then
    End If
    If cellWithData.HasFormula And Not objFixedFormulas Then
        Call CopyFormulasWithAbsoluteShift(objtype, cellWithData, cellToModify, deltaRow, deltaCol)
    End If
    Call setHeightAndWidth(objtype, cellToModify, cellWithData, changedRows, changedColumns)
End Sub

Public Sub CopyFormulasWithAbsoluteShift(objtype As InstructionType, ByVal sourceRange As range, ByVal destRange As range, ByVal deltaRow As Long, ByVal deltaCol As Long)
    ' Copies formulas from a source range to a destination range, adjusting ALL
    ' references (relative, mixed, and absolute) based on the change in position.
    ' This method is highly efficient for large ranges.
    ' --- Validation ---
    If sourceRange Is Nothing Or destRange Is Nothing Then
        MsgBox "Source and Destination ranges must be valid.", vbCritical, "Error"
        Exit Sub
    End If

    ' Write the entire array of new formulas back to the sheet in one operation.
    ' Excel will raise an error here if a calculated reference is invalid (e.g., R0 or R-5).
    On Error Resume Next
    destRange.formulaR1C1 = ConvertFormulaR1C1(objtype, sourceRange.formulaR1C1, sourceRange.row, sourceRange.Column, deltaRow, deltaCol)
    If Err.Number <> 0 Then
        MsgBox "An error occurred while writing the updated formulas." & vbCrLf & _
            "This can happen if a formula shift results in an invalid reference (e.g., row 0 or a negative row/column).", _
            vbCritical, "Formula Error"
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub

Private Function ConvertFormulaR1C1(objtype As InstructionType, ByVal formulaR1C1 As String, sourceRow As Long, sourceColumn As Long, ByVal deltaRow As Long, ByVal deltaCol As Long) As String
    ' It parses a formula string in R1C1 notation and adjusts any absolute
    ' row or column references by the provided delta values.
    ' It correctly ignores relative references (e.g., R[1]) and references
    ' inside string literals (e.g., ="Report for R2C3").
    Dim result As New StringBuilder
    Dim i As Long
    Dim char As String
    i = 1
    Do While i <= Len(formulaR1C1)
        char = Mid(formulaR1C1, i, 1)
        
        Select Case char
            Case """"
                ' Found the start of a string literal.
                ' Find the closing quote and append the entire literal without parsing.
                Dim endQuotePos As Long
                endQuotePos = InStr(i + 1, formulaR1C1, """")

                If endQuotePos > 0 Then
                    result.Append Mid(formulaR1C1, i, endQuotePos - i + 1)
                    i = endQuotePos
                Else
                    ' Malformed formula with an unclosed quote, append the rest of the string.
                    result.Append Mid(formulaR1C1, i)
                    i = Len(formulaR1C1)
                End If
                
            Case "R", "C"
                ' Found a potential row or column reference.
                Dim prevChar As String
                If i = 1 Then
                    prevChar = ""
                Else
                    prevChar = Mid(formulaR1C1, i - 1, 1)
                End If


                If i = Len(formulaR1C1) Then
                    ' R or C is the last character in the formula.
                    result.Append char
                    GoTo NextIteration
                End If
                
                Dim nextChar As String
                nextChar = Mid(formulaR1C1, i + 1, 1)

                Dim j As Long
                If prevChar = "!" And IsNumeric(nextChar) Then
                    ' This is a absolute reference to another sheet 'Network Parameters'!R4C2.
                    ' These do not need to be changed. Append the entire segment as-is.
                    Dim lastRefNumPos As Long
                    j = i + 1
                    Do While j <= Len(formulaR1C1) And (IsNumeric(Mid(formulaR1C1, j, 1)) Or Mid(formulaR1C1, j, 1) = "[")
                        'Println Mid(formulaR1C1, j, 1)
                        lastRefNumPos = j
                        j = j + 1
                    Loop
                    If lastRefNumPos = 0 Or Mid(formulaR1C1, j, 1) <> "C" Then
                        ' Malformed or not a reference, just append the character.
                        result.Append char
                        GoTo NextIteration
                    End If
                    ' reset index so that can be verified later
                    lastRefNumPos = 0
                    j = j + 1
                    Do While j <= Len(formulaR1C1) And IsNumeric(Mid(formulaR1C1, j, 1))
                        lastRefNumPos = j
                        j = j + 1
                    Loop
                    If lastRefNumPos = 0 Then
                        ' Malformed, just append the character.
                        result.Append char
                    Else
                        result.Append Mid(formulaR1C1, i, lastRefNumPos - i + 1)
                        i = lastRefNumPos
                    End If
                ElseIf nextChar = "[" Then
                    ' This is a relative reference (e.g., R[1] or C[-2]).
                    ' These do not need to be changed. Append the entire segment as-is.
                    Dim endBracketPos As Long
                    endBracketPos = InStr(i + 1, formulaR1C1, "]")
                    If endBracketPos > 0 Then
                        result.Append Mid(formulaR1C1, i, endBracketPos - i + 1)
                        i = endBracketPos
                    Else
                        ' Malformed, just append the character.
                        result.Append char
                    End If
                ElseIf IsNumeric(nextChar) Then
                    ' This is an absolute reference (e.g., R1, C5).
                    ' Absolute references are always positive.
                    ' Extract the number, adjust it by the delta, and rebuild.
                    Dim numStr As String
                    j = i + 1
                    numStr = ""

                    ' Extract the full numeric part.
                    Do While j <= Len(formulaR1C1) And IsNumeric(Mid(formulaR1C1, j, 1))
                        numStr = numStr & Mid(formulaR1C1, j, 1)
                        j = j + 1
                    Loop

                    If Len(numStr) > 0 Then
                        Dim originalNum As Long, newNum As Long
                        originalNum = CLng(numStr)
                        ' Apply the appropriate delta.
                        If char = "R" Then
                            If objtype = Title Then
                                newNum = originalNum + deltaRow + (sourceRow - originalNum)
                            Else
                                newNum = originalNum + deltaRow
                            End If
                        Else ' char = "C"
                            newNum = originalNum + deltaCol
                        End If

                        result.Append char & CStr(newNum)
                        i = j - 1 ' Move parser index to the end of the processed number.
                    Else
                        ' Should not happen if IsNumeric(nextChar) is true, but as a fallback.
                        result.Append char
                    End If
                Else
                    ' Not a standard reference (e.g., "RC" in a named range). Treat as literal.
                    result.Append char
                End If
                
            Case Else
                ' Any other character, append directly.
                result.Append char
        End Select
NextIteration:
        i = i + 1
    Loop

    ConvertFormulaR1C1 = result.ToString
End Function

' Sets the height and width of the specified cell based on the object type and changed rows/columns.
Function setHeightAndWidth(objtype As InstructionType, cellToModify As range, cellWithData As range, changedRows As Scripting.Dictionary, changedColumns As Scripting.Dictionary)

    If Not changedRows.Exists(cellToModify.row) Then
        cellToModify.RowHeight = cellWithData.RowHeight
        changedRows.Add cellToModify.row, cellWithData.RowHeight
    ElseIf changedRows.Item(cellToModify.row) < cellWithData.RowHeight Then
        cellToModify.RowHeight = cellWithData.RowHeight
        changedRows.Item(cellToModify.row) = cellWithData.RowHeight
    End If

    If Not changedColumns.Exists(cellToModify.Column) Then
        cellToModify.ColumnWidth = cellWithData.ColumnWidth
        changedColumns.Add cellToModify.Column, cellWithData.ColumnWidth
    ElseIf changedColumns.Item(cellToModify.Column) < cellWithData.ColumnWidth Then
        cellToModify.ColumnWidth = cellWithData.ColumnWidth
        changedColumns.Item(cellToModify.Column) = cellWithData.ColumnWidth
    End If
End Function

' GetSize function calculates the size of the input range based on the specified parameters
Function GetSize(inputSheet As Worksheet, firstCell As range, width As Long, maxSize As Long, objtype As InstructionType) As Long
    Dim lineSize() As Long: ReDim lineSize(0 To width - 1)
    Dim lastPossibleCell As range
    If objtype = InstructionType.Header Then
        For i = 0 To width - 1
            Set lastPossibleCell = inputSheet.Cells(firstCell.row + i, firstCell.Column + maxSize - 1)
            If IsEmpty(lastPossibleCell) Then
                lineSize(i) = lastPossibleCell.End(xlToLeft).Column - (firstCell.Column - 1)
            Else
                lineSize(i) = lastPossibleCell.Column - (firstCell.Column - 1)
            End If
        Next
    Else
        For i = 0 To width - 1
            Set lastPossibleCell = inputSheet.Cells(firstCell.row + maxSize - 1, firstCell.Column + i)
            If IsEmpty(lastPossibleCell) Then
                lineSize(i) = lastPossibleCell.End(xlUp).row - (firstCell.row - 1)
            Else
                lineSize(i) = lastPossibleCell.row - (firstCell.row - 1)
            End If
        Next
    End If
    GetSize = WorksheetFunction.Max(lineSize)
End Function

' GetInput function retrieves a range of cells based on the specified parameters
Function GetInput(inputSheet As Worksheet, firstCell As range, width As Long, maxSize As Long, objtype As InstructionType) As range
    Dim length As Long: length = GetSize(inputSheet, firstCell, width, maxSize, objtype)
    Dim lastCell As range
    If objtype = InstructionType.Header Then
        Set lastCell = inputSheet.Cells(firstCell.row + width - 1, firstCell.Column + length - 1)
    Else
        Set lastCell = inputSheet.Cells(firstCell.row + length - 1, firstCell.Column + width - 1)

    End If
    Set GetInput = inputSheet.range(firstCell, lastCell)
End Function

' GetInt function retrieves an integer value from a string, with an optional default value
Function GetInt(str As String, Optional Default As Long = 1) As Long
    If IsNumeric(str) Then
        If str > 0 Then
            GetInt = str
        Else
            GetInt = Default
        End If
    Else
        GetInt = Default
    End If
End Function

' configSheet global variable should have been set before function call
Function GetType(str As String) As InstructionType
    str = LCase(str)
    If (str = "output") Then
        GetType = Output
    ElseIf (str = "header") Then
        GetType = Header
    ElseIf (str = "column") Then
        GetType = Column
    ElseIf (str = "title") Then
        GetType = Title
    ElseIf (str = "years") Then
        GetType = years
    Else
        GetType = na
    End If
End Function

' Verifies if a worksheet with the given name exists in the workbook
Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = (ThisWorkbook.Worksheets(shtName).Name = shtName)
End Function

' Converts a column number to its corresponding letter representation
Function Col_Letter(lngCol As Long) As String
    Col_Letter = Split(Cells(1, lngCol).Address(True, False), "$")(0)
End Function

' Returns the string representation of an InstructionType enumeration value
Function EnumName(i As Long) As String
    EnumName = Array("na", "Output", "Header", "Column", "Title", "Years")(i + 1)
End Function

' Prints a string to the debug output
Sub Println(str As String)
    Debug.Print str
End Sub

' End of Main.bas