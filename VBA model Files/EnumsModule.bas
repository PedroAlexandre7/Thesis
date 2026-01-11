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
