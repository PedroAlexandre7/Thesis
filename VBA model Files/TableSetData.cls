VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sheet As Worksheet
Public iType As InstructionType
Public sheetHeader As InputCells
Public firstRow As Long
Public firstColumn As Long
Public columnShift As Long
Public rowShift As Long
Public createSheets As Boolean
Public ClearData As Boolean
