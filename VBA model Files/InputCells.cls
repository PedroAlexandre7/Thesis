VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public iType As InstructionType
Public range As range
Public columnShift As Long
Public rowShift As Long
Public titleIsHeader As Boolean
Public lastIsTotal As Boolean
Public countInTotal As Boolean
Public fixedFormulas As Boolean
Public deleteSheet As Boolean
Public hasFormatOnly As Boolean

