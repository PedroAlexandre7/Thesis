' Start of InputData.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputData"
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


' End of InputData.cls


' Start of OffsetReturnData.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OffsetReturnData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public maxColumnCount As Long
Public currentRowOffset As Long

' End of OffsetReturnData.cls


' Start of SheetData.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public changedRows As New Scripting.Dictionary
Public changedColumns As New Scripting.Dictionary

' End of SheetData.cls


' Start of StringBuilder.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_arrBuffer
Private m_strDelimiter
Private Sub Class_Initialize()
    m_arrBuffer = Array()
    m_strDelimiter = ��
End Sub

Private Sub Class_Terminate()
    m_arrBuffer = Empty
End Sub

Public Property Get Delimiter()
    Delimiter = m_strDelimiter
End Property

Public Property Let Delimiter(strDelimiter)
    m_strDelimiter = strDelimiter
End Property

Public Sub Append(strValue)
    ReDim Preserve m_arrBuffer(UBound(m_arrBuffer) + 1)
    m_arrBuffer(UBound(m_arrBuffer)) = strValue
End Sub

Public Sub AppendLine(strValue)
    Me.Append strValue & vbCrLf
End Sub

Public Sub Compact()
    If Not Me.Delimiter = �� Then
        strOriginalDelimiter = Me.Delimiter
        Me.Delimiter = ��
    End If
    strTemp = Me.ToString
    m_arrBuffer = Array()
    Me.Append strTemp
    Me.Delimiter = strOriginalDelimiter
End Sub

Public Function ToArray()
    ToArray = m_arrBuffer
End Function

Public Function ToString()
    ToString = Join(m_arrBuffer, m_strDelimiter)
End Function

' End of StringBuilder.cls


' Start of TableSetData.cls

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
Public sheetHeader As InputData
Public firstRow As Long
Public firstColumn As Long
Public columnShift As Long
Public rowShift As Long
Public createSheets As Boolean
Public ClearData As Boolean

' End of TableSetData.cls