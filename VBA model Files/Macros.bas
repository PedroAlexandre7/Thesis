Attribute VB_Name = "Macros"
Sub ShowAllSheets()
Attribute ShowAllSheets.VB_ProcData.VB_Invoke_Func = "r\n14"
If ActiveWorkbook.Sheets.Count > 15 Then
    Application.CommandBars("Workbook tabs").Controls("More Sheets...").Execute
Else
  Application.CommandBars("Workbook tabs").ShowPopup
End If
End Sub

Sub AllCaps()
Attribute AllCaps.VB_ProcData.VB_Invoke_Func = "e\n14"
    For Each myCell In Selection.Cells: myCell.Value = UCase(myCell.Value): Next
End Sub

Sub UpdateSheetsList()
Attribute UpdateSheetsList.VB_ProcData.VB_Invoke_Func = "t\n14"
    
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
Attribute FillRandom.VB_ProcData.VB_Invoke_Func = "u\n14"
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

