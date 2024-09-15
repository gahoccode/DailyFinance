Attribute VB_Name = "Module1"

Option Explicit

Sub Update_Case()
    ' Turn off screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' Define the workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("So lieu KCB_Final.xlsx")
    Dim ws As Worksheet
    Set ws = wb.Sheets("Management Report")
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("DATA")
    
    ' Update SUMIFS formulas
    Dim searchDate As Date
    Dim colIndex As Long
    Dim found As Boolean
    Dim cell As Range
    Dim formulaColumn As String
    Dim lastRowC As Long
    Dim lastRowD As Long
    Dim lastRow As Long
    
    ' Get the date from cell B5 in the Management Report sheet
    searchDate = ws.Range("B5").Value
    
    ' Initialize found flag
    found = False
    
    ' Loop through the first row in the DATA sheet to find the matching date
    For Each cell In wsData.Rows(1).Cells
        If cell.Value = searchDate Then
            colIndex = cell.Column
            found = True
            Exit For
        End If
    Next cell
    
    ' Check if the date was found
    If Not found Then
        MsgBox "Date not found in the DATA sheet!", vbExclamation
        Exit Sub
    End If
    
    ' Convert column index to column letter
    formulaColumn = Split(wsData.Cells(1, colIndex).Address, "$")(1)
    
    ' Find the last row with data in column C, starting from row 14
    lastRowC = ws.Range("C14").End(xlDown).Row
    
    ' Find the last row with data in column D in the DATA sheet
    lastRowD = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row
    
    ' Write the SUMIFS formula to cell C14 with absolute references for the sum range
    ws.Cells(14, 3).Formula = "=SUMIFS(DATA!$" & formulaColumn & "$2:$" & formulaColumn & "$" & lastRowD & ", DATA!$D$2:$D$" & lastRowD & ", 'Management Report'!$A14)"
    
    ' Fill down the formula from C14 to the last row
    Dim FillRange As Range
    Set FillRange = ws.Range("C14:C" & lastRowC)
    ws.Range("C14").AutoFill Destination:=FillRange, Type:=xlFillValues
    
    ' Notify about the SUMIFS update
    'MsgBox "Formulas updated successfully!", vbInformation
    
    ' Format ranges to be bold
    Dim boldRanges As Variant
    boldRanges = Array("C13:C14", "C18", "C24", "C33:C36", "C41", "C45", "C54:C57", "C61", "C67")
    
    ' Apply bold formatting to specified ranges
    Dim i As Integer
    For i = LBound(boldRanges) To UBound(boldRanges)
        ws.Range(boldRanges(i)).Font.Bold = True
    Next i
    
    ' List of source and destination ranges for formulas
    Dim ranges As Variant
    ranges = Array("J13", "C13", "J15", "C15", "J19", "C19", "J26", "C26", "J40", "C40", "J45", "C45", "J50", "C50", "J58", "C58", "J61", "C61", "J66", "C66")
    
    ' Copy and paste formulas
    For i = 0 To UBound(ranges) Step 2
        ws.Range(ranges(i)).Copy
        ws.Range(ranges(i + 1)).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    Next i
     
    ' Notify the user
    'MsgBox "Formatting and formulas copied successfully.", vbInformation
    
    '------------------------------------------
    ' TI LE CHI DINH CLS
    ' Get the date from cell B5 in the Management Report sheet
    searchDate = ws.Range("B5").Value
    
    ' Loop through the first row in the DATA sheet to find the matching date
    For Each cell In wsData.Rows(1).Cells
        If cell.Value = searchDate Then
            colIndex = cell.Column
            found = True
            Exit For
        End If
    Next cell
    
    ' Check if the date was found
    If Not found Then
        MsgBox "Date not found in the DATA sheet!", vbExclamation
        Exit Sub
    End If
    
    ' Convert column index to column letter
    formulaColumn = Split(wsData.Cells(1, colIndex).Address, "$")(1)
    
    ' Find the last row with data in column C in the Management Report sheet
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Define the range where the formula will be placed
    Dim formulaRange As Range
    Set formulaRange = ws.Range("C76:C" & lastRow)
    
    ' Write the XLOOKUP formula to cell C76 with the determined column
    ws.Cells(76, 3).Formula = "=XLOOKUP(A76, DATA!$D$2:$D$" & lastRowD & ", DATA!$" & formulaColumn & "$2:$" & formulaColumn & "$" & lastRowD & ", """")"
    
    ' Fill down the formula from C76 to the last row
    ws.Range("C76").AutoFill Destination:=formulaRange, Type:=xlFillValues
        
    ' Turn on screen updating
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Notify the user
    MsgBox "Formulas updated successfully!", vbInformation

End Sub

Sub CopyPasteFormats(ws As Worksheet, sourceCell As String, targetCell As String)
    ws.Range(sourceCell).Copy
    ws.Range(targetCell).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

