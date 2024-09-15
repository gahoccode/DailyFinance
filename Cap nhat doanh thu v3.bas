Attribute VB_Name = "Module1"

Option Explicit

Sub Update_Revenue()
    Dim wb As Workbook
    Dim wsDaily As Worksheet
    Dim wsReport As Worksheet
    Dim searchDate As Date
    Dim lastRow As Long
    Dim lastRowA As Long
    Dim colIndex As Long
    Dim formulaColumn As String
    Dim cell As Range
    Dim properDateFormat As Boolean
    
    ' Turn off screen updating and set manual calculation for optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set reference to the workbook
    Set wb = Workbooks("Daily Revenue 2024.xlsx")
    
    ' Set references to the sheets
    Set wsDaily = wb.Sheets("Daily")
    Set wsReport = wb.Sheets("Management Report")
    
    ' Get the date from cell D7 in the Management Report sheet
    searchDate = wsReport.Range("D7").Value
    
    ' Loop through row 4 in the Daily sheet to find the matching date
    properDateFormat = False
    For Each cell In wsDaily.Rows(4).Cells
        If IsDate(cell.Value) Then
            properDateFormat = True
            If cell.Value = searchDate Then
                colIndex = cell.Column
                Exit For
            End If
        End If
    Next cell
    
    ' Convert column index to column letter
    formulaColumn = Split(wsDaily.Cells(4, colIndex).Address, "$")(1)
    
    ' Find the last row with data in column A of the Daily sheet using xlDown
    lastRowA = wsDaily.Cells(5, "A").End(xlDown).Row
    
    ' Find the last row with data in column D of the Management Report sheet, starting from cell D10
    lastRow = wsReport.Cells(wsReport.Rows.Count, "D").End(xlUp).Row
    
    ' Write the XLOOKUP formula to cell D11 with the determined last row of column A
    wsReport.Range("D11").Formula = "=XLOOKUP(A11, 'Daily'!$A$5:$A$" & lastRowA & ", 'Daily'!" & formulaColumn & "$5:$" & formulaColumn & "$" & lastRowA & ", """")"
    
    ' Fill down the formula from D11 to the last row without changing cell format
    wsReport.Range("D11").AutoFill Destination:=wsReport.Range("D11:D" & lastRow), Type:=xlFillValues
    
    ' Turn on screen updating and reset calculation to automatic
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Display success message
    MsgBox "Formula updated and formatting applied successfully!", vbInformation
End Sub

