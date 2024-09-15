Attribute VB_Name = "Module2"
Option Explicit
Private wsHRDrug As Worksheet
Private wsDept As Worksheet
Private wsHR As Worksheet
Private wsChart As Worksheet
Public wb As Workbook

Sub Initialize()
    ' Initialize the module-level variables
    Set wb = Workbooks("Monthly P&L 2024_PCSG.xlsx")
    Set wsDept = wb.Sheets("Departments")
    Set wsHRDrug = wb.Sheets("Thuoc-VTTH")
    Set wsHR = wb.Sheets("HR")
    Set wsChart = wb.Sheets("Chart")
End Sub

Sub English_Version()
' Ensure global variables are initialized
Initialize
' Run each subroutine in the module
Chart_English
RadarChart_English
HR_English
Thuoc_English
English_PL
End Sub
Sub Vietnamese_Version()
' Ensure global variables are initialized
Initialize
' Run each subroutine in the module
Chart_VN
RadarChart_VN
HR_VN
Thuoc_VN
Vietnam_PL
End Sub
Sub Thuoc_VN()
     
    ' Ensure the workbook and worksheet are initialized
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Update values in the specified cells
    wsHRDrug.Range("G6").Value = "Thu" & ChrW(7889) & "c"
    wsHRDrug.Range("G7").Value = "V" & ChrW(7853) & "t t" & ChrW(432) & _
        " tiêu hao và hóa ch" & ChrW(7845) & "t"
    wsHRDrug.Range("G8").Value = "T" & ChrW(7893) & "ng "
    wsHRDrug.Range("G9").Value = "T" & ChrW(7881) & " l" & ChrW(7879) & _
        " % trên doanh thu"
    wsHRDrug.Range("G10").Value = "T" & ChrW(7893) & "ng thu" & ChrW(7889) & _
        "c và VTTH"
    wsHRDrug.Range("G11").Value = "Chi phí thu" & ChrW(7889) & "c"
    wsHRDrug.Range("G12").Value = "Chi phí VTTH"
    wsHRDrug.Range("G14").Value = "T" & ChrW(7881) & " l" & ChrW(7879) & _
        " % chi phí thu" & ChrW(7889) & "c trên doanh thu thu" & ChrW(7889) & "c"
        
    wsHRDrug.Range("H4").Value = "Chi phí thu" & ChrW(7889) & "c và v" & ChrW(7853 _
        ) & "t t" & ChrW(432) & " tiêu hao"
    wsHRDrug.Range("K5").Value = "Ch" & ChrW(7881) & " tiêu"
    wsHRDrug.Range("M5").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsHRDrug.Range("N5").Value = "So v" & ChrW(7899) & "i k" & ChrW(7871) & " ho" _
        & ChrW(7841) & "ch"
    
    ' Restore screen updating and calculation settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub Thuoc_English()
    ' Ensure the workbook and worksheet are initialized
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' Update values in the specified cells
    wsHRDrug.Range("G6").Value = "Pharmacy"
    wsHRDrug.Range("G7").Value = "Consumable & Chemical"
    wsHRDrug.Range("G8").Value = "Total"
    wsHRDrug.Range("G9").Value = "% of Net Revenue"
    wsHRDrug.Range("G10").Value = "Total Pharmacy, Cons. & Chemical Cost"
    wsHRDrug.Range("G11").Value = "Pharmacy Cost"
    wsHRDrug.Range("G12").Value = "Consumable & Chemical Cost"
    wsHRDrug.Range("G14").Value = "Pharmacy cost, % of Pharmacy Revenue"
    wsHRDrug.Range("H4").Value = "PCSG Pharmacy, Consumable and Chemical Cost (VND bn)"
    wsHRDrug.Range("K5").Value = "Target"
    wsHRDrug.Range("M5").Value = "Last Month Variance"
    wsHRDrug.Range("N5").Value = "Target Variance"
    ' Restore screen updating and calculation settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub Vietnam_PL()
    ' Ensure the workbook and worksheet are initialized
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Clear contents from the target range
    wsDept.Range("E7:E25").ClearContents
    wsDept.Range("D7:D25").ClearContents
    ' Populate the cells with specific text and add a small wait time after each line
    wsDept.Range("E7").Value = "N" & ChrW(7897) & "i trú"
    wsDept.Range("E8").Value = "S" & ChrW(7843) & "n sanh"
    wsDept.Range("E9").Value = "S" & ChrW(7843) & "n d" & ChrW(432) & ChrW(7905) & "ng"
    wsDept.Range("E11").Value = "Ngo" & ChrW(7841) & "i trú"
    wsDept.Range("E12").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    wsDept.Range("E13").Value = "N" & ChrW(7897) & "i trú"
    wsDept.Range("E14").Value = "Ngo" & ChrW(7841) & "i trú"
    wsDept.Range("E15").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    wsDept.Range("E16").Value = "N" & ChrW(7897) & "i trú"
    wsDept.Range("E17").Value = "NICU"
    wsDept.Range("E18").Value = "Ngo" & ChrW(7841) & "i trú"
    wsDept.Range("E19").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    wsDept.Range("E20").Value = "N" & ChrW(7897) & "i trú"
    wsDept.Range("E21").Value = "Ngo" & ChrW(7841) & "i trú"
    wsDept.Range("E22").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    wsDept.Range("E23").Value = "N" & ChrW(7897) & "i trú"
    wsDept.Range("E24").Value = "Ngo" & ChrW(7841) & "i trú"
    wsDept.Range("E25").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    ' Populate the cells in column D
    wsDept.Range("D6").Value = "Chuyên khoa"
    wsDept.Range("D7").Value = "S" & ChrW(7843) & "n khoa"
    wsDept.Range("D13").Value = "Ph" & ChrW(7909) & " khoa"
    wsDept.Range("D16").Value = "Nhi khoa"
    wsDept.Range("D20").Value = ChrW(272) & "a khoa"
    wsDept.Range("D23").Value = "IVF"
    wsDept.Range("D26").Value = "T" & ChrW(7893) & "ng c" & ChrW(7897) & "ng"
    
    ' Populate headers in other columns
    wsDept.Range("O5").Value = "S" & ChrW(7889) & " ca"
    wsDept.Range("V5").Value = "Bill bình quân (tri" & ChrW(7879) & "u " & ChrW(273) & ChrW(7891) & "ng)"
    wsDept.Range("H5").Value = "Doanh thu (t" & ChrW(7881) & " " & ChrW(273) & ChrW(7891) & "ng)"
    wsDept.Range("I6").Value = "Ch" & ChrW(7881) & " tiêu"
    wsDept.Range("P6").Value = "Ch" & ChrW(7881) & " tiêu"
    wsDept.Range("L6").Value = "% Th" & ChrW(7921) & "c hi" & ChrW(7879) & "n k" _
         & ChrW(7871) & " ho" & ChrW(7841) & "ch"
    wsDept.Range("S6").Value = "% Th" & ChrW(7921) & "c hi" & ChrW(7879) & "n k" _
         & ChrW(7871) & " ho" & ChrW(7841) & "ch"
    wsDept.Range("Z6").Value = "% Th" & ChrW(7921) & "c hi" & ChrW(7879) & "n k" _
         & ChrW(7871) & " ho" & ChrW(7841) & "ch"
    wsDept.Range("K6").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsDept.Range("R6").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsDept.Range("Y6").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    
End Sub

Sub English_PL()
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Clear contents from the target range
    wsDept.Range("E7:E25").ClearContents
    wsDept.Range("D7:D25").ClearContents
    ' Populate the cells with specific text and add a small wait time after each line
    wsDept.Range("E7").Value = "Inpatient"
    wsDept.Range("E8").Value = "Delivery"
    wsDept.Range("E9").Value = "Threatened Preterm Labor"
    wsDept.Range("E11").Value = "Outpatient"
    wsDept.Range("E12").Value = "Total"
    wsDept.Range("E13").Value = "Inpatient"
    wsDept.Range("E14").Value = "Outpatient"
    wsDept.Range("E15").Value = "Total"
    wsDept.Range("E16").Value = "Inpatient"
    wsDept.Range("E17").Value = "NICU"
    wsDept.Range("E18").Value = "Outpatient"
    wsDept.Range("E19").Value = "Total"
    wsDept.Range("E20").Value = "Inpatient"
    wsDept.Range("E21").Value = "Outpatient"
    wsDept.Range("E22").Value = "Total"
    wsDept.Range("E23").Value = "Inpatient"
    wsDept.Range("E24").Value = "Outpatient"
    wsDept.Range("E25").Value = "Total"
    
    ' Populate the cells in column D
    wsDept.Range("D6").Value = "Specialty"
    wsDept.Range("D7").Value = "Obstetrics"
    wsDept.Range("D13").Value = "Gynecology"
    wsDept.Range("D16").Value = "Pediatrics"
    wsDept.Range("D20").Value = "Polyclinic"
    wsDept.Range("D23").Value = "IVF"
    wsDept.Range("D26").Value = "Grand Total"
    ' Populate headers in other columns
    wsDept.Range("O5").Value = "Number of Cases"
    wsDept.Range("V5").Value = "Average Bill Size"
    wsDept.Range("H5").Value = "Revenue"
    wsDept.Range("I6").Value = "Target"
    wsDept.Range("P6").Value = "Target"
    wsDept.Range("L6").Value = "Budget Variance"
    wsDept.Range("S6").Value = "Budget Variance"
    wsDept.Range("Z6").Value = "Budget Variance"
    wsDept.Range("K6").Value = "Last Month Variance"
    wsDept.Range("R6").Value = "Last Month Variance"
    wsDept.Range("Y6").Value = "Last Month Variance"
    
End Sub

Sub HR_English()
    ' Ensure the workbook and worksheet are initialized
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' Clear the target range
    wsHR.Range("D8:D16").ClearContents
    ' Populate the cells with specific text
    wsHR.Range("D8").Value = "Frontline"
    wsHR.Range("D9").Value = "Physicians (Hospital-Employed)"
    wsHR.Range("D10").Value = "Nurses and Midwives"
    wsHR.Range("D11").Value = "Others"
    wsHR.Range("D12").Value = "Physicians from Agencies"
    wsHR.Range("D13").Value = "Support staff"
    wsHR.Range("D15").Value = "Support from other branches"
    wsHR.Range("D16").Value = "Total"
    wsHR.Range("E6").Value = "Total Salary (VND Bil)"
    wsHR.Range("K6").Value = "Number of Employees"
    wsHR.Range("P6").Value = "Average Salary (VND Mil)"
    wsHR.Range("U6").Value = "Average fixed salary per employee"
    wsHR.Range("AA6").Value = "Average variable salary per employee"
    wsHR.Range("J7").Value = "Last Month Variance"
    wsHR.Range("O7").Value = "Last Month Variance"
    wsHR.Range("T7").Value = "Last Month Variance"
    wsHR.Range("Y7").Value = "Last Month Variance"
    wsHR.Range("AD7").Value = "Last Month Variance"
    'Restore screen updating and calculation settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub HR_VN()
    ' Ensure the workbook and worksheet are initialized
    If wb Is Nothing Or wsDept Is Nothing Then Initialize
    ' Clear the target range
    wsHR.Range("D8:D16").ClearContents
    
    ' Populate the cells with specific text
    wsHR.Range("D8").Value = "Kh" & ChrW(7889) & "i chuyên môn"
    wsHR.Range("D9").Value = "Bác s" & ChrW(297) & " c" & ChrW(417) & " h" & _
        ChrW(7919) & "u"
    wsHR.Range("D10").Value = ChrW(272) & "i" & ChrW(7873) & "u d" & ChrW(432) _
        & ChrW(7905) & "ng/NHS"
    wsHR.Range("D11").Value = "Khác"
    wsHR.Range("D12").Value = "Bác s" & ChrW(297) & " h" & ChrW(417) & "p tác"
    wsHR.Range("D13").Value = "Nhân viên v" & ChrW(7853) & "n hành"
    wsHR.Range("D15").Value = "H" & ChrW(7895) & " tr" & ChrW(7907) & " t" & _
        ChrW(7915) & " chi nhánh khác"
    wsHR.Range("D16").Value = "T" & ChrW(7893) & "ng "
    wsHR.Range("E6").Value = "T" & ChrW(7893) & "ng qu" & ChrW(7929) & " l" & _
        ChrW(432) & ChrW(417) & "ng"
    wsHR.Range("K6").Value = "S" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) _
        & "ng nhân s" & ChrW(7921)
    wsHR.Range("P6").Value = "L" & ChrW(432) & ChrW(417) & "ng bình quân (tri" _
         & ChrW(7879) & "u " & ChrW(273) & ChrW(7891) & "ng)"
    wsHR.Range("U6").Value = "L" & ChrW(432) & ChrW(417) & "ng c" & ChrW(7889) _
         & " " & ChrW(273) & ChrW(7883) & "nh bình quân (tri" & ChrW(7879) & "u " _
         & ChrW(273) & ChrW(7891) & "ng)"
    wsHR.Range("AA6").Value = "L" & ChrW(432) & ChrW(417) & "ng s" & ChrW(7843) _
         & "n ph" & ChrW(7849) & "m bình quân (tri" & ChrW(7879) & "u " & ChrW( _
        273) & ChrW(7891) & "ng)"
    wsHR.Range("J7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsHR.Range("O7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsHR.Range("T7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsHR.Range("Y7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
'   wsHR.Range("AC7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    wsHR.Range("AD7").Value = "So v" & ChrW(7899) & "i tháng tr" & ChrW(432) & _
        ChrW(7899) & "c"
    
    
End Sub
Sub Chart_English()
' Ensure the workbook and worksheet are initialized
If wb Is Nothing Or wsDept Is Nothing Then Initialize
wsChart.Range("D23") = "Cost of Sales"
wsChart.Range("D24") = "SG&A"
wsChart.Range("D25") = "Employee's Benefits"
wsChart.Range("D26") = "EBITDA (Before Welfare)"
wsChart.Range("D28") = "EBITDA (After Welfare)"
End Sub

Sub Chart_VN()
' Ensure the workbook and worksheet are initialized
If wb Is Nothing Or wsDept Is Nothing Then Initialize
wsChart.Range("D23") = "Giá v" & ChrW(7889) & "n"
wsChart.Range("D24") = "Chi phí Qu" & ChrW(7843) & "n lý"
wsChart.Range("D25") = "Phúc l" & ChrW(7907) & "i"
wsChart.Range("D26") = "EBITDA (ch" & ChrW(432) & "a bao g" & ChrW(7891) _
         & "m phúc l" & ChrW(7907) & "i)"
wsChart.Range("D28") = "EBITDA (bao g" & ChrW(7891) & "m phúc l" & ChrW( _
        7907) & "i)"
End Sub

Sub RadarChart_English()
' Ensure the workbook and worksheet are initialized
If wb Is Nothing Or wsDept Is Nothing Then Initialize
wsChart.Range("D31") = "Cost of Sales"
wsChart.Range("D32") = "SG&A"
wsChart.Range("D33") = "Employee's Benefits"
wsChart.Range("D34") = "EBITDA (Before Welfare)"
wsChart.Range("D35") = "Revenue"
wsChart.Range("D36") = "EBITDA (After Welfare)"
End Sub

Sub RadarChart_VN()
' Ensure the workbook and worksheet are initialized
If wb Is Nothing Or wsDept Is Nothing Then Initialize
wsChart.Range("D31") = "Giá v" & ChrW(7889) & "n"
wsChart.Range("D32") = "Chi phí Qu" & ChrW(7843) & "n lý"
wsChart.Range("D33") = "Phúc l" & ChrW(7907) & "i"
wsChart.Range("D34") = "EBITDA (ch" & ChrW(432) & "a bao g" & ChrW(7891) _
         & "m phúc l" & ChrW(7907) & "i)"
wsChart.Range("D35") = "Doanh thu"
wsChart.Range("D36") = "EBITDA (bao g" & ChrW(7891) & "m phúc l" & ChrW( _
        7907) & "i)"
End Sub



