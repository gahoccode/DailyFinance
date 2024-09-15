Attribute VB_Name = "Module1"
Option Explicit
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare PtrSafe Function SetForegroundWindow Lib "user32" _
(ByVal hWnd As Long) As Long

Declare PtrSafe Function SetCursorPos Lib "user32" _
(ByVal x As Long, ByVal y As Long) As Long

Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, _
ByVal dwExtraInfo As Long)

Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Const MOUSEEVENTF_LEFTUP As Long = &H4
Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Const MOUSEEVENTF_RIGHTUP As Long = &H10
Sub SendMonthlyReport()

    Dim wbPL As Workbook
    Dim wsPL As Worksheet
    Dim wsHR As Worksheet
    Dim wsDept As Worksheet
    Dim tblPL As Range
    Dim tblHR As Range
    Dim tblDept As Range

    ' Set references to the workbook and sheets
    Set wbPL = Workbooks("Monthly P&L 2024_PCSG.xlsx")
    Set wsPL = wbPL.Sheets("PL Details")
    Set wsHR = wbPL.Sheets("HR")
    Set wsDept = wbPL.Sheets("Departments")

    ' Set references to the ranges
    Set tblPL = wsPL.Range("D1:P26")
    Set tblHR = wsHR.Range("D6:T16")
    Set tblDept = wsDept.Range("D5:Y26")

    ' Copy each range individually
    tblPL.Copy
    Call OpenZaloAndMoveMouse
'    Application.Wait Now + TimeValue("00:00:01")
    tblHR.Copy
'    Application.Wait Now + TimeValue("00:00:01")
    Call OpenZaloAndMoveMouse3
    tblDept.Copy
    Call OpenZaloAndMoveMouse2
    ' After copying, the content will be available on the clipboard
    ' and can be pasted wherever needed
    'MsgBox "Tables have been copied.", vbInformation

End Sub

Sub OpenZaloAndMoveMouse()
    ' Open Zalo
    Shell "C:\Users\tamml\AppData\Local\Programs\Zalo\Zalo.exe", vbNormalFocus
    
    ' Wait for Zalo to open (adjust sleep time as needed)
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Find Zalo window
    Dim ZaloHwnd As Long
    ZaloHwnd = FindWindow(vbNullString, "Zalo")
    
    ' Bring Zalo window to foreground
    SetForegroundWindow ZaloHwnd
    
    ' Move mouse cursor to X:283 Y:80 and click the left mouse button
    'Dua chuot vao o tim kiem ten nhom Zalo
    SetCursorPos 283, 80
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    ' Simulate typing "PCSG - "
    'Go ten nhom Zalo vao o tim kiem
    SendKeys "PCSG - Phòng Tài Chính "
    'SendKeys "vo "
    Application.Wait Now + TimeValue("00:00:01")
    ' Move mouse cursor to X:363 Y:380 and click the left mouse button
    'Chon nhom zalo
    SetCursorPos 309, 279
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    ' Move mouse cursor to X:949 Y:967 and right click the mouse
    'Bam chuot phai de paste
    SetCursorPos 949, 967
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    Application.Wait Now + TimeValue("00:00:01")
    'Click chon paste
    ' Move mouse cursor to X:1010 Y:940 and click the left mouse button
    SetCursorPos 1010, 940
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

    'Hit Send
    SetCursorPos 1234, 876
    Application.Wait Now + TimeValue("00:00:01")
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

    'Hit Send again in the dialouge box
'    SetCursorPos 1364, 793
'    Application.Wait Now + TimeValue("00:00:02")
'    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Sub OpenZaloAndMoveMouse2()
    ' Open Zalo
    Shell "C:\Users\tamml\AppData\Local\Programs\Zalo\Zalo.exe", vbNormalFocus
    
    ' Wait for Zalo to open (adjust sleep time as needed)
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Find Zalo window
    Dim ZaloHwnd As Long
    ZaloHwnd = FindWindow(vbNullString, "Zalo")
    
    ' Bring Zalo window to foreground
    SetForegroundWindow ZaloHwnd
    
    'Move mouse to pasting window and Right Click
    SetCursorPos 1234, 785
    Application.Wait Now + TimeValue("00:00:02")
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    'Move mouse to paste option and click paste
    Application.Wait Now + TimeValue("00:00:01")
    SetCursorPos 1287, 891
    Application.Wait Now + TimeValue("00:00:01")
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    Application.Wait Now + TimeValue("00:00:01")
    SetCursorPos 1227, 838
    Application.Wait Now + TimeValue("00:00:01")
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    
'    SetCursorPos 1287, 891
'    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
End Sub

Sub OpenZaloAndMoveMouse3()
    ' Open Zalo
    Shell "C:\Users\tamml\AppData\Local\Programs\Zalo\Zalo.exe", vbNormalFocus
    
    ' Wait for Zalo to open (adjust sleep time as needed)
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Find Zalo window
    Dim ZaloHwnd As Long
    ZaloHwnd = FindWindow(vbNullString, "Zalo")
    
    ' Bring Zalo window to foreground
    SetForegroundWindow ZaloHwnd
    
'    'Move mouse to pasting window and Right Click
    SetCursorPos 1234, 785
    Application.Wait Now + TimeValue("00:00:02")
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
'
'    'Move mouse to paste option and click paste
    Application.Wait Now + TimeValue("00:00:01")
    SetCursorPos 1287, 891
    Application.Wait Now + TimeValue("00:00:01")
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    'Move mouse to paste option and click paste
    Application.Wait Now + TimeValue("00:00:01")
    SetCursorPos 1229, 797
    Application.Wait Now + TimeValue("00:00:01")
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    
'    SetCursorPos 1287, 891
'    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
End Sub
