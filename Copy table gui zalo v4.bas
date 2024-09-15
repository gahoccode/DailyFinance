Attribute VB_Name = "Module2"
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
'Sub Initialize()
'    ' Call the subroutine to copy the table and pictures
'    Call CopyTablePic
'
'    ' Call the subroutine to open Zalo and move the mouse
'    Call OpenZaloAndMoveMouse
'End Sub

Sub CopyTablePic()
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Define the source ranges and workbooks
    Dim rangeKCB As Range, rangeRevenue As Range
    Dim wbKCB As Workbook, wbRevenue As Workbook
    Dim wbKCBExists As Boolean, wbRevenueExists As Boolean
    Dim picIndex As Integer
    Dim picsPasted As Integer
    Dim pasteTop As Double, pasteLeft As Double
    Dim lastRow As Long
    Dim firstPic As Object, secondPic As Object ' Variables to hold the pictures for copying

    On Error Resume Next
    ' Attempt to set the workbooks
    Set wbKCB = Workbooks("So lieu KCB_Final.xlsx")
    wbKCBExists = Not wbKCB Is Nothing
    Set wbRevenue = Workbooks("Daily Revenue 2024.xlsx")
    wbRevenueExists = Not wbRevenue Is Nothing
    On Error GoTo 0

    ' Add a new workbook
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    Application.Wait Now + TimeValue("00:00:05")

    ' Initialize picIndex, picsPasted, and initial paste coordinates
    picIndex = 1
    picsPasted = 0
    pasteTop = Range("A1").Top
    pasteLeft = Range("A1").Left

    ' Copy and paste the first source range as a linked picture if it exists
    If wbKCBExists Then
        On Error Resume Next
        With wbKCB.Sheets("Management Report")
            lastRow = .Cells(.Rows.Count, "H").End(xlUp).Row
            Set rangeKCB = .Range("B1:H" & lastRow)
        End With
        If Not rangeKCB Is Nothing Then
            wbKCB.Sheets("Management Report").Activate
            rangeKCB.Copy
            newWorkbook.Sheets(1).Activate
            ActiveSheet.Pictures.Paste(Link:=True).Select
            Selection.Top = pasteTop
            Selection.Left = pasteLeft
            Selection.Name = "Pic" & picIndex
            picIndex = picIndex + 1
            picsPasted = picsPasted + 1
            pasteTop = Selection.Top + Selection.Height + 10 ' Adjust 10 for spacing
            DoEvents
            Application.Wait Now + TimeValue("00:00:03")
            If picIndex = 2 Then ' Store the first picture for later copying
                Set firstPic = Selection
            End If
        End If
        On Error GoTo 0
    End If

    ' Resize Pictures
    On Error Resume Next
    Dim pic As Picture
    For Each pic In newWorkbook.Sheets(1).Pictures
        With pic.ShapeRange
            .LockAspectRatio = msoTrue ' Lock aspect ratio
            .Width = 10.68 * 72 ' Convert inches to points (1 inch = 72 points)
        End With
    Next pic

    ' Copy the first picture after resizing and paste it to Zalo
    If Not firstPic Is Nothing Then
        firstPic.Copy
        Call OpenZaloAndMoveMouse
    End If

    ' Copy and paste the second source range as a linked picture if it exists
    If wbRevenueExists Then
        On Error Resume Next
        With wbRevenue.Sheets("Management Report")
            lastRow = .Cells(.Rows.Count, "I").End(xlUp).Row
            Set rangeRevenue = .Range("B1:I" & lastRow)
        End With
        If Not rangeRevenue Is Nothing Then
            wbRevenue.Sheets("Management Report").Activate
            rangeRevenue.Copy
            newWorkbook.Sheets(1).Activate
            ActiveSheet.Pictures.Paste(Link:=True).Select
            Selection.Top = pasteTop
            Selection.Left = pasteLeft
            Selection.Name = "Pic" & picIndex
            picIndex = picIndex + 1
            picsPasted = picsPasted + 1
            pasteTop = Selection.Top + Selection.Height + 10 ' Adjust 10 for spacing
            DoEvents
            Application.Wait Now + TimeValue("00:00:03")
            If picIndex = 3 Then ' Store the second picture for later copying
                Set secondPic = Selection
            End If
        End If
        On Error GoTo 0
    End If

    ' Copy the second picture and paste it to Zalo
    If Not secondPic Is Nothing Then
        secondPic.Copy
        Call OpenZaloAndMoveMouse2
    End If

    ' Turn on screen updating
    Application.ScreenUpdating = True

    ' Notify the user about the number of pictures pasted
    'MsgBox picsPasted & " pictures were successfully pasted in the new workbook and are now linked. Please paste them where needed.", vbInformation
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
'    SetCursorPos 1234, 876
'    Application.Wait Now + TimeValue("00:00:01")
'    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

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
    
    ' Move mouse cursor to X:949 Y:967 and right click the mouse
    'Bam chuot phai de paste
    SetCursorPos 1071, 790
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    
    
    Application.Wait Now + TimeValue("00:00:01")
    SetCursorPos 1143, 888
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    'Final Send button
'    Application.Wait Now + TimeValue("00:00:01")
'    SetCursorPos 1356, 793
'    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    
    
End Sub




