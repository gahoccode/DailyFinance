Attribute VB_Name = "Module3"
Sub MovePicture5()
    Dim pic As Shape
    
    ' Set the reference to Picture 5
    Set pic = ActiveSheet.Shapes("Picture 5")
    
    ' Set the position and size
    With pic
        .Top = 0
        .Left = 91.5584259
        .Width = 446.8104858
        .Height = 172.2077942
    End With
End Sub
Sub GetSelectedImagePosition()
    Dim selectedShape As Shape
    Dim newWorkbook As Workbook
    Dim ws As Worksheet
    
    ' Check if any shape is selected
    If TypeName(Selection) <> "Picture" And TypeName(Selection) <> "Shape" Then
        MsgBox "Please select a picture or shape first.", vbExclamation
        Exit Sub
    End If
    
    ' Set the selected shape
    Set selectedShape = Selection.ShapeRange(1)
    
    ' Open a new workbook
    Set newWorkbook = Workbooks.Add
    Set ws = newWorkbook.Sheets(1)
    
    ' Write the position data in the new workbook
    ws.Range("A1").Value = "Selected Shape Position"
    ws.Range("A2").Value = "Top"
    ws.Range("B2").Value = selectedShape.Top
    ws.Range("A3").Value = "Left"
    ws.Range("B3").Value = selectedShape.Left
    ws.Range("A4").Value = "Width"
    ws.Range("B4").Value = selectedShape.Width
    ws.Range("A5").Value = "Height"
    ws.Range("B5").Value = selectedShape.Height
    
    ' Notify the user
    MsgBox "The position of the selected shape has been recorded in a new workbook.", vbInformation
End Sub


Sub ToggleOffDates()
'
' ToggleOffDates Macro


    Range("B1:C5").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub
Sub ToggleOnDates()

' ToggleOnDates Macro


    Range("B1:C5").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
End Sub
