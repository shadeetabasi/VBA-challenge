Sub ConditionalFormatting()

'Loop Through All Sheets
For Each ws In Worksheets

'Declare variable
Dim i As Long

'Define last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop Until Last Row
    For i = 2 To LastRow

If ws.Cells(i, 10) > 0 Then
ws.Cells(i, 10.Interior.ColorIndex = 4

Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

Next ws

End Sub

