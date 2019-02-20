Attribute VB_Name = "Module7"
Sub ConditionalFormatting()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

    Dim PercentChange As Double
    
    Dim LastRow As Long
    LastRow = ws.Range("K" & Rows.Count).End(xlUp).Row

    For i = 2 To LastRow
    With ws.Cells(i, 11)
        
            If .Value < 0 Then                              'If the cell's value is less than 0, then turn Red
               ws.Cells(i, 11).Interior.ColorIndex = 3      'Turn the cell color Red
               
            Else                                            'Else, Turn Green
                ws.Cells(i, 11).Interior.ColorIndex = 4     'Turn the cell color Green
                
             End If
           End With
        Next i

Next ws


End Sub

