Attribute VB_Name = "Module4"
'YearPerformance
Sub YearPerformance()

Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
    
        'Count how many rows in Column
        Dim LastRow As Long
        LastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
            
            'Begin Loop to determine amount change and percent change from open to close
            For i = 2 To LastRow
            
                If ws.Cells(i, 20).Value <> 0 Then
                
            
                    'Determine change in value from open to close
                    ws.Cells(i, 10).Value = (ws.Cells(i, 21).Value - ws.Cells(i, 20).Value)
        
                    'Determine percent increase or decrease
                    ws.Cells(i, 11).Value = (ws.Cells(i, 10).Value / ws.Cells(i, 20).Value)
            Else
            End If
            
            'Next Row
            Next i
    Next ws

End Sub


