Attribute VB_Name = "Module5"
'This code finds the highest volume in a range and returns additional information
Sub FindMaxVol()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    Dim i As Long
    Dim firstRow As Integer
    Dim columnNumber As Integer
    Dim max As Double
    Dim tag As String
    
    firstRow = 2
    columnNumber = 12
    'Count how many rows in Column
    Dim LastRow As Long
    LastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
        If ws.UsedRange.Rows.Count <= 1 Then max = 0 Else max = ws.Cells(2, 12)
        
        For i = firstRow To LastRow
           With ws.Cells(i, 12)
        
             If .Value > max Then
               max = .Value
               tag = .Offset(0, -3).Value
             End If
           End With
        Next i
        
        ws.Cells(4, 17) = max
        ws.Cells(4, 16).Value = tag
    Next ws
    
End Sub

'This code finds the highest Percent in a range and returns additional information
Sub FindMaxPer()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    Dim i As Long
    Dim firstRow As Integer
    Dim columnNumber As Integer
    Dim max As Double
    Dim tag As String
    
    firstRow = 2
    columnNumber = 11
    'Count how many rows in Column
    Dim LastRow As Long
    LastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
        If ws.UsedRange.Rows.Count <= 1 Then max = 0 Else max = ws.Cells(2, 11)
        
        For i = firstRow To LastRow
           With ws.Cells(i, 11)
        
             If .Value > max Then
               max = .Value
               tag = .Offset(0, -2).Value
             End If
           End With
        Next i
        
        ws.Cells(2, 17) = max
        ws.Cells(2, 16).Value = tag
    Next ws
    
End Sub

'This code finds the lowest Percent in a range and returns additional information
Sub LowPer()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

    Dim i As Long
    Dim firstRow As Integer
    Dim columnNumber As Integer
    Dim min As Double
    Dim tag As String
    
    firstRow = 2
    columnNumber = 11
    'Count how many rows in Column
    Dim LastRow As Long
    LastRow = ws.Range("I" & Rows.Count).End(xlUp).Row
    
        If ws.UsedRange.Rows.Count <= 1 Then min = 0 Else min = ws.Cells(2, 11)
        
        For i = firstRow To LastRow
           With ws.Cells(i, 11)
        
             If .Value < min Then
               min = .Value
               tag = .Offset(0, -2).Value
             End If
           End With
        Next i
        
        ws.Cells(3, 17) = min
        ws.Cells(3, 16).Value = tag
    Next ws
    
End Sub


