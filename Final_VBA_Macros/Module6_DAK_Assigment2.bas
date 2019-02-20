Attribute VB_Name = "Module6"

'Formatting
Sub Formatting()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

                        'Headers Text
                        ws.Cells(1, 9).Value = "Ticker"
                        ws.Cells(1, 10).Value = "Yearly Change"
                        ws.Cells(1, 11).Value = "Percent Change"
                        ws.Cells(1, 12).Value = "Total Stock Volume"
                        ws.Cells(1, 20).Value = "Open"
                        ws.Cells(1, 21).Value = "Close"
                        
                        
                        'Summary Text
                        ws.Cells(2, 15).Value = "Greatest % Increase"
                        ws.Cells(3, 15).Value = "Greatest % Decrease"
                        ws.Cells(4, 15).Value = "Greatest Total Volume"
                        ws.Cells(1, 16).Value = "Ticker"
                        ws.Cells(1, 17).Value = "Value"
                        
                        'Make Text Bold
                        ws.Range("O2").EntireColumn.Font.Bold = True
                        ws.Range("P1:Q1").Font.Bold = True
                        
                        'Font for Price Change
                        ws.Range("J2").EntireColumn.NumberFormat = "0.000000"
                        
                        'Font for Total Stock Volume
                        ws.Range("L2").EntireColumn.NumberFormat = "0"
                        
                        'Autofit Columns
                        ws.Range("I1:L1").EntireColumn.AutoFit
                        ws.Range("O1:Q1").EntireColumn.AutoFit
                        
                        'Make Number %
                        ws.Range("K2").EntireColumn.NumberFormat = "0.00%"
                        ws.Cells(2, 17).NumberFormat = "0.00%"
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                           
    'next worksheet
    Next ws
    
End Sub




