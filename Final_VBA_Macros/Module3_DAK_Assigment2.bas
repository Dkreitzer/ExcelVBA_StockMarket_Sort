Attribute VB_Name = "Module3"

'SORT AND TOTALS

Sub SortAndTotals()
Dim ws As Worksheet
For Each ws In Worksheets

    'count the rows
    Dim RowNumber As Long
    RowNumber = ws.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    'Defining Variables
    Dim Brand_Name As String
    
    Dim Brand_Total As Double
    Brand_Total = 0
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'Declare and set a value for Opening Price
    Dim OpenPrice As Double
    OpenPrice = ws.Range("C2").Value
    ws.Cells(2, 20).Value = OpenPrice
    
    'Declare and set a value for Closing Price
    Dim ClosePrice As Double
    ClosePrice = 0


            'Compiling Loop Begins
            For i = 2 To RowNumber
            
                
                'check if we are in same card range, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                    'Change values as follows:
                    Brand_Name = ws.Cells(i, 1).Value                            'Brand Name
                    Brand_Total = Brand_Total + ws.Cells(i, 7).Value             'Brand Total
                    ClosePrice = ws.Cells(i, 6).Value                          'close of current ticker
                    OpenPrice = ws.Cells(i + 1, 3).Value                       'Open Price for next ticker
                    
                    'Print Brand / Total Values
                    ws.Range("I" & Summary_Table_Row).Value = Brand_Name
                    ws.Range("L" & Summary_Table_Row).Value = Brand_Total
                    ws.Range("U" & Summary_Table_Row).Value = ClosePrice
                                        
                    'Add a row to the Summary_Table_Row
                    Summary_Table_Row = Summary_Table_Row + 1
                                       
                    'Set Open Price for next comany
                    ws.Range("T" & Summary_Table_Row).Value = OpenPrice
                    
                    'Reset counter to 0
                    Brand_Total = 0
                    
            'if it is the same brand card then
            Else
    
            'Add the existing Brand_Total to the next total in (i,3)
            Brand_Total = Brand_Total + ws.Cells(i, 7).Value
        
        End If

    Next i
   
    Next ws
    

End Sub

