Attribute VB_Name = "Module2"
'Run Stock Market Master Page


Sub StockMarket()

'PURPOSE: Determine how many seconds it took for code to completely run
'Timer Starts Now!

Dim StartTime As Double
Dim SecondsElapsed As Double
'Remember time when macro starts
  StartTime = Timer

'SORT AND Totals
    Call SortAndTotals          'in Module3
    Call YearPerformance        'in Module4
    Call FindMaxVol             'in Module5
    Call FindMaxPer             'in Module5
    Call LowPer                 'in Module5
    Call Formatting             'in Module6
    Call ConditionalFormatting  'in Module7

'END TIMER
'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation


End Sub





