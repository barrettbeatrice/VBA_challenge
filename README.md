# VBA_challenge
Challenge 2 VBA for Bootcamp Code Solution


'Looping across multiple worksheets for all years 2018, 2019, 2020

Sub LoopStockAnalysis()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Select
        Call StockAnalysis
    Next
    
End Sub

Sub StockAnalysis()

'Create a script that loops through all the stocks and output on each worksheet


'Use conditional formatting in the yearly change (Red if <0, Green if>0)

'Create Column Titles
Cells(1, 9) = "Ticker_Symbol"
Cells(1, 10) = "Yearly_Change"
Cells(1, 11) = "Percent_Change"
Cells(1, 12) = "Total_Stock_Volume"

'Bold the column titles
Cells(1, 9).Font.Bold = True
Cells(1, 10).Font.Bold = True
Cells(1, 11).Font.Bold = True
Cells(1, 12).Font.Bold = True


'step 1:Declare the variables and give them initialization values
'Assign integer for the loop to start
    
Dim Ticker As Integer
Ticker = 1

Dim SummaryTableRow As Long
SummaryTableRow = 2

Dim TotalVolume  As Double
TotalVolume = 0

Dim OpeningValue As Double

Dim ClosingValue As Double

Dim ClosingMinusOpening As Double


' Step 2: Determine the total length/number of worksheet rows
rowmax = Cells(Rows.Count, "A").End(xlUp).Row

'Step 3: Loop through the rows in the column
    For I = 2 To rowmax


'Step 4: identify the cumulative volumes for each group of Tickers
TotalVolume = TotalVolume + Cells(I, 7).Value
           
'For when the ticker changes to another stock
    If Cells(I - 1, Ticker).Value <> Cells(I, Ticker).Value Then
    OpeningValue = Cells(I, 3).Value
    
    End If

'Step 5: Search for when the value of the next row's cell is different than the previous cells value
    'identifies when the ticker code changes). Use IF-Then statement
'Step 6: Output the ticker value into column "I"
    
    If Cells(I + 1, Ticker).Value <> Cells(I, Ticker).Value Then
    'set closing value variable
    ClosingValue = Cells(I, 6).Value
    
    ClosingMinusOpening = ClosingValue - OpeningValue
    
 
    Cells(SummaryTableRow, 9).Value = Cells(I, Ticker).Value
    
    Cells(SummaryTableRow, 10).Value = ClosingMinusOpening
    
    
'Step 7: Determine the Percent Change
'Divide the value in "Step 1" by the value in "Step 2" and multiply by 100
   
    Cells(SummaryTableRow, 11).Value = ClosingMinusOpening / OpeningValue
    
'Step 8: format  'Percent Change' column 11 (K) format to number format and a %
    Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
    
'Now calculate the TotalVolume for Column 12(L)
    Cells(SummaryTableRow, 12).Value = TotalVolume
    
    SummaryTableRow = SummaryTableRow + 1

    'need to reset between tickers
    TotalVolume = 0

    End If
    
Next I
    
'Use conditional formatting for Percent Change. IF value > 0, green. If value <0 = red
'"When" loop statements.....then format it like this

   Columns("K:K").Select
   
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  
    Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
    End With
    Selection.FormatConditions(1).StopIfTrue = False
  
'Remove formatting created above from  "K1 & J1" (the column titles)
Range("K1").Select
    Selection.FormatConditions.Delete
    
Range("J1").Select
    Selection.FormatConditions.Delete
    
'BONUS/Advanced Part *************************************************

'adding functionality to display the stock with the greatest % increase,greatest % decrease
' - and the greatest total volume

'create column and row titles
Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"
Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"

'Extend the length of column 14 to fit the titles
Columns(14).AutoFit

'Bold the column titles
Cells(1, 15).Font.Bold = True
Cells(1, 16).Font.Bold = True
Cells(2, 14).Font.Bold = True
Cells(3, 14).Font.Bold = True
Cells(4, 14).Font.Bold = True



'Declare variables that will be used

Dim MaxValue As Double
Dim MinValue As Double
Dim GreatestTotalVolume As Double

'Find max value within column k
    MaxValue = Application.WorksheetFunction.Max(Range("K:k"))

'display max value found in P2, convert value to percentage
        Cells(2, 16) = MaxValue
        Cells(2, 16).NumberFormat = "0.00%"

'Find the min Value of column k
    MinValue = Application.WorksheetFunction.Min(Range("K:k"))

'display min value in P3, converet value to percentage
        Cells(3, 16) = MinValue
        Cells(3, 16).NumberFormat = "0.00%"

'Find the greatest total volume in column L
    GreatestTotalVolume = Application.WorksheetFunction.Max(Range("L:l"))

'display the greatest toal volume in P4
        Cells(4, 16) = GreatestTotalVolume

'find the location of the ticker that corresponds to the above values

'declare variables that will be used
Dim inc_loc As Integer
Dim dec_loc As Integer
Dim totalvolloc As Integer

'use match function, to find value corresponding to MaxValue, Min Value, and GreatestTotalVolume
inc_loc = WorksheetFunction.Match(MaxValue, Range("K:K"), 0)
dec_loc = WorksheetFunction.Match(MinValue, Range("K:K"), 0)
totalvolloc = WorksheetFunction.Match(GreatestTotalVolume, Range("L:L"), 0)

' assign them to the table cells (add 1 beacuse range above didnt include the header row)
Range("O2") = Cells(inc_loc + 1, 9)
Range("O3") = Cells(dec_loc + 1, 9)
Range("O4") = Cells(totalvolloc + 1, 9)


End Sub

