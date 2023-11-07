Sub Stock_initiative()

'Creating Variables
Dim StockStartDate As Double
Dim StockEndDate As Double
Dim YearlyStockChange As Double
Dim TotalVolume As Double

Dim PercentChange As Double
Dim StockTicker As String
Dim count As Double
Dim GPI As Double
Dim GPD As Double
Dim GTV As Double

'creating range
Dim Rng As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition
    
'loop through all sheets

For Each ws In Worksheets


count = 2
TotalVolume = 0
GPI = 0
GPD = 0
GTV = 0

'last rownumber
LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row

'Labeling columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

'For Loop to find each individual stock ticker

    For i = 2 To LastRow
               
    'Finding the Yearly Change for stocks
    
        If ws.Cells(i, 2).Value Like "*0102" Then
            StockTicker = ws.Cells(i, 1).Value
            StockStartDate = ws.Cells(i, 3).Value
        End If
        
        If ws.Cells(i, 2).Value Like "*1231" Then
            StockEndDate = ws.Cells(i, 6).Value
            YearlyStockChange = StockEndDate - StockStartDate
        End If
    
    'Finding Percent Change
        
        If ws.Cells(i, 2).Value Like "*1231" Then
            PercentChange = (StockEndDate - StockStartDate) / StockStartDate
        End If
    
    'Calculating Total Volume of Stock Ticker
    
         If StockTicker = ws.Cells(count, 9) Then
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
         End If
    
    'Sorting the Different Stock Ticker, Yearly Change, Percent Change, and Total Volume
    
        If ws.Cells(i, 1).Value = StockTicker And ws.Cells(i, 2).Value Like "*1231" Then
            ws.Cells(count, 9).Value = StockTicker
            ws.Cells(count, 10) = YearlyStockChange
            ws.Cells(count, 11) = PercentChange
            ws.Cells(count, 12) = TotalVolume
    'Using "count" variable to neatly add rows for each individual stock ticker
        count = count + 1
        
        End If
    
    'Reseting Total Volume
    
        If ws.Cells(i, 1).Value = StockTicker And ws.Cells(i, 2).Value Like "*1231" Then
            TotalVolume = 0
        End If
    
    'Finding Greatest Percent Increase
    
        If PercentChange > GPI Then
            GPI = PercentChange
            ws.Range("O2").Value = StockTicker
        End If
    
    'Finding Greatest Percent Decrease
    
        If PercentChange < GPD Then
            GPD = PercentChange
            ws.Range("O3").Value = StockTicker
        End If
    
    'Finding Greatest Total Volume of Individual Stock
            
        If GTV < ws.Cells(count, 12).Value Then
            GTV = ws.Cells(count, 12).Value
            ws.Range("O4").Value = ws.Cells(count, 9).Value
        End If
        
    'Placing Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume
    
    
    ws.Range("P2").Value = GPI
    ws.Range("P3").Value = GPD
    ws.Range("P4").Value = GTV
        
    
        
    'Iterating next i
    Next i
    
    'Formatting to percentage
    ws.Range(ws.Cells(2, 11), ws.Cells(LastRow, 11)).NumberFormat = "0.00%"
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    
    
    'setting the range of the conditional formatting
    Set Rng = ws.Range(ws.Cells(2, 10), ws.Cells(LastRow, 11))
    
    'Clear any conditional formatting
    Rng.FormatConditions.Delete
    
    'Setting the criteria for the range
    Set condition1 = Rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set condition2 = Rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    'Setting the color for the conditions
    With condition1
    .Interior.ColorIndex = 4
    End With
   
    With condition2
    .Interior.ColorIndex = 3
    End With
    
 
      Next ws
      
    
      
End Sub
