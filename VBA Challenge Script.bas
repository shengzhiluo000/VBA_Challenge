Attribute VB_Name = "Module11"
Sub ticker_Alpha()


'Declaration for step 1

Dim Ticker As String

Dim YearlyChange As Double

Dim PercentChange As Double

Dim TotalStockVolume As Double

Dim TickerRow As Integer

'Counter of how many iteration for same ticker... used to store open value

Dim OpenCount As Integer

'Decalaration for step 2

Dim GreatestTicker As String

Dim GreatestPercentChange As Double

Dim LeastTicker As String

Dim LeastPercentChange As Double

Dim HighestTicker As String

Dim HighestVolume As Double

'worksheet loop

For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    TotalStockVolume = 0
    
    TickerRow = 2
    
    OpenCount = 0
    
    'name title for each column for step 1
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'name title for each column for step 2
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    'Step 1: find yearly change, percent change and total volumn for each distinct ticker
    
    For i = 2 To LastRow
    
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Ticker Column
            
            Ticker = ws.Cells(i, 1).Value
            
            ws.Range("I" & TickerRow).Value = Ticker
            
            'Total Stock Volume Column
            
            TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
            
            ws.Range("L" & TickerRow).Value = TotalStockVolume
            
            'Close value = ws.Range("F" & i).Value
            
            'Open value of the year = ws.Range("C" & (i - OpenCount)
            
            'Year Change and Percent Change Column
            
            YearlyChange = ws.Range("F" & i).Value - ws.Range("C" & (i - OpenCount)).Value
            
            ws.Range("J" & TickerRow).Value = YearlyChange
            
            PercentChange = YearlyChange / ws.Range("C" & (i - OpenCount)).Value
            
            ws.Range("K" & TickerRow).Value = FormatPercent(PercentChange)
            
            'Recount to start a different ticker
            
            TotalStockVolume = 0
            
            TickerRow = TickerRow + 1
            
            OpenCount = 0
            
        Else
        
            TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
            
            'increase count if ticker is the same for open value
            
            OpenCount = OpenCount + 1
        
        End If
    
    Next i
    
    'Step 2: find the greatest and least percent change, and greatest total volumn and their respective ticker
    
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    GreatestTicker = ws.Range("I2").Value
    GreatestPercentChange = ws.Range("K2").Value
    LeastTicker = ws.Range("I2").Value
    LeastPercentChange = ws.Range("K2").Value
    HighestTicker = ws.Range("I2").Value
    HighestVolumn = ws.Range("L2").Value
    
    'find the greatest and least percent change
    
    For i = 2 To LastRow2
    
        If ws.Range("K" & i).Value > GreatestPercentChange Then
        
            GreatestTicker = ws.Range("I" & i).Value
            GreatestPercentChange = ws.Range("K" & i).Value
        
        ElseIf ws.Range("K" & i).Value < LeastPercentChange Then
        
            LeastTicker = ws.Range("I" & i).Value
            LeastPercentChange = ws.Range("K" & i).Value
        
        Else
        
            GreatestTicker = GreatestTicker
            GreatestPercentChange = GreatestPercentChange
            LeastTicker = LeastTicker
            LeastPercentChange = LeastPercentChange
        
        End If
    
    Next i
    
    ' find the greatest total volumn
    
    ws.Range("P2").Value = GreatestTicker
    ws.Range("Q2").Value = FormatPercent(GreatestPercentChange)
    ws.Range("P3").Value = LeastTicker
    ws.Range("Q3").Value = FormatPercent(LeastPercentChange)
    
    For i = 2 To LastRow2
    
        If ws.Range("L" & i).Value > HighestVolumn Then
            
            HighestVolumn = ws.Range("L" & i).Value
            HighestTicker = ws.Range("I" & i).Value
            
        Else
        
            HighestTicker = HighestTicker
            HighestVolumn = HighestVolumn
    
        End If
    
    Next i
    
    ws.Range("P4").Value = HighestTicker
    ws.Range("Q4").Value = HighestVolumn
    
    'Conditional formatting for Yearly change and percent change
    
    For i = 2 To LastRow2
    
        'Green if value >= 0
        
        If ws.Range("J" & i).Value >= 0 Then
        
            ws.Range("J" & i).Interior.ColorIndex = 4
            ws.Range("K" & i).Interior.ColorIndex = 4
        
        'Red if value < 0
        
        Else:
        
            ws.Range("J" & i).Interior.ColorIndex = 3
            ws.Range("K" & i).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    'Auto format Column width
    
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub
