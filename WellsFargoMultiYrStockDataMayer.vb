Sub WellsFargoMultiYrStockDataMayer()

    Dim WS_Count As Integer
    Dim I As Integer

    'create loop for all worksheets in the active workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    For I = 1 To WS_Count
    
        'ActiveWorkbook.Worksheets(I).Activate
    
        'declare variables for ticker and volume - "easy" segments
        Dim ws As Worksheets
        Dim Ticker As String
        Dim TotalVol As Long
        Dim lastrow As Long
        
        'declare variables for yearly change & percent change - "moderate" segments
        Dim TickerOpen As Integer
        Dim TickerClose As Integer
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        'declare variables for 'Greatest' chart - 'hard' segment
        Dim SClastrow As Long
        Dim MaxPercent As Double
        Dim MinPercent As Double
        Dim GreatestIncreaseTicker As String
        Dim GreatestDecreaseTicker As String
        
        
        'this prevents my overflow error
        On Error Resume Next
    
        'label all worksheets with all Titles for new Summary Chart
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'label Last Set of requested data - 'Hard' segment
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        
        'variables and starter for pulling all new data into two new summary charts
        Dim SummaryChartRow As Integer
        Dim BonusSummaryChartRow As Integer
        SummaryChartRow = 2
        BonusSummaryChartRow = 2
        
        'define lastrow
        lastrow = Cells(Rows.Count, 1).End(xlUp).row

        'create integer for loop -------------------------------------------------------------------------------
        For row = 2 To lastrow
        
            'check for Jan 1 date)
            If Cells(row, 2).Value = (ActiveSheet.Name * 10000) + 101 Then
                
                'grab jan 01 open ticker price
                OpenPrice = Cells(row, 3).Value
            
            Else
            
                'locate when new ticker name appears
                If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
                
                    'grab ticker name for Summary Chart
                    Ticker = Cells(row, 1).Value
            
                    'add to Total Volume
                    TotalVol = TotalVol + Cells(row, 7).Value
        
                    'print Ticker in Summary Chart
                    Range("I" & SummaryChartRow) = Ticker
         
                    'print TotalVolume in Summary Chart
                    Range("L" & SummaryChartRow) = TotalVol
            
                    'grab Dec 31 close ticker price
                    ClosePrice = Cells(row, 6).Value
            
                    'calculate difference between opening and closing price
                    YearlyChange = ClosePrice - OpenPrice
    
                    'print yearly change in Summary Chart
                    Range("J" & SummaryChartRow) = YearlyChange
                    
                    'calculate tickers percent change
                    PercentChange = (ClosePrice - OpenPrice) / OpenPrice
         
                    'print percent change in Summary Chart
                    Range("K" & SummaryChartRow) = PercentChange
                    
                    'format percent change column correctly
                    Range("K:K").NumberFormat = "0.00%"
                       
                    'condition percent change cell in Summary Chart to....
                    If Cells(SummaryChartRow, 10).Value < 0 Then
            
                        'Red if negative
                        Cells(SummaryChartRow, 10).Interior.ColorIndex = 3
        
                    Else
            
                        'Black if positive
                        Cells(SummaryChartRow, 10).Interior.ColorIndex = 4
                
                    End If
            
                    'Reset the Total Volume counter
                    TotalVol = 0
        
                    'add one to Summary Chart counter to continue in new line in chart
                    SummaryChartRow = SummaryChartRow + 1
                
                'if the ticker name is same
                Else
        
                    'keep accumulating Total Volume  - Tried long,int,double
                    TotalVol = TotalVol + Cells(row, 7).Value
            
                End If
                
            End If
            
        Next row

            'LOCATE GREATEST INCREASE
        
            'format min & max cells in BonusChart
            Cells(2, 16).NumberFormat = "0.00%"
            Cells(3, 16).NumberFormat = "0.00%"
            
            'define lastrow in SummaryChart
            SClastrow = Cells(Rows.Count, 9).End(xlUp).row
        
            'set max
            MaxPercent = 0
            
            'reset SummaryChartRow
            SummaryChartRow = 2
    
            'begin loop to find Greatest Increase
            For SummaryChartRow = 2 To SClastrow
    
                'find values larger than latest max and setting new max
                If Cells(SummaryChartRow, 11).Value > MaxPercent Then
            
                    'new Greatest Increase
                    MaxPercent = Cells(SummaryChartRow, 11).Value
                    
                    'locate Greatest Increase Ticker
                    GreatestIncreaseTicker = Cells(SummaryChartRow, 9).Value
                    
                End If
            
            Next SummaryChartRow
        
            'print Greatest Increase Ticker
            Cells(BonusSummaryChartRow, 15).Value = GreatestIncreaseTicker
    
            'Print Greatest Increase
            Cells(BonusSummaryChartRow, 16).Value = MaxPercent
    
            'add one to Bonus Summary Chart counter to continue in new line in bonus chart
            BonusSummaryChartRow = BonusSummaryChartRow + 1
            
            'LOCATE GREATEST DECREASE
            
            'set min
            MinPercent = 0
            
            'reset SummaryRowChart
            SummaryRowChart = 2
            
            'begin loop to find Greatest Decrease
            For SummaryChartRow = 2 To SClastrow
    
                'find values larger than latest max and setting new max
                If Cells(SummaryChartRow, 11).Value < MinPercent Then
            
                    'new Greatest Decrease
                    MinPercent = Cells(SummaryChartRow, 11).Value
            
                    'locate Greatest Decrease Ticker
                    GreatestDecreaseTicker = Cells(SummaryChartRow, 9).Value
                    
                End If
    
            Next SummaryChartRow
        
            'print Greatest Decrease Ticker
            Cells(BonusSummaryChartRow, 15).Value = GreatestDecreaseTicker
    
            'Print Greatest Decrease
            Cells(BonusSummaryChartRow, 16).Value = MinPercent
    
            'add one to Bonus Summary Chart counter to continue in new line in bonus chart
            BonusSummaryChartRow = BonusSummaryChartRow + 1
    
            'LOCATE GREATEST TOTAL VOLUME
            
            'set max
            MaxVol = 0
    
            'reset SummaryRowChart
            SummaryRowChart = 2
            
            'begin loop to find Greatest Total Volume
            For SummaryChartRow = 2 To SClastrow
    
                'find values larger than latest max and setting new max
                If Cells(SummaryChartRow, 12).Value > MaxVol Then
            
                    'new Greatest Total Volume
                    MaxVol = Cells(SummaryChartRow, 12).Value
        
                    'locate Greatest Total Volume Ticker
                    GreatestTotalVolTicker = Cells(SummaryChartRow, 9).Value
        
                End If
    
            Next SummaryChartRow
    
            'print Greatest Total Volume Ticker
            Cells(BonusSummaryChartRow, 15).Value = GreatestTotalVolTicker
    
            'Print Greatest Increase
            Cells(BonusSummaryChartRow, 16).Value = MaxVol
           
    Next I
    
End Sub


