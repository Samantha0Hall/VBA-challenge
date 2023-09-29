Attribute VB_Name = "Module1"
Sub StockMarketLoop()

    For Each ws In Worksheets
    
    
    'variables
        Dim Ticker As String
        
        Dim OpeningPrice As Double
        
        Dim ClosingPrice As Double
        
        Dim YearlyChange As Double
        
        Dim PercentChange As Double
        
        Dim TotalStockVolume As Double
        
        Dim LastRow As Long
        
        Dim SummaryTableRow As Integer
        
        Dim GreatestIncrease As Long
        Dim GreatestDecrease As Long
        Dim GreatestTotalVolume As Long
        
        
        
        
        SummaryTableRow = 2
        
      
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
        
    'summary table headers
            ws.Cells(1, 9).Value = "Ticker"
            
            ws.Cells(1, 10).Value = "Yearly Change"
            
            ws.Cells(1, 11).Value = "Percent Change"
            
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
     'greatest summary table headers
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"
                
            
    'code provided by Mark E.
            TickerStartIndex = 2
            
            
    'loop
        For i = 2 To LastRow
                    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                Ticker = ws.Cells(i, 1).Value
                        
                OpeningPrice = ws.Cells(i, 3).Value
                        
                ClosingPrice = ws.Cells(i, 6).Value
            
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                                            
                YearStartOpen = ws.Cells(TickerStartIndex, 3).Value
                                            
                YearlyChange = ClosingPrice - ws.Cells(TickerStartIndex, 3).Value
                                                               
                PercentChange = YearlyChange / YearStartOpen
                        
                       TickerStartIndex = (i + 1)
                        
    'summary table values
            ws.Cells(SummaryTableRow, 9).Value = Ticker
            ws.Cells(SummaryTableRow, 10).Value = YearlyChange
            ws.Cells(SummaryTableRow, 11).Value = PercentChange
            ws.Cells(SummaryTableRow, 12).Value = TotalStockVolume
           
            
          SummaryTableRow = SummaryTableRow + 1
          
          OpeningPrice = 0
            ClosingPrice = 0
            TotalStockVolume = 0
            PercentChange = 0
          
          Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                                                            
        
          End If
                     
           ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
                     
    'summary table yearly change colors
                     
            If YearlyChange > 0 Then
                 ws.Cells(SummaryTableRow - 1, 10).Interior.ColorIndex = 4
                    
            ElseIf YearlyChange < 0 Then
                  ws.Cells(SummaryTableRow - 1, 10).Interior.ColorIndex = 3
                                   

            End If
                                                           
                     
                Next i
        
           
        Next ws
        
End Sub

