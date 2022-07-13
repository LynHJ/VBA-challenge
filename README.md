     Sub stock()

       For Each ws In Worksheets

         'creat summary table
        
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yealy_Change"
            ws.Range("K1").Value = "Percent_Change"
            ws.Range("L1").Value = "Total_Stock_volume"
            
            
            'Initial setting
            Last_Row = ws.Range("A" & Rows.Count).End(xlUp).Row
            Summary_row_count = 2
            Initial_price_for_ticker_name = 2
     '-----------------------------------------------------------------
 
                For I = 2 To Last_Row

                    If (ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value) Then
                    
                        'summary table first column
                            
                            Ticker_Name = ws.Cells(I, 1).Value
                            ws.Range("I" & Summary_row_count).Value = Ticker_Name

                        'summary table second column
                            'last ticker close price- first ticker open price
                            
                            Last_ticker_close_Price = ws.Cells(I, 6).Value
                            
                            First_ticker_open_Price = ws.Cells(Initial_price_for_ticker_name, 3).Value
                            
                            Yearly_Change = Last_ticker_close_Price - First_ticker_open_Price
                            
                            ws.Range("J" & Summary_row_count).Value = Yearly_Change
                            
                        'color set
                            
                            If (ws.Range("J" & Summary_row_count).Value > 0) Then
                            
                            ws.Range("J" & Summary_row_count).Interior.ColorIndex = 4
                            
                            Else
                            
                            ws.Range("J" & Summary_row_count).Interior.ColorIndex = 3
                            
                            End If
                            
                         'summary table third column
                            
                            Percent_Change = (Yearly_Change / First_ticker_open_Price)
                            
                            Percent_Change = Format(Percent_Change, "0.00%")
                            
                            ws.Range("K" & Summary_row_count).Value = Percent_Change
                            
                         'summary table Fourth column
                            
                            Total_Stock_volume = Total_Stock_volume + ws.Cells(I, 7).Value
                            
                            ws.Range("L" & Summary_row_count).Value = Total_Stock_volume
                            
                            Total_Stock_volume = 0
   
                          'summary row control
                        
                            Summary_row_count = Summary_row_count + 1
                        
                            Initial_price_for_ticker_name = I + 1
                            
                            Else
                            
                            Total_Stock_volume = Total_Stock_volume + ws.Cells(I, 7).Value
   
                    
                    End If

                Next I
     '--------------------------------------------------------------------------------
     'BONUS
           'creat a table
            
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
           'Max&min
           
            Rng = ws.Range(ws.Cells(2, 11), ws.Cells(Summary_row_count, 11))
            ws.Range("Q2") = Application.WorksheetFunction.Max(Rng)
            ws.Range("Q3") = Application.WorksheetFunction.Min(Rng)
            
           'Max Vol

            RngVol = ws.Range(ws.Cells(2, 12), ws.Cells(Summary_row_count, 12))
            ws.Range("Q4") = Application.WorksheetFunction.Max(RngVol)
            
           'Xloop
            
            RngT = ws.Range(ws.Cells(2, 9), ws.Cells(Summary_row_count, 9)).Value
            ws.Range("P2") = ws.Application.WorksheetFunction.XLookup(ws.Range("Q2"), Rng, RngT, "not_found")
            ws.Range("P3") = ws.Application.WorksheetFunction.XLookup(ws.Range("Q3"), Rng, RngT, "not_found")
            ws.Range("P4") = ws.Application.WorksheetFunction.XLookup(ws.Range("Q4"), RngVol, RngT, "not_found")

     '------------------------------------------------------------------------------------------------------------
     'makeup
       ws.Range("A:Z").Columns.AutoFit
       ws.Range(ws.Cells(1, 9), ws.Cells(Summary_row_count - 1, 12)).Borders.LineStyle = xlContinuous
       ws.Range(ws.Cells(1, 15), ws.Cells(4, 17)).Borders.LineStyle = xlContinuous
       ws.Range("I1,J1,K1,L1").Interior.ColorIndex = 43
       ws.Range("P1,Q1").Interior.ColorIndex = 43
       ws.Range("O2,O3,O4").Interior.ColorIndex = 46

      Next ws

    End Sub
