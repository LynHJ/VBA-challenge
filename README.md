    Sub stock()
   
       For Each ws In Worksheets
    
             'creat summary table
            
                ws.range("I1").Value = "Ticker"
                ws.range("J1").Value = "Yealy_Change"
                ws.range("K1").Value = "Percent_Change"
                ws.range("L1").Value = "Total_Stock_volume"
                
                
                'Initial setting
                Last_Row = ws.range("A" & Rows.Count).End(xlUp).Row
                Summary_row_count = 2
                Initial_price_for_ticker_name = 2
      '-----------------------------------------------------------------
                    For i = 2 To Last_Row
 
                        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                        
                            'summary table first column
                                
                                Ticker_Name = ws.Cells(i, 1).Value
                                ws.range("I" & Summary_row_count).Value = Ticker_Name
    
                            'summary table second column
                                'last ticker close price- first ticker open price
                                
                                Last_ticker_close_Price = ws.Cells(i, 6).Value
                                
                                First_ticker_open_Price = ws.Cells(Initial_price_for_ticker_name, 3).Value
                                
                                Yearly_Change = Last_ticker_close_Price - First_ticker_open_Price
                                
                                ws.range("J" & Summary_row_count).Value = Yearly_Change
                                
                            'color set
                                
                                If (ws.range("J" & Summary_row_count).Value > 0) Then
                                
                                ws.range("J" & Summary_row_count).Interior.ColorIndex = 4
                                
                                Else
                                
                                ws.range("J" & Summary_row_count).Interior.ColorIndex = 3
                                
                                End If
                                
                             'summary table third column
                                
                                Percent_Change = (Yearly_Change / First_ticker_open_Price)
                                
                                Percent_Change = Format(Percent_Change, "0.00%")
                                
                                ws.range("K" & Summary_row_count).Value = Percent_Change
                                
                             'summary table Fourth column
                                
                                Total_Stock_volume = Total_Stock_volume + ws.Cells(i, 7).Value
                                
                                ws.range("L" & Summary_row_count).Value = Total_Stock_volume
                                
                                ws.range("L" & Summary_row_count).NumberFormat = "0"
                                
                                Total_Stock_volume = 0
       
                              'summary row control
                            
                                Summary_row_count = Summary_row_count + 1
                            
                                Initial_price_for_ticker_name = i + 1
                                
                                Else
                                
                                Total_Stock_volume = Total_Stock_volume + ws.Cells(i, 7).Value
       
                        
                        End If

                    Next i
     '--------------------------------------------------------------------------------
     'BONUS
            'creat a table
                ws.range("O2").Value = "Greatest % Increase"
                ws.range("O3").Value = "Greatest % Decrease"
                ws.range("O4").Value = "Greatest Total Volume"
                ws.range("P1").Value = "Ticker"
                ws.range("Q1").Value = "Value"
                
                'Max&min
               
                Rng = ws.range(ws.Cells(2, 11), ws.Cells(Summary_row_count, 11))
                ws.range("Q2") = Application.WorksheetFunction.Max(Rng)
                ws.range("Q3") = Application.WorksheetFunction.Min(Rng)
                
                'Max Vol
                
                RngVol = ws.range(ws.Cells(2, 12), ws.Cells(Summary_row_count, 12))
                ws.range("Q4") = Application.WorksheetFunction.Max(RngVol)
                ws.range("Q4").NumberFormat = "0"
                
                'Xloop
                
                RngT = ws.range(ws.Cells(2, 9), ws.Cells(Summary_row_count, 9)).Value
                ws.range("P2") = ws.Application.WorksheetFunction.XLookup(ws.range("Q2"), Rng, RngT, "not_found")
                ws.range("P3") = ws.Application.WorksheetFunction.XLookup(ws.range("Q3"), Rng, RngT, "not_found")
                ws.range("P4") = ws.Application.WorksheetFunction.XLookup(ws.range("Q4"), RngVol, RngT, "not_found")

     '--------------------------------------------------------------------------------------------------------------------
             'makeup
              ws.range("A:Z").Columns.AutoFit
              ws.range(ws.Cells(1, 9), ws.Cells(Summary_row_count - 1, 12)).Borders.LineStyle = xlContinuous
              ws.range(ws.Cells(1, 15), ws.Cells(4, 17)).Borders.LineStyle = xlContinuous
              ws.range("I1,J1,K1,L1").Borders.Weight = xlMedium
              ws.range("I1,J1,K1,L1").Interior.ColorIndex = 43
              ws.range("P1,Q1").Interior.ColorIndex = 43
              ws.range("O2,O3,O4").Interior.ColorIndex = 46
 
         Next ws
   
      End Sub


