Sub Stock_market():

    For Each ws In Worksheets
  
    
    Dim Ticker_Symbol As String
    Dim Results_Panel_Row As Integer
    Dim Yearly_Change As Double
    Dim Price_Open As Double
    Dim Price_Close As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Great_pct_inc As Double
    Dim Great_pct_dec As Double
    Dim Great_tot_vol As Double
  
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
  

    Total_Stock_Volume = 0
      
    Results_Panel_Row = 2
      
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    
    For i = 2 To LastRow
        
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Price_Open = ws.Cells(i, 3).Value
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                      
            Ticker_Symbol = ws.Cells(i, 1).Value
            Price_Close = ws.Cells(i, 6).Value
                        
            ws.Range("I" & Results_Panel_Row).Value = Ticker_Symbol
                        
            Yearly_Change = Price_Close - Price_Open
            ws.Range("J" & Results_Panel_Row).Value = Yearly_Change
                        
            Percent_Change = (Yearly_Change / Price_Open)
            ws.Range("K" & Results_Panel_Row).Value = Percent_Change
                        
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Results_Panel_Row).Value = Total_Stock_Volume
                        
            Results_Panel_Row = Results_Panel_Row + 1
                  
            Total_Stock_Volume = 0
                 
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
             
        End If
        
    Next i
    
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    Great_pct_inc = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow2).Value)
    ws.Cells(2, 17).Value = Great_pct_inc
           
    Great_pct_dec = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow2).Value)
    ws.Cells(3, 17).Value = Great_pct_dec
                
    Great_tot_vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow2).Value)
    ws.Cells(4, 17).Value = Great_tot_vol
            

    For j = 2 To LastRow2
        
        If ws.Cells(j, 11).Value = Great_pct_inc Then
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                   
        ElseIf ws.Cells(j, 11).Value = Great_pct_dec Then
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                    
        ElseIf ws.Cells(j, 12).Value = Great_tot_vol Then
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
                  
        End If
        
    Next j
      

    For k = 2 To LastRow2
            
        If ws.Cells(k, 10).Value >= 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 4
                
        ElseIf ws.Cells(k, 10).Value < 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
                    
        End If
            
    Next k
         
    For m = 2 To LastRow2
            
            ws.Range("K2:K" & LastRow2).NumberFormat = "0.00%"
            ws.Range("J2:J" & LastRow2).NumberFormat = "0.00"
                
    Next m
                 
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
  Next ws
    

    MsgBox ("Calculations complete")


End Sub

