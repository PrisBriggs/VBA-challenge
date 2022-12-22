'VBA scripting that analyzes generated stock market data.
'Code created by Priscila Menezes Briggs in December 2022.

'-----------------------------------------------------------INITIAL PROCEDURES------------------------------------------------------------------------------------------------------

'Initializing sub-routine
Sub Stock_market():

    'Code must perform in all Worksheets of the Workbook
    For Each ws In Worksheets
  
    'Setting the variables and its types
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
  
    'Naming columns and rows headers, making sure that Cells are assigned with "ws" in front so the code will perform the same in all worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
  
'-------------------------------------------------------------CALCULATING THE VALUES FOR THE RESULTS PANEL--------------------------------------------------------------------------

    'Making sure that initial Stock Volume is zero so we can add the real values and store them
    Total_Stock_Volume = 0
  
    'Assigning the first row where results should start to be displayed in the Results Panel
    Results_Panel_Row = 2
  
    'Formula to find the last row with data, by finding out the total of rows in column 1 of the worksheet, then from the last cell in Column 1 going
    'back up to the last value/filled cell in the column and then finding out which row it belongs to
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    'Loop through all the rows with raw data
    For i = 2 To LastRow

        'Checking if the ticker symbol is the same until when it's not by finding out if a given cell is different from the one immediately before. This comparison is necessary
        'to find out the first opening price of the year and store it, so it will serve as the basis for the calculation of the Yearly Change in the following IF statement.
        'In summary, this condition will find the first row of the series of the same ticker symbol.
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
            'Setting the value that will be stored by the variable Price_Open, which is the opening price of the stock for a given ticker, once the change in tickers is detected
            Price_Open = ws.Cells(i, 3).Value
        
            'Starting to add the first amount of stock volume for each ticker to the variable that stores the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
        'Checking if the ticker symbol in the next row is different than the one in the current row. In such case, the following values from the current row will be stored, printed
        'or calculated before advancing to the next row
        'In summary, this condition will find the last row of the series of the same ticker symbol.
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
            'Storing the ticker in the variable Ticker_Symbol by keeping the last repeated symbol before it turns to another. Similarly with the closing price, which is being stored
            'under the variable Price_Close
            Ticker_Symbol = ws.Cells(i, 1).Value
            Price_Close = ws.Cells(i, 6).Value
            
            'Printing the Ticker Symbol in the designated cell
            ws.Range("I" & Results_Panel_Row).Value = Ticker_Symbol
            
            'Calculating the yearly change for each ticker by the difference between the previously found values for closing and opening prices, and then printing this value
            Yearly_Change = Price_Close - Price_Open
            ws.Range("J" & Results_Panel_Row).Value = Yearly_Change
            
            'Calculating the Percent Change between the closing and the opening prices and printing it to the designated cell
            Percent_Change = (Yearly_Change / Price_Open)
            ws.Range("K" & Results_Panel_Row).Value = Percent_Change
            
            'Calculating the final Stock Volume for the last row of the ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Results_Panel_Row).Value = Total_Stock_Volume
            
            'Add one to the Results Panel row, so next ticker and its correspondent information will be printed in the next row of the panel
            Results_Panel_Row = Results_Panel_Row + 1
      
            'Reset the Total_Stock_Volume for the next ticker
            Total_Stock_Volume = 0
            
        'If next ticker is the same as current ticker, then execute this command to accumulate the stock volume of a given ticker in the Total_Stock_Volume variable
        Else
            
            'Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                            
        'End of cycle of conditions
        End If
    
    'Goes to the next row until the last row of raw data
    Next i
    
'-------------------------------------------------------------CALCULATING THE MAXIMUM AND MINIMUM VALUES------------------------------------------------------------------------------

    'Same formula used for LastRow, but now looking at column 9
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Formula to calculate the maximum value in a range, and then assigning this value (greatest percent increase) to a designated cell
    Great_pct_inc = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow2).Value)
    ws.Cells(2, 17).Value = Great_pct_inc
       
    'Formula to calculate the minimum value in a range, and then assigning this value (greatest percent decrease) to a designated cell
    Great_pct_dec = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow2).Value)
    ws.Cells(3, 17).Value = Great_pct_dec
            
    'Formula to calculate the maximum value in a range, and then assigning this value (greatest stock volume) to a designated cell
    Great_tot_vol = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow2).Value)
    ws.Cells(4, 17).Value = Great_tot_vol
            
'-------------------------------------------------------------ASSIGNING MAXIMUM AND MINIMUM VALUES TO THEIR CORRESPONDENT TICKER SYMBOLS-----------------------------------------------
    
    'Loop through all the rows with data in the Results Panel
    For j = 2 To LastRow2

        'Checking if the greatest percent increase value matches any of the values in column Percent Change
        If ws.Cells(j, 11).Value = Great_pct_inc Then
            'Retrieving the ticker symbol associated with the greatest percent increase and entering it into the correspondent cell.
            ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
            
        'Checking if the greatest percent decrease value matches any of the values in column Percent Change
        ElseIf ws.Cells(j, 11).Value = Great_pct_dec Then
            'Retrieving the ticker symbol associated with the greatest percent decrease and entering it into the correspondent cell.
            ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
            
        'Checking if the greatest stock volume value matches any of the values in column Total Stock Volume
        ElseIf ws.Cells(j, 12).Value = Great_tot_vol Then
            'Retrieving the ticker symbol associated with the greatest total stock volume and entering it into the correspondent cell.
            ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
            
        'End of IF statement
        End If
    
    
    'Continues the loop through all the rows of the Results Panel
    Next j
      
'-------------------------------------------------------------FORMATTING--------------------------------------------------------------------------------------------------------------
    
    'Applying requested format to the column Yearly Change
    For k = 2 To LastRow2
    
        'Applying color green for positive values
        If ws.Cells(k, 10).Value >= 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 4
        
        'Applying color red for negative values
        ElseIf ws.Cells(k, 10).Value < 0 Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
            
        'End of IF statement
        End If
        
    'Continues the loop through all the rows of the Results Panel
    Next k
     
    'Applying requested format to the columns Yearly Change and Percent Change
    For m = 2 To LastRow2

            'Applying requested format to the Percent Change. Making sure that there are two decimals and the data format is percent.
            ws.Range("K2:K" & LastRow2).NumberFormat = "0.00%"
            
            'Making sure that there are two decimals for values in column Yearly Change.
            ws.Range("J2:J" & LastRow2).NumberFormat = "0.00"
            
    'Continues the loop through all the rows of the Results Panel
    Next m
     
            'Applying requested format to the cells "greatest percent". Making sure that there are two decimals and the data format is percent.
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"
           
  'Perform all the above code in the next worksheet and so on
  Next ws
    
'-------------------------------------------------------------FINAL PROCEDURES--------------------------------------------------------------------------------------------------------
    
    'CALCULATIONS COMPLETE. Message box to warn the user that Excel has finished to apply the code to all sheets.
    MsgBox ("Calculations complete")

'End of the Stock_market subroutine.
End Sub
