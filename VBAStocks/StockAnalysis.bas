Attribute VB_Name = "Module1"
Sub Calculate_Stock_Stats():

'Variable for ticker symbol
Dim ticker As String

'Variable to serialise tickers
Dim tickerNum As Integer

'Variable to store last row
Dim lastRow As Long

' Variable for Opening Price
Dim O_Price As Double

'variable for closing price
Dim C_Price As Double

'variable for  yearly change
Dim Y_Change As Double

' variable to keep track of percent change
Dim P_Change As Double

' variable to keep track of total stock volume
Dim TStockVol As Double

' variable for Greatest Percent Increase value for the year
Dim GPI As Double

'Ticker that has the Greatest Percent Increase
Dim GPI_ticker As String

' Variable for Greatest Percent Decrease Value for the year
Dim GPD As Double

'Ticker that has the Greatest Percent Decrease.
Dim GPD_ticker As String

' variable to keep track of the greatest stock volume value for specific year.
Dim GStockVol As Double

' ticker that has the Greatest Stock Volume.
Dim GStockVol_ticker As String

' loop over each worksheet in the workbook
For Each ws In Worksheets

    ' Make the worksheet active.
    ws.Activate

    ' Find the last row of each worksheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Add header columns for each worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Initialize variables for each worksheet.
    tickerNum = 0
    ticker = ""
    Y_Change = 0
    O_Price = 0
    P_Change = 0
    TStockVol = 0
    
    ' Skipping the header row, loop through the list of tickers.
    For i = 2 To lastRow

        ' Get the value of the ticker symbol we are currently calculating for.
        ticker = Cells(i, 1).Value
        
        'Fetch Opening Price once
        If O_Price = 0 Then
            O_Price = Cells(i, 3).Value
        End If
        
        ' Add up the total stock volume for the ongoing ticker
        TStockVol = TStockVol + Cells(i, 7).Value
        
        'On encountering a different ticker
        If Cells(i + 1, 1).Value <> ticker Then
            'Update the serial number of tickers
            tickerNum = tickerNum + 1
            'write ticker to column 9
            Cells(tickerNum + 1, 9) = ticker
            
            'fetch closing price for the ticker(only for the end of the year)
            C_Price = Cells(i, 6)
            
            ' Calculate Yearly Change and write to column 10
            Y_Change = C_Price - O_Price
            
            'write yearly change to column 10
           Cells(tickerNum + 1, 10).Value = Y_Change
          
          ' If yearly change value is greater than 0, color cell green.
            If Y_Change > 0 Then
                Cells(tickerNum + 1, 10).Interior.ColorIndex = 4
            ' If yearly change value is less than 0, color cell red.
            ElseIf Y_Change < 0 Then
                Cells(tickerNum + 1, 10).Interior.ColorIndex = 3
            ' If yearly change value is 0, color cell yellow.
            Else
                Cells(tickerNum + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change value for ticker.
            If O_Price = 0 Then
                P_Change = 0
            Else
                P_Change = (Y_Change / O_Price)
            End If
                        
            ' Format the percent_change value as a percent.
            Cells(tickerNum + 1, 11).Value = Format(P_Change, "Percent")
            
            ' Write total stock volume to column 12
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            'Set opening price and Stock Volume to zero
            O_Price = 0
            TStockVol = 0
            
            
            
        End If
        
    Next i
    
    
    
   'BONUS
    Range("O2").Value = "Greatest Percentage Increase"
    Range("O3").Value = "Greatest Percentage Decrease"
    Range("O4").Value = "Greatest Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Get the last row
    lastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    GPI = Cells(2, 11).Value
    GPI_ticker = Cells(2, 9).Value
    GPD = Cells(2, 11).Value
    GPD_ticker = Cells(2, 9).Value
    GStockVol = Cells(2, 12).Value
    GStockVol_ticker = Cells(2, 9).Value
    
    
    ' skipping the header row, loop through the list of tickers.
    For i = 2 To lastRow
    
        ' ticker with the greatest percent increase.
        If Cells(i, 11).Value > GPI Then
            GPI = Cells(i, 11).Value
            GPI_ticker = Cells(i, 9).Value
        End If
        
        'ticker with the greatest percent decrease.
        If Cells(i, 11).Value < GPD Then
            GPD = Cells(i, 11).Value
            GPD_ticker = Cells(i, 9).Value
        End If
        
        ' ticker with the greatest stock volume.
        If Cells(i, 12).Value > GStockVol Then
            GStockVol = Cells(i, 12).Value
            GStockVol_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    'Write to the cells
    Range("P2").Value = Format(GPI_ticker, "Percent")
    Range("Q2").Value = Format(GPI, "Percent")
    Range("P3").Value = Format(GPD_ticker, "Percent")
    Range("Q3").Value = Format(GPD, "Percent")
    Range("P4").Value = GStockVol_ticker
    Range("Q4").Value = GStockVol
    
Next ws


End Sub

