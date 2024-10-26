Attribute VB_Name = "Module1"
Sub StockTicker():

    'establishing a for loop to loop across the different sheets in the file
    For Each ws In Worksheets
        
        'establishing column headers for Ticker, Quarterly Change, Percent Change, and Total Stock Volume
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'adjust column width so that headers are legible
        ws.Columns().AutoFit
        
        'we want to descend through column A for everywhere the ticker symbol changes, from row 2 to the last populated row
        'establish the last populated row as a variable lastRow, this will allow the code to be run on sheets with different numbers of rows
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'set variable TickerSymbol as a string, to hold the ticker symbol
        Dim TickerSymbol As String
        
        'set variable FirstOpen as double, to hold the first opening price for a ticker symbol in the quarter
        Dim FirstOpen As Double
        FirstOpen = ws.Cells(2, 3).Value 'save the contents of cell C2 as the first opening price of the sheet
        
        'set variable FinalClose as double, to hold the final close price for a ticker symbol in the quarter
        Dim FinalClose As Double
        
        'set variable TotalVolume as a double, to hold the Total Stock Volume.
        Dim TotalVolume As Double
        
        'initialize TotalVolume as zero
        TotalVolume = 0
        
        'set a variable to hold the row location of Ticker and Volume outputs, since we have headers in row 1, we'll start with row 2 as the first location
        Dim StockRow As Integer
        StockRow = 2
    
        'loop through all of the rows, starting at row 2 and ending at the last row
        For Row = 2 To lastRow
        
            'check to see if ticker symbol has changed by comparing values of adjacent rows
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
                    'check if the FirstOpen price (stored at the end of this for loop) is equal to zero.
                    'If it is, select the next opening price in the sequence, as the zero will make percentage change calculation impossible.
                    If FirstOpen = 0 Then
                        FirstOpen = ws.Cells(Row + 2, 3).Value
                    End If
            
                'save ticker symbol for row just before changeover as TickerSymbol
                TickerSymbol = ws.Cells(Row, 1).Value
                
                'put TickerSymbol in the Ticker column, Column I (9)
                ws.Cells(StockRow, 9).Value = TickerSymbol
                
                
                'save closing price, located in Column F (6) at the point of Ticker Symbol changeover
                FinalClose = ws.Cells(Row, 6).Value
                
                'put Quarterly Change into column J (10)
                ws.Cells(StockRow, 10).Value = FinalClose - FirstOpen
                
                'format Quarterly Change with two decimal places
                ws.Cells(StockRow, 10).NumberFormat = "0.00"
                
                'put Percentage Change into Column K (11)
                ws.Cells(StockRow, 11).Value = (FinalClose - FirstOpen) / FirstOpen
                
                'format Percentage Change as a percent
                ws.Cells(StockRow, 11).NumberFormat = "0.00%"
                
                
                'add the TotalVolume value for that row (found in Column G(7)) to the Total Volume (this is the last value to be added to the running volume of the TickerSymbol)
               TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                
                'Add the TotalVolume to Column L
                ws.Range("L" & StockRow).Value = TotalVolume
                
                'format Quarterly Change in column J to be red if less than 0 and green if greater than 0
                If ws.Range("J" & StockRow).Value < 0 Then
                    ws.Range("J" & StockRow).Interior.ColorIndex = 3 'turns cells less than 0 red
                ElseIf ws.Range("J" & StockRow).Value > 0 Then
                    ws.Range("J" & StockRow).Interior.ColorIndex = 4 'turns cells greater than 0 green
                End If
                
                
                'reset the TotalVolume to zero, for usage with the next Ticker Symbol
                TotalVolume = 0
                
                'looks at the first row after the symbol change, logs the value in column C(3) as the first opening price, which will be used in the next loop.
                FirstOpen = ws.Cells(Row + 1, 3).Value
                    

                
                'StockRow moves down one row, to continue with the next Ticker Symbol
                StockRow = StockRow + 1
                
                
                
            Else 'if the Ticker Symbol does not change, add the listed volume in Column G(7) to the Total Stock Volume for the current Ticker Symbol, which will populate column L when a new Ticker Symbol is discovered
                TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
                
            End If 'ends if statement
            
        Next Row 'moves to the next row in the for loop
        
        
    
        
        
        'add labels for the Ticker and Value of Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'autofit the columns to contain the headers
        ws.Columns().AutoFit
        
        'calculate max value in Percent Change column (column K) as Greatest % Increase, store it in cell Q2
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & StockRow))
        
        'Format Max % as a percentage
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'calculate minimum value in Percent Change column (column K) as Greatest % Decrease, store it in cell Q3
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & StockRow))
        
        'Format Min % as a percentage
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'calculate Greatest Total Volume using max function on Total Stock Volume column (column L), store it in cell Q4
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & StockRow))
        
        
        
        'for loop to cycle through percent change and total stock volume columns, from row 2 to the last row of the sheet.
        For Row = 2 To lastRow
        
            'if a value in column K (11) matches the Max % stored in Q2, print the corresponding ticker symbol from column I (9) in Cell P2
            If ws.Cells(Row, 11).Value = ws.Range("Q2").Value Then
                ws.Range("P2").Value = ws.Cells(Row, 9).Value
                
            'if a value in column K (11) matches the Min % stored in Q3, print the corresponding ticker symbol from column I (9) in Cell P3
            ElseIf ws.Cells(Row, 11).Value = ws.Range("Q3").Value Then
                ws.Range("P3").Value = ws.Cells(Row, 9).Value
                
                
            'if a value in column L (12) matches the Greatest Total Volume stored in Q4, print the corresponding ticker symbol from column I (9) in Cell P4
            ElseIf ws.Cells(Row, 12).Value = ws.Range("Q4").Value Then
                ws.Range("P4").Value = ws.Cells(Row, 9).Value
                
        End If 'ends if statement
        
        Next Row 'moves to the next row in the for loop
        
    Next ws 'moves to the next sheet in the file
    
End Sub

