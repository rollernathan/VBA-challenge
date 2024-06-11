Attribute VB_Name = "Module1"
Sub stock_price()

    ' Set an initial variable for holding the stock name
    Dim stock_name As String
    
    ' Set an initial variable for holding the opening price of the quarter per stock name
    Dim stock_open As Double
    stock_open = 0

    ' Set an initial variable for holding the closing price of the quarter per stock name
    Dim stock_close As Double
    stock_close = 0

    ' Set an initial variable for holding the quarterly change of the quarter per stock name
    Dim quarterly_change As Double
    quarterly_change = 0

    ' Set an initial variable for holding the volume
    Dim stock_volume As Double
    stock_volume = 0


    ' Loop through all sheets
    For Each ws In Worksheets
    
    ' Keep track of the location for each stock name in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
        
        ' Find last row of the sheet
        Dim LastRow As Long
        last_row = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        ' Add header row titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' Loop through all stock prices
        For i = 2 To last_row
    
            ' Check if we are still within the same stock, if it is not...
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
                ' Set the stock name
                stock_name = ws.Cells(i, 1).Value
                
                ' Set the opening price of the quarter
                stock_open = ws.Cells(i, 3).Value
                
                ' Print the stock name in the Summary Table
                ws.Cells(summary_table_row, 9).Value = stock_name
                
                ' Add to the stock_volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            ' If the cell immediately following a row is the new stock...
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' set the closing price of the quarter
                stock_close = ws.Cells(i, 6).Value
                
                ' calculate quarterly change
                quarterly_change = stock_close - stock_open
            
                ' Add to the stock_volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                ' Print the quarterly change in the Summary Table
                ws.Cells(summary_table_row, 10).Value = quarterly_change
                
                ' Set the condition for color of the quarterly change
                If quarterly_change < 0 Then
                
                    ' Set the Cell Colors to Red
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
    
                ElseIf quarterly_change > 0 Then
                
                    ' Set the Font Color to Green
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    
                End If
            
                ' Print the percent change in the Summary Table
                ws.Cells(summary_table_row, 11).Value = FormatPercent(quarterly_change / stock_open)
                
                ' Print the stock volume in the Summary Table
                ws.Cells(summary_table_row, 12).Value = stock_volume
     
                ' Reset the stock volume
                stock_volume = 0
    
                ' Add one to the summary table row
                summary_table_row = summary_table_row + 1
        
            Else
                ' Add to the stock_volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
    
            End If
    
        Next i
        
        ' Label Second Table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Set an initial variable for storing maximums and minimums
        Dim greatest_increase As Double
        greatest_increase = ws.Cells(2, 11).Value

        Dim greatest_decrease As Double
        greatest_decrease = ws.Cells(2, 11).Value

        Dim greatest_volume As Double
        greatest_volume = ws.Cells(2, 12).Value

        ' Find greatest percent increase in stock prices
        For i = 2 To last_row
            If ws.Cells(i, 11).Value >= greatest_increase Then
                greatest_increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9)
            End If
        Next i

        ' Assign greatest percent increase to table
        ws.Cells(2, 17).Value = FormatPercent(greatest_increase)

        ' Find greatest percent decrease in stock prices
        For i = 2 To last_row
            If ws.Cells(i, 11).Value <= greatest_decrease Then
                greatest_decrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9)
            End If
        Next i

        ' Assign greatest percent decrease to table
        ws.Cells(3, 17).Value = FormatPercent(greatest_decrease)
        
        ' Find greatest total volume in stock prices
        For i = 2 To last_row
            If ws.Cells(i, 12).Value >= greatest_volume Then
                greatest_volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9)
            End If
        Next i

        ' Assign greatest total volume to table
        ws.Cells(4, 17).Value = greatest_volume
        
        
        ' Autofit to display data
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub

