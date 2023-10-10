Attribute VB_Name = "Module1"
Sub vba_challenge()
' VBA function designed to iterate through large stock ticker data sets
' returns a 4 element by X element long table, with each row containing ticker symbol, yearly change from opening price to closing price, percent change, and total stock volume
'iterates through every sheet in workbook
Dim ws As Worksheet, a As Range

For Each ws In ActiveWorkbook.Worksheets
    'initial variable for number of worksheets
    Dim worksheet_count As Integer

    'initial variables for iterations
    'set initial variable for holding ticker symbol
    Dim ticker_symbol As String
    'set initial variable for holding opening price at beginning of year
    Dim opening_price As Double
    'set initial variable for holding closing price at end of year
    Dim closing_price As Double


    'set variable for yearly change
    Dim yearly_change As Double
    'set variable for percent change
    Dim percent_change As Double

    'set initial variable for holding volume total, set to zero
    Dim total_volume As Double
    total_volume = 0

    'Set summary table headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'bolds summary table headings
    Range("I1:P1").Font.Bold = True
    
    'keep track of stock ticker data summary in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
        'iterate through every used stock info row in individual worksheet
        For i = 2 To 22771
     
            'check to see if ticker symbol is same as previous row (for finding first value of each ticker symbol)
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
                'set ticker symbol
                ticker_symbol = Cells(i, 1).Value
                
                'Print the stock ticker in the Summary Table
                Range("I" & summary_table_row).Value = ticker_symbol
                'set opening price
                opening_price = Cells(i, 3).Value
                       
            'check to see if ticker symbol is same as next row (for finding last value of each ticker symbol)
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           
                'set closing price
                closing_price = Cells(i, 6).Value
             
                'Print the total volume to the Summary Table
                Range("L" & summary_table_row).Value = total_volume
                
                'set yearly change
                yearly_change = closing_price - opening_price
                
                'print the yearly change
                Range("J" & summary_table_row).Value = yearly_change
                
                'set percent change
                percent_change = 100 * (closing_price - opening_price) / opening_price
                
                'print the percent change
                Range("K" & summary_table_row).Value = percent_change
    
                'Add one to the summary table row
                summary_table_row = summary_table_row + 1
    
                'Reset the volume total
                total_volume = 0
    
    
            'If the cell immediately following a row is the same symbol...
            Else
                'Add to the total volume
                total_volume = total_volume + Cells(i, 7).Value
            End If
      Next i
    
    'Set largest value table headings
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    'bolds summary table headings
    Range("N2:N4").Font.Bold = True
    
    
    Dim ticker_number As Integer
        
        For j = 2 To 2000
            If Cells(j, 11).Value > max_value Then
                max_value = Cells(j, 11).Value
                max_ticker = Cells(j, 9).Value
            ElseIf Cells(j, 11).Value < min_value Then
                min_value = Cells(j, 11).Value
                min_ticker = Cells(j, 9).Value
            ElseIf Cells(j, 12).Value > max_vol Then
                max_vol = Cells(j, 12).Value
                max_vol_ticker = Cells(j, 9).Value
            End If
        Next j
    
    Cells(2, 15).Value = max_ticker
    Cells(3, 15).Value = min_ticker
    Cells(4, 15).Value = max_vol_ticker
    
    Cells(2, 16).Value = max_value
    Cells(3, 16).Value = min_value
    Cells(4, 16).Value = max_vol
    
    
    
    'setting coniditional formatting parameters
    Set greater_than_zero = Range("J2:J2000").FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set less_than_zero = Range("J2:J2000").FormatConditions.Add(xlCellValue, xlLess, "=0")
       With greater_than_zero
        .Interior.Color = vbGreen
       End With
       With less_than_zero
         .Interior.Color = vbRed
       End With
Next
End Sub


