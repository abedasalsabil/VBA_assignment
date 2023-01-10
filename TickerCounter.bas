Attribute VB_Name = "Module2"
Sub TickerCounter()

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Volume"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"

Dim Ticker As String
Dim Volume_Total As Double
Dim Summary_table_row As Double
Dim ticker_open_close_counter As Double
Dim yearly_open, yearly_close As Double
Dim j As Integer

Volume_Total = 0
Summary_table_row = 2 ' keep track of row
ticker_open_close_counter = 2 ' keep track of row for opening and closing values

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    Ticker = Cells(i, 1).Value
    Volume_Total = Volume_Total + Cells(i, 7).Value
    yearly_open = Cells(ticker_open_close_counter, 3)

    'summarize if ticker is different
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        yearly_close = Cells(i, "F")
        Range("I" & Summary_table_row).Value = Ticker
        Cells(Summary_table_row, "K").Value = yearly_close - yearly_open
        
        If yearly_open = 0 Then
            Cells(Summary_table_row, "L").Value = Null
            Else
                Cells(Summary_table_row, "L").Value = (yearly_close - yearly_open) / yearly_open
                 
        End If
        
        ' colour positives as green and negatives as red using conditional formatting
        
        If Cells(Summary_table_row, 11).Value > 0 Then
            
                    Cells(Summary_table_row, 11).Interior.ColorIndex = 4
                    
                ElseIf Cells(Summary_table_row, 11).Value < 0 Then
                
                        Cells(Summary_table_row, 11).Interior.ColorIndex = 3
                Else
                        Cells(Summary_table_row, 11).Interior.ColorIndex = 6
                    
            End If
        
        Range("J" & Summary_table_row).Value = Volume_Total
        
        ' use conditional formatting to add % in percent change column
        Cells(Summary_table_row, "L").NumberFormat = "0.00%"
        
        ' increment the data
        Summary_table_row = Summary_table_row + 1
        Volume_Total = 0
        ticker_open_close_counter = i + 1
        
    End If
             
Next i

' calculate greatest % increase, greatest % decrease, and greatest total volume after first part is done

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

Dim greatest_increase, greatest_decrease, greatest_volume As Double
Dim ticker_increase_index, ticker_decrease_index, ticker_volume_index As String

lastrow_ticker = Cells(Rows.Count, "I").End(xlUp).Row

' calculate greatest % increase
greatest_increase = WorksheetFunction.Max(Range("L2:L" & lastrow_ticker))
ticker_increase_index = WorksheetFunction.Index(Range("I2:I" & lastrow_ticker), WorksheetFunction.Match(greatest_increase, Range("L2:L" & lastrow_ticker), 0))
Range("O2").Value = ticker_increase_index
Range("P2").Value = greatest_increase
'format value to %
Range("P2").NumberFormat = "0.00%"

' calculate greatest % decrease
greatest_decrease = WorksheetFunction.Min(Range("L2:L" & lastrow_ticker))
ticker_decrease_index = WorksheetFunction.Index(Range("I2:I" & lastrow_ticker), WorksheetFunction.Match(greatest_decrease, Range("L2:L" & lastrow_ticker), 0))
Range("O3").Value = ticker_decrease_index
Range("P3").Value = greatest_decrease
' format value to %
Range("P3").NumberFormat = "0.00%"

'calculate greatest total volume
greatest_volume = WorksheetFunction.Max(Range("J2:J" & lastrow_ticker))
ticker_volume_index = WorksheetFunction.Index(Range("I2:I" & lastrow_ticker), WorksheetFunction.Match(greatest_volume, Range("J2:J" & lastrow_ticker), 0))
Range("O4").Value = ticker_volume_index
Range("P4").Value = greatest_volume

End Sub
