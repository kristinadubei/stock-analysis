Sub Stock_data_summary()

    'Setting up variables for PART 1. Summary table analysis
    Dim ticker_name As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim total_volume As LongLong
    Dim y_change As Double
    Dim p_change As Double
    Dim summary_table_row As Long
    Dim lastrow As Long
    
    'Setting up variables for PART 2. Find Greatest Increase, Descrease and Total Volume
    Dim greatestTicker As String
    Dim greatestValue As Double
    Dim summarylastrow As Long


    'Make this script run on every worksheet in this workbook
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        'PART 1. Summary table -------------------------------------------------------------------

        'count number or rows to plug into for loop    
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        summary_table_row = 2
        
        'create summary table column headers (Ticker, Yearly Change, Percentage Change, Total Volume)
        Range("i1").Value = "Ticker"
        Range("j1").Value = "Yearly Change"
        Range("k1").Value = "Percentage Change"
        Range("l1").Value = "Total Volume"
        
        'assign first ticker name of the table and pick up its opening price and first volume value
        ticker_name = Range("a2")
        opening_price = Range("c2")
        total_volume = Range("g2")
    
        'starting from row 3 data onwards, check each ticker info. I skipped row 2 since I manually assigned starting values for first ticker
        For i = 3 To lastrow
            'if next ticker is different from previous row
            If Cells(i, 1).Value <> ticker_name Then
                'pick up the closing price for current ticker
                closing_price = Cells(i - 1, 6).Value
                'calculate the yearly change
                y_change = closing_price - opening_price
                'calculate the precentage change
                p_change = y_change / opening_price
                'assign ticker name, yearly change, percentage change and total volume values into the summary table
                Cells(summary_table_row, 9).Value = ticker_name
                Cells(summary_table_row, 10).Value = y_change
                Cells(summary_table_row, 11).Value = p_change
                Cells(summary_table_row, 12).Value = total_volume
                
                    'format yearly change cell to be green, red or yellow based on value
                    If Cells(summary_table_row, 10).Value > 0 Then
                        Cells(summary_table_row, 10).Interior.ColorIndex = 4
                    ElseIf Cells(summary_table_row, 10).Value < 0 Then
                        Cells(summary_table_row, 10).Interior.ColorIndex = 3
                    Else
                        'if yearly change is 0 I decided to highlight it with yellow
                        Cells(summary_table_row, 10).Interior.ColorIndex = 6 
                    End If
                
                'format percentage change number as percent
                Cells(summary_table_row, 11).NumberFormat = "0.00%"
                'increment summary table row count by 1
                summary_table_row = summary_table_row + 1
                
                'after the current ticker info is all recorded into the summary table I reset all info for new ticker
                ticker_name = Cells(i, 1).Value
                opening_price = Cells(i, 3).Value
                closing_price = 0
                total_volume = Cells(i, 7).Value
            Else
                'if next cell is same as current ticker then add up the total volume
                total_volume = total_volume + Cells(i, 7).Value
            End If
            
        Next i
        
        'autofit columns
        Columns("A:R").AutoFit
        
        'PART 2. Find Greatest Increase, Descrease and Total Volume -------------------------------------
        
        'count rows in the summary table to plug into for loop later; set holder value to 0
        summarylastrow = Cells(Rows.Count, "I").End(xlUp).Row
        greatestValue = 0
        
        'create table headers
        Range("p2").Value = "Greatest % Increase"
        Range("p3").Value = "Greatest % Decrease"
        Range("p4").Value = "Greatest Total Volume"
        Range("q1").Value = "Ticker"
        Range("r1").Value = "Value"
    
        
        'Loop to find greatest % increase
        For i = 2 To summarylastrow
            'everytime the bigger increase is found it will replace the holder's value and be output into the table
            If Cells(i, 11).Value > greatestValue Then
                greatestValue = Cells(i, 11).Value
                greatestTicker = Cells(i, 9).Value
                Range("q2").Value = greatestTicker
                Range("r2").Value = greatestValue
                Cells(2, 18).NumberFormat = "0.00%"
            End If
        Next i
        
        'reset holder
        greatestValue = 0
        
        'Loop to find greatest % decrease
        For i = 2 To summarylastrow
            'everytime the bigger decrease is found it will replace the holder's value and be output into the table
            If Cells(i, 11).Value < greatestValue Then
                greatestValue = Cells(i, 11).Value
                greatestTicker = Cells(i, 9).Value
                Range("q3").Value = greatestTicker
                Range("r3").Value = greatestValue
                Cells(3, 18).NumberFormat = "0.00%"
            End If
        Next i
        
        'reset holder
        greatestValue = 0
        
        'Loop to find greatest total volume
        For i = 2 To summarylastrow
            'everytime the bigger volume is found it will replace the holder's value and be output into the table
            If Cells(i, 12).Value > greatestValue Then
                greatestValue = Cells(i, 12).Value
                greatestTicker = Cells(i, 9).Value
                Range("q4").Value = greatestTicker
                Range("r4").Value = greatestValue
            End If
        Next i
        
        Columns("A:R").AutoFit
    
    Next ws

End Sub