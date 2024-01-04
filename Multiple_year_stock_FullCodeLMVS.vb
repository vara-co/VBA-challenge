Sub YearlyStock()

'Declare the worksheet and variables
For Each ws In Worksheets

    'Define Variables to hold values
    Dim LastRow As Long
    Dim Ticker As String
    Dim openingP As Double
    Dim closingP As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim TotalVol As Double
    Dim Summary_Table_Row As Integer
    
    'Find the last row for the whole dataset
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Summary Table to input in Row 2
    Summary_Table_Row = 2
    
'----------------------------------------------------------------------
'------------                 L  O  O  P                   ------------
'----------------------------------------------------------------------

    'Loop through each row
    For i = 2 To LastRow
        
        '-----------------------------------------------
        '--  Start your conditionals / IF statements  --
        '-----------------------------------------------
        
        '**This 1st If statement, contains the Ticker value and OpeningP**
        
        'Check if current row has different ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
        'Set your Ticker name
            Ticker = ws.Cells(i, 1).Value
        
        'Get opening price for the new year
            openingP = ws.Cells(i, 3).Value
        
        'Reset total vol for the new year
            TotalVol = 0
            
        'End of 1st If statement
        End If
        
        '**This 2nd If statement, contains the closingP value and
        '**the calculations: Yearly Change, Percent Change, and Total Volume.
        '**Note that your End If closes at the end of this Main Loop.
        '**Be careful with the indentation for the nested loops.
        
        'Check if current row is the last one for the current ticker
        'You can use "Ticker" because the value was already assigned above,
        'or "ws.Cells(i, 1).Value"
        If i = LastRow Or ws.Cells(i + 1, 1).Value <> Ticker Then
        
        'Since it is not the same, add closing Price
            closingP = ws.Cells(i, 6).Value
            
        'Calculate yearly change
            yearlyChange = closingP - openingP
            'If I were to add the $ symbol to the yearly column I would
            'use the next line in comments, however, it encloses the
            'negative numbers within parenthesis and now it won't match
            'the homework's image. So be aware that I know how to do it
            'Your instructions are not clear about which to follow
            'ws.Cells(Summary_Table_Row, 10).Style = "Currency"
            
        'Avoiding division by zero
            If openingP <> 0 Then
            'calculate the percent change
                percentChange = (yearlyChange / openingP)
            Else
                percentChange = 0
            End If
        
        'Add the Total Volume
        TotalVol = TotalVol + ws.Cells(i, 7).Value
        
        'Print the Values in the Summary Table Row
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("J" & Summary_Table_Row).Value = yearlyChange
        ws.Range("K" & Summary_Table_Row).Value = percentChange
        ws.Range("L" & Summary_Table_Row).Value = TotalVol
        
        'Trying color format 2
        If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4 'Green
            Else
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3 'Red
            End If
            
        'Add one to the summary table row to continue in the next row when it prints
        Summary_Table_Row = Summary_Table_Row + 1
        
        
        '--------IF THE SAME VALUES ELSE ------------
        'If the cell immediately below to the current is the same ticker.
        'Note that this Else statement is part of the Main Loop but it is attached
        'as part of the 2nd Loop. I'm aware there are other ways to write this.
        Else
        
        'Add the TotalVol, when the ticker is the same as above
        TotalVol = TotalVol + ws.Cells(i, 7).Value
              
        'This is the *End If* for the Main Loop, on the 2nd If Statement
        End If
    
        'This is the *next i* for the Whole Loop.
        'It will loop through the columns that calculates:
        'yearlyChange, percentChange, & TotalVol and change the colors
    Next i
    
'-------------------------------------------------------------------------
'----                  FORMATING OUTSIDE THE LOOP                     ----
'-------------------------------------------------------------------------
    
    'Add format to the Percent Change column !! This is correct
    ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
    
    'Add format to the Total Stock Volume to number !! This is correct
    ws.Cells(Summary_Table_Row, 12).NumberFormat = "0"
    
    'Add format to make sure the content of the cells fit accordingly
    ws.Columns("A:R").AutoFit
    
'-------------------------------------------------------------------------
'----               OUTPUT THE VALUES INTO ROW HEADER                 ----
'-------------------------------------------------------------------------

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'These two Row Header values are for the second part of the challenge
    'Which are part of the 2nd Summary Header
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
'------------------------------------------------------------------------
'----     Greatest % Increase, Decrease and Greatest Total Vol       ----
'-----                    In a 2nd summary table                    -----
'------------------------------------------------------------------------
    
    'If ws.range("I:I").Value (Figure out the way to say greatest value) Then
    'print the result in ws.cell(2,16).Value for ticker and ws.cell(2,17).Value
    'for the Value
    
    'Return stock with the greatest % increase
    Dim MaxPercentInc As Double
    Dim TickerMaxPercentInc As String
        
    'Return stock with the greatest % decrease
    Dim MaxPercentDec As Double
    Dim TickerMaxPercentDec As String

    'Return stock with the greatest total volume
    Dim MaxTotalVol As LongLong
    Dim TickerMaxTotalVol As String
        
    'Reset the Variables
    MaxPercentInc = 0
    MaxPercentDec = 0
    MaxTotalVol = 0
        
    'Loop through the 2nd summary table to find max percent
    'increase, decrease, and total
    For i = 2 To Summary_Table_Row - 1
        
        If ws.Cells(i, 11).Value > MaxPercentInc Then
            MaxPercentInc = ws.Cells(i, 11).Value
            TickerMaxPercentInc = ws.Cells(i, 9).Value
        End If
            
        If ws.Cells(i, 11).Value < MaxPercentDec Then
            MaxPercentDec = ws.Cells(i, 11).Value
            TickerMaxPercentDec = ws.Cells(i, 9).Value
        End If
            
        If ws.Cells(i, 12).Value > MaxTotalVol Then
            MaxTotalVol = ws.Cells(i, 12).Value
            TickerMaxTotalVol = ws.Cells(i, 9).Value
        End If
    Next i
        
'--------------------------------------------------------------------
'------------  Output the values to the greatest % ------------------
'------------  Increase,Decrease, and Total Volume ------------------
'------------        In a 2nd Summary Table        ------------------
'--------------------------------------------------------------------
    
    'Output the values for this summary
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Increase"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
        
    'Output the Ticker Values for Max PercentInc/Dec/TotalVol
    ws.Cells(2, 16).Value = TickerMaxPercentInc
    ws.Cells(3, 16).Value = TickerMaxPercentDec
    ws.Cells(4, 16).Value = TickerMaxTotalVol
    
    'Output Max PercentInc/Dec/TotalVol
    ws.Cells(2, 17).Value = MaxPercentInc
    ws.Cells(3, 17).Value = MaxPercentDec
    ws.Cells(4, 17).Value = MaxTotalVol
        
    'Format Max PercentInc/Dec/TotalVol to display 2 decimals % and long numbers
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0"
       
    
'This is to go to the next worksheet and do the whole thing again
Next ws

End Sub