Attribute VB_Name = "Module1"
Sub stockMarketEval():


    'Declare a variable to iterate until the last row,
    'and a variable that sets the current worksheet name.
    Dim lastRow As Long
    Dim current As Worksheet

    'Loop through the worksheets. Remember to add 'current.'
    'to every Cells() and Range() function at the beginning.
    For Each current In Worksheets

        'Initialize the lastRow variable
        'and set to the last row in each sheet.
        lastRow = current.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create headers for the summary table
        'in each sheet.
        current.Range("J1") = "Ticker"
        current.Range("K1") = "Ticker Volume"
        current.Range("L1") = "Yearly Change"
        current.Range("M1") = "Percent Change"
        
        'Format the columns that need formatting, namely the
        'columns where 'Yearly Change' and 'Percent Change'
        'will be pushed.
        
        
        
        
        'Declare the variable that will hold the aggregate
        'market volume for that ticker. Initialize to zero.
        Dim runVolume As LongLong
        runVolume = 0
        
        'Declare the tickerName variable, which will hold
        'the value for the ticker name within the next 'for' loop.
        Dim tickerName As String
        
        'Declare tickerTableRow, which will hold
        'the pointer to the summary table for new ticker entries.
        Dim tickerTableRow As Integer
        tickerTableRow = 2

        'Set the first value for openPrice, which holds the first
        'opening price for the first ticker.
        Dim openPrice As Double
        openPrice = current.Cells(2, 3).Value
    
            'Loop through each individual sheet from 2,
            'which is the first row after the header, until
            'the last row, delcared above.
            For j = 2 To lastRow
                
                'Declare closePrice, which will hold the closing price
                'for that ticker until the pointer reaches the end.
                Dim closePrice As Double
                closePrice = current.Cells(j, 6).Value
                
                'Declare and initialize yChange, which will hold
                'the yearly change in price for the year for that ticker.
                Dim yChange As Double
                yChange = 0
                
                'Declare and initialize pChange, which will hold
                'the percent change in price for the year for that ticker.
                Dim pChange As Double
                pChange = 0
                
                
                'Compare the current iteration pointing to the current ticker
                'to the next. If they're different, execute the following code:
                If current.Cells(j + 1, 1).Value <> current.Cells(j, 1).Value Then
                
                    'tickerName gets initialized to the last entry for that ticker symbol
                    'for that year, the final market volume entry gets added to the aggregate volume,
                    'the close price is updated to the last entry for close price for that symbol,
                    'and yearly change is calculated.
                    tickerName = current.Cells(j, 1).Value
                    runVolume = runVolume + current.Cells(j, 7).Value
                    closePrice = current.Cells(j, 6).Value
                    yChange = closePrice - openPrice
                    
                    
                    
                    
                
                    'Pushes to the summary table the current tickerName,
                    'the total aggregate market volume,
                    'and the yearly change.
                    current.Range("J" & tickerTableRow).Value = tickerName
                    current.Range("K" & tickerTableRow).Value = runVolume
                    current.Range("L" & tickerTableRow).NumberFormat = "$#,##0.00"
                    current.Range("L" & tickerTableRow).Value = yChange
                    
                    'Colors the cells with yearly change red or green according
                    'to their value. If positive, green, else negative, red.
                    If current.Range("L" & tickerTableRow).Value < 0 Then
                       current.Range("L" & tickerTableRow).Interior.ColorIndex = 3
                        
                    ElseIf current.Range("L" & tickerTableRow).Value > 0 Then
                        current.Range("L" & tickerTableRow).Interior.ColorIndex = 4
                    Else
                    End If

                    
                    
                    
                    'Takes into consideration an error you get when dividing by zero.
                    If openPrice = 0 Then
                        current.Range("M" & tickerTableRow).Value = "N/A"
                    Else
                        pChange = yChange / openPrice * 100
                        current.Range("M" & tickerTableRow).NumberFormat = "General\%"
                        current.Range("M" & tickerTableRow).Value = pChange
                    End If

                    
                    'Change the pointer to the summary table by adding one.
                    tickerTableRow = tickerTableRow + 1
                
                    'Reset the aggregate market volume.
                    runVolume = 0

                    'Change the value of your closing price and open price to be the next ticker.
                    'DO NOT change the open price when you restart the loop.
                    closePrice = current.Cells(j + 1, 6).Value
                    openPrice = current.Cells(j + 1, 3).Value


                    'Reset the yearly change and percent change values.
                    yChange = 0
                    pChange = 0
                    
                    
                'If the current ticker symbol is the same as the next ticker symbol,
                'just keep adding to the aggregate market volume
                'and change the close price to the next row.
                Else
                    runVolume = runVolume + current.Cells(j, 7).Value
                    closePrice = current.Cells(j, 6).Value
                
                End If
            Next j
    Next current
            
End Sub

