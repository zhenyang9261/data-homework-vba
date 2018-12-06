Attribute VB_Name = "Module3"
'Tasks -------------------------------------------------
'Loop through each year of stock data and diaplay
'
'Table 1:
'1. Ticker symbol
'2. Yearly change from what the stock opened the year at to what the closing price was.
'3. The percent change from the what it opened the year at to what it closed.
'4. Total amount of volume each stock had over the year.
'5. Highlight positive change in green and negative change in red
'
'Table 2:
'1. Greatest % increase
'2. Greatest % decrease
'3. Greatest total volume
'-------------------------------------------------------

Sub Multiple_year_stock_data_hard():

    'Define variables begin -------------------

    'loop counter
    Dim i As Long

    'last row number
    Dim lastRow As Long

    'ticker symbol
    Dim ticker As String

    'open value
    Dim openVal As Double

    'close value
    Dim closeVal As Double

    'total volume
    Dim totalPerTicker As Double

    'column number of ticker in dataset
    Dim tickerColData As Integer

    'column number of Volume in dataset
    Dim volColData As Integer

    'column number of open in dataset
    Dim openColData As Integer

    'column number of close in dataset
    Dim closeColData As Integer

    'column number of Ticker in result
    Dim tickerColResult As Integer

    'column number of Volume in result
    Dim volColResult As Integer

    'column number of Yearly Change in result
    Dim yearlyChangeColResult As Integer

    'column number of Percent Change in result
    Dim percentChangeColResult As Integer
    
    'column number for table2 labels
    Dim labelColResult2 As Integer
    
    'column number for table2 ticker
    Dim tickerColResult2 As Integer
    
    'column number or table2 value
    Dim valueColResult2 As Integer

    'Yearly Change
    Dim yearlyChange As Double

    'Percent Change
    Dim percentChange As Double

    'counter row number for result set
    Dim resultRow As Long
    
    'greatest % increase
    Dim greatestPercentIncrease As Double
    
    'greatest % increase ticker
    Dim tickerGreatestIncrease As String
    
    'greatest % decrease ticker
    Dim tickerGreatestDecrease As String
    
    'greatest total volume ticker
    Dim tickerGreatestTotal As String

    'greatest % decrease
    Dim greatestPercentDecrease As Double
    
    'greatest total volume
    Dim greatestTotal As Double

    'Define variables end --------------------------

    'Initialize column numbers begin ---------------------
    tickerColData = 1
    openColData = 3
    closeColData = 6
    volColData = 7
    tickerColResult = 9
    yearlyChangeColResult = 10
    percentChangeColResult = 11
    volColResult = 12
    labelColResult2 = 15
    tickerColResult2 = 16
    valueColResult2 = 17
    'Initialize column numbers end ------------------------
    
    'Loop through all work sheet
    For Each ws In Worksheets
    
        'Reset total volume per ticker variable
        totalPerTicker = 0
        
        'Reset greatest % increase for this sheet
        greatestPercentIncrease = 0
        
        'Reset greatest % decrease for this sheet
        greatestPercentDecrease = 0
        
        'Reset greatest total volume for this sheet
        greatestTotal = 0
    
        'Reset first row in result set
        resultRow = 2
    
        'Get first ticker name of current sheet
        ticker = ws.Cells(2, tickerColData).Value
        tickerGreatestIncrease = ws.Cells(2, tickerColData).Value
        tickerGreatestDecrease = ws.Cells(2, tickerColData).Value
        tickerGreatestTotal = ws.Cells(2, tickerColData).Value
        
        'Get first open value of current ticker in current sheet
        openVal = ws.Cells(2, openColData).Value

        'Get last row number of current sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Write header - result table 1
        ws.Cells(1, tickerColResult).Value = "Ticker"
        ws.Cells(1, yearlyChangeColResult).Value = "Yearly Change"
        ws.Cells(1, percentChangeColResult).Value = "Percent Change"
        ws.Cells(1, volColResult).Value = "Total Stock Volume"
        
        'Write header and labels - result table 2
        ws.Cells(1, tickerColResult2).Value = "Ticker"
        ws.Cells(1, valueColResult2).Value = "Value"
        ws.Cells(2, labelColResult2).Value = "Greatest % Increase"
        ws.Cells(3, labelColResult2).Value = "Greatest % Decrease"
        ws.Cells(4, labelColResult2).Value = "Greatest Total Volume"

        'Loop through dataset
        For i = 2 To lastRow

            'If open value is 0, this row is bad data, ignore, continue to next row
            If openVal = 0 Then
                openVal = ws.Cells(i + 1, openColData).Value
            
            Else
            
                'Add current line volume to totalPerTicker
                totalPerTicker = totalPerTicker + ws.Cells(i, volColData).Value
    
                'Check whether the next ticker is the same as this one.
                'If yes, this -if- block will not execute.
                'if no,
                '1. calculate result,
                '2. populate result cells in table1,
                '3. reset variables to make ready for next ticker
                '4. check whether the current percent change is the greatest % increase/decrease,
                '   check whether the current total volume is the greatest total volume,
                '   if yes, store the current value as greatest % increase/greatest % decrease/greatest total
                '   In cases above, also store the current ticker
                If ws.Cells(i + 1, tickerColData).Value <> ws.Cells(i, tickerColData).Value Then
    
                    'Calculate Yearly Change and Percent Change
                    closeVal = ws.Cells(i, closeColData).Value
                    yearlyChange = closeVal - openVal
                    percentChange = yearlyChange / openVal
        
                    'Populate result
                    ws.Cells(resultRow, tickerColResult).Value = ticker
                    ws.Cells(resultRow, yearlyChangeColResult).Value = yearlyChange
                    ws.Cells(resultRow, percentChangeColResult).Value = FormatPercent(Str(percentChange), 2)
                    ws.Cells(resultRow, volColResult).Value = totalPerTicker
        
                    'Color Yearly Change cell. Green if positive. Red if negtive
                    If yearlyChange >= 0 Then
                        ws.Cells(resultRow, yearlyChangeColResult).Interior.ColorIndex = 4
                    Else
                        ws.Cells(resultRow, yearlyChangeColResult).Interior.ColorIndex = 3
                    End If
                    
                    'Check whether this percent change is bigger than previous greatest % increase,
                    'if yes, store this one as greatest % increase, also store the ticker
                    If percentChange > greatestPercentIncrease Then
                        greatestPercentIncrease = percentChange
                        tickerGreatestIncrease = ticker
                    End If
                        
                    'Check whether this percent change is smaller than previous greatest % decrease,
                    'if yes, store this one as greatest % decrease, also store the ticker
                    If percentChange < greatestPercentDecrease Then
                        greatestPercentDecrease = percentChange
                        tickerGreatestDecrease = ticker
                    End If
                        
                    'Check whether this total volume is bigger than previous greatest total,
                    'if yes, store this one as greatest total, also store the ticker
                    If totalPerTicker > greatestTotal Then
                        greatestTotal = totalPerTicker
                        tickerGreatestTotal = ticker
                    End If
                    
                    'Reset variables to make ready for next ticker
                    resultRow = resultRow + 1
                    ticker = ws.Cells(i + 1, tickerColData).Value
                    openVal = ws.Cells(i + 1, openColData).Value
                    totalPerTicker = 0
                    
                End If
            
            End If

        Next i
        
        'Populate result cells with the result from the for loop
        ws.Cells(2, tickerColResult2).Value = tickerGreatestIncrease
        ws.Cells(2, valueColResult2).Value = FormatPercent(Str(greatestPercentIncrease), 2)
        
        ws.Cells(3, tickerColResult2).Value = tickerGreatestDecrease
        ws.Cells(3, valueColResult2).Value = FormatPercent(Str(greatestPercentDecrease), 2)
        
        ws.Cells(4, tickerColResult2).Value = tickerGreatestTotal
        ws.Cells(4, valueColResult2).Value = greatestTotal
        
    
    Next ws

End Sub


