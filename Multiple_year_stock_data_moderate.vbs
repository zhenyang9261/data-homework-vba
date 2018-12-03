Attribute VB_Name = "Module2"
'Tasks -------------------------------------------------
'Loop through each year of stock data and diaplay
'1. Ticker symbol
'2. Yearly change from what the stock opened the year at to what the closing price was.
'3. The percent change from the what it opened the year at to what it closed.
'4. Total amount of volume each stock had over the year.
'5. Highlight positive change in green and negative change in red
'-------------------------------------------------------

Sub Multiple_year_stock_data_moderate():

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

    'Yearly Change
    Dim yearlyChange As Double

    'Percent Change
    Dim percentChange As String

    'counter row number for result set
    Dim resultRow As Long

    'Define variables end --------------------------

    'Initialize variables begin ---------------------
    tickerColData = 1
    openColData = 3
    closeColData = 6
    volColData = 7
    tickerColResult = 9
    yearlyChangeColResult = 10
    percentChangeColResult = 11
    volColResult = 12
    'Initialize variables end ------------------------
    
    'Loop through all work sheet
    For Each ws In Worksheets
    
        'Reset total volume per ticker variable
        totalPerTicker = 0
    
        'Reset first row in result set
        resultRow = 2
    
        'Get first ticker name of current sheet
        ticker = ws.Cells(2, tickerColData).Value
        
        'Get first open value of current ticker in current sheet
        openVal = ws.Cells(2, openColData).Value

        'Get last row number of current sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Write header
        ws.Cells(1, tickerColResult).Value = "Ticker"
        ws.Cells(1, yearlyChangeColResult).Value = "Yearly Change"
        ws.Cells(1, percentChangeColResult).Value = "Percent Change"
        ws.Cells(1, volColResult).Value = "Total Stock Volume"

        'Loop through dataset
        For i = 2 To lastRow

            'Add current line volume to totalPerTicker
            totalPerTicker = totalPerTicker + ws.Cells(i, volColData).Value
    
            'Check whether the next ticker is the same as this one.
            'If yes, this -if- block will not execute.
            'if no,
            '1. calculate result,
            '2. populate result cells,
            '3. reset variables to make ready for next ticker
            If ws.Cells(i + 1, tickerColData).Value <> ws.Cells(i, tickerColData).Value Then
    
                'Calculate Yearly Change
                closeVal = ws.Cells(i, closeColData).Value
                yearlyChange = FormatNumber((closeVal - openVal), 9)
        
                'Calculate Percent Change. In case open value equals 0, set percentChange to Infinite
                If openVal = 0 Then
                    percentChange = "Infinite"
                Else
                    percentChange = FormatPercent(Str(yearlyChange / openVal), 2)
                End If
        
                'Populate result
                ws.Cells(resultRow, tickerColResult).Value = ticker
                ws.Cells(resultRow, yearlyChangeColResult).Value = yearlyChange
                ws.Cells(resultRow, percentChangeColResult).Value = percentChange
                ws.Cells(resultRow, volColResult).Value = totalPerTicker
        
                'Color Yearly Change cell. Green if positive. Red if negtive
                If yearlyChange >= 0 Then
                    ws.Cells(resultRow, yearlyChangeColResult).Interior.ColorIndex = 4
                Else
                    ws.Cells(resultRow, yearlyChangeColResult).Interior.ColorIndex = 3
                End If
        
                'Reset variables to make ready for next ticker
                resultRow = resultRow + 1
                ticker = ws.Cells(i + 1, tickerColData).Value
                openVal = ws.Cells(i + 1, openColData).Value
                totalPerTicker = 0
            End If

        Next i
    
    Next ws

End Sub

