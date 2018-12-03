Attribute VB_Name = "Module1"
'Tasks -------------------------------------------------
'Loop through each year of stock data and display:
'1. Ticker symbol
'2. Total amount of volume each stock had over the year
'-------------------------------------------------------

Sub Multiple_year_stock_data_easy():

    'Define variables begin -------------------

    'loop counter
    Dim i As Long

    'last row number
    Dim lastRow As Long

    'ticker symbol
    Dim ticker As String

    'total volume
    Dim totalPerTicker As Double

    'column number of ticker in dataset
    Dim tickerColData As Integer

    'column number of total volume in dataset
    Dim volColData As Integer

    'column number of ticker in result
    Dim tickerColResult As Integer

    'column number of total volume in result
    Dim volColResult As Integer

    'counter row number for result set
    Dim resultRow As Long

    'Define variables end --------------------------

    'Initialize variables begin -----------------------
    tickerColData = 1
    volColData = 7
    tickerColResult = 9
    volColResult = 10
    'Initialize variables end ----------------------------

    'Loop through all work sheet
    For Each ws In Worksheets
    
        'Reset total volume per ticker variable
        totalPerTicker = 0
    
        'Reset first row in result set
        resultRow = 2
    
        'Get first ticker name of current sheet
        ticker = ws.Cells(2, tickerColData).Value

        'Get last row number of current sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Write header
        ws.Cells(1, tickerColResult).Value = "Ticker"
        ws.Cells(1, volColResult).Value = "Total Stock Volume"

        'Loop through dataset
        For i = 2 To lastRow

            'Add current line volume to totalPerTicker
            totalPerTicker = totalPerTicker + ws.Cells(i, volColData).Value
    
            'Check whether the next ticker is the same as this one.
            'If yes, this -if- block will not execute.
            'if no, populate result cells and reset variables to make ready for next ticker
            If ws.Cells(i + 1, tickerColData).Value <> ws.Cells(i, tickerColData).Value Then
                'Populate result
                ws.Cells(resultRow, tickerColResult).Value = ticker
                ws.Cells(resultRow, volColResult).Value = totalPerTicker
        
                'Reset variables to make ready for next ticker
                resultRow = resultRow + 1
                ticker = ws.Cells(i + 1, tickerColData).Value
                totalPerTicker = 0
            End If

        Next i

    Next ws
    
End Sub
