HW 2 VBA-Challenge

'Use alphabet testing sheet to write script before putting in master doc
'To Do
'Loop through all the stocks for one year and display
'1. the ticker symbol
'2. yearly change of opening price at the beginning and closing price at end of year
'3. percent change of opening price at the beginning and closing price at end of year
'4. total stock volume of the stock
'5. Use conditional formatting to highlight positive change (green) and negative change (red)

Sub Multiple_Year_Stock_Data_Moderate():

'Define variables

'loop counter
Dim i As Long
'last row
Dim LastRow As Long
'ticker symbol
Dim ticker As String
'open value
Dim openVal As Double
'close value
Dim closeVal As Double
'total ticker volume
Dim totalPerTicker As Double
'ticker column data
Dim tickerColData As Integer
'column of Volume
Dim volColData As Integer
'column of open
Dim openColData As Integer
'column of close
Dim closeColData As Integer
'column of ticker results
Dim tickerColResult As Integer
'column of volume results
Dim volColResult As Integer
'column of yearly change results
Dim yearlyChangeColResult As Integer
'column of percent change results
Dim percentChangeColResult As Integer
'yearly change
Dim yearlyChange As Double
'percent change
Dim percentChange As Double
'counter row for results
Dim ResultRow As Long

'assign column values to variables
tickerColData = 1
openColData = 3
closeColData = 6
volColData = 7
tickerColResult = 9
yearlyChangeColResult = 10
percentChangeColResult = 11
volColResult = 12

'loop through worksheet
For Each ws In Worksheets
    'reset total volumer per ticker variable
    totalPerTicker = 0
    'reset first row in results
    ResultRow = 2
    'retrieve ticker name of current sheet
    ticker = ws.Cells(2, tickerColData).Value
    'pull first open value of current ticker in current sheet
    openVal = ws.Cells(2, openColData).Value
    'find last row number in current sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'label headers
    ws.Cells(1, tickerColResult).Value = "Ticker"
    ws.Cells(1, yearlyChangeColResult).Value = "Yearly Change"
    ws.Cells(1, percentChangeColResult).Value = "Percent Change"
    ws.Cells(1, volColResult).Value = "Total Stock Volume"
    
    'loop through dataset
    For i = 2 To LastRow
        If openVal = 0 Then
            openVal = ws.Cells(i + 1, openColData).Value
        Else
            totalPerTicker = totalPerTicker + ws.Cells(i, volColData).Value
            
            If ws.Cells(i + 1, tickerColData).Value <> ws.Cells(i, tickerColData).Value Then
                closeVal = ws.Cells(i, closeColData).Value
                yearlyChange = closeVal - openVal
                percentChange = yearlyChange / openVal
                
                ws.Cells(ResultRow, tickerColResult).Value = ticker
                ws.Cells(ResultRow, yearlyChangeColResult).Value = yearlyChange
                ws.Cells(ResultRow, percentChangeColResult).Value = FormatPercent(Str(percentChange), 2)
                ws.Cells(ResultRow, volColResult).Value = totalPerTicker
                
                If yearlyChange >= 0 Then
                    ws.Cells(ResultRow, yearlyChangeColResult).Interior.ColorIndex = 4
                Else
                    ws.Cells(ResultRow, yearlyChangeColResult).Interior.ColorIndex = 3
                End If
                
                'reset for next ticker
                ResultRow = ResultRow + 1
                ticker = ws.Cells(i + 1, tickerColData).Value
                openVal = ws.Cells(i + 1, openColData).Value
                totalPerTicker = 0
            End If
        End If
    Next i
Next ws
    
    

End Sub