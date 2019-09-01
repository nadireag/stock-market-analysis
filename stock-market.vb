Sub StockData():

    For Each ws In Worksheets

   'Create variables to hold ticker value and total
    Dim Ticker As String
    Dim Total_Vol As Double
    Total_Vol = 0

    'Create title colums for ticker and total vol
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"

    'Create a variable to keep tack for ticker and total row
    Dim tablerow As Integer
    tablerow = 2

    'Create lastrow variable to count rows
    Dim lastrow As Long
    lastrow = ws.Application.CountA(Columns(1))
    

        'loop through tickers to find total for each year
        For i = 2 To lastrow

            'Check to see if same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Set tickers
                Ticker = ws.Cells(i, 1).Value
    
                'Add Vol
                Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
                'Print tickers and total vol
                ws.Range("I" & tablerow).Value = Ticker
                ws.Range("J" & tablerow).Value = Total_Vol
    
                'Add one to tablerow to keep track
                tablerow = tablerow + 1
    
                'Reset total
                Total_Vol = 0
    
            'if tickers are different
            Else
    
                'Add total
                Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
            End If
    
    Next i

Next ws

End Sub
