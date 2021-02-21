Sub stockticker():

For Each ws In Worksheets

On Error Resume Next
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
tickerrow = 2

ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

ws.Cells(1, "O").Value = "Ticker"
ws.Cells(1, "P").Value = "Value"
ws.Cells(2, "N").Value = "Greatest % Increase"
ws.Cells(3, "N").Value = "Greatest % Decrease"
ws.Cells(4, "N").Value = "Greatest Total Volume"

    For i = 1 To (lastrow - 1)
    
        'If to generate Ticker list
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'MsgBox (Cells(i + 1, 1))
            ticker = (ws.Cells(i + 1, 1))
            ws.Cells(tickerrow, "I").Value = ticker
            
            'Sumif to get annual change
            opendate = Application.WorksheetFunction.MinIfs(ws.Range("B:B"), ws.Range("A:A"), ws.Cells(tickerrow, "I"))
            closedate = Application.WorksheetFunction.MaxIfs(ws.Range("B:B"), ws.Range("A:A"), ws.Cells(tickerrow, "I"))
            'MsgBox (opendate)
            'MsgBox (closedate)
            
            openprice = Application.WorksheetFunction.SumIfs(ws.Range("C:C"), ws.Range("A:A"), ws.Cells(tickerrow, "I"), ws.Range("B:B"), opendate)
            'MsgBox (openprice)
            closeprice = Application.WorksheetFunction.SumIfs(ws.Range("F:F"), ws.Range("A:A"), ws.Cells(tickerrow, "I"), ws.Range("B:B"), closedate)
            'MsgBox (closeprice)
            
            pricechg = closeprice - openprice
            ws.Cells(tickerrow, "J").Value = pricechg
            
            'Conditional formatting
            If pricechg > 0 Then
               ws.Cells(tickerrow, "J").Interior.ColorIndex = 4
               
            ElseIf pricechg < 0 Then
               ws.Cells(tickerrow, "J").Interior.ColorIndex = 3
            
            Else
                ws.Cells(tickerrow, "J").Interior.ColorIndex = 48
            End If
            
            'percent change
            ws.Cells(tickerrow, "K").Value = pricechg / openprice
            ws.Cells(tickerrow, "K").NumberFormat = "0.00%"
                   
            'Sumif to get total volume
            ws.Cells(tickerrow, "L").Value = Application.WorksheetFunction.SumIfs(ws.Range("G:G"), ws.Range("A:A"), ws.Cells(tickerrow, "I"))
            
            
            tickerrow = tickerrow + 1
                         
        End If
        'Exit For
    Next i
    
    tickerrow = 2
    Summary_lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    val_increase = 0
    val_decrease = 0
    greatest_volume = 0

    
    For j = 2 To (Summary_lastrow - 1)
        If ws.Cells(j, "K").Value > val_increase Then
            val_increase = ws.Cells(j, "K").Value
            ticker_increase = ws.Cells(j, "I").Value
            
        ElseIf ws.Cells(j, "K").Value < val_decrease Then
            val_decrease = ws.Cells(j, "K").Value
            ticker_decrease = ws.Cells(j, "I").Value
        End If
    Next j
    
    ws.Cells(2, "O").Value = ticker_increase
    ws.Cells(2, "P").Value = val_increase
    ws.Cells(2, "P").NumberFormat = "0.00%"
    ws.Cells(3, "O").Value = ticker_decrease
    ws.Cells(3, "P").Value = val_decrease
    ws.Cells(3, "P").NumberFormat = "0.00%"
        
   For k = 2 To (Summary_lastrow - 1)
        If ws.Cells(k, "L").Value > greatest_volume Then
            greatest_volume = ws.Cells(k, "L").Value
            ticker_volume = ws.Cells(k, "I").Value
            
        End If
    Next k
    
    ws.Cells(4, "P").Value = greatest_volume
    ws.Cells(4, "O").Value = ticker_volume
        
   ws.Columns("I:P").AutoFit

  
        

'Exit For

Next ws


End Sub
