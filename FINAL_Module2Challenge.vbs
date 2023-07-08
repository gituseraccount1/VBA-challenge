Attribute VB_Name = "Module1"
Sub tickerList()

' Holds the ticker symbol name
Dim tickerSymbolName As String

' Goes to a new line to add a new ticker symbol name
Dim listOfTickers As String

' Holds the total stock volume
Dim tickerTotal As Double

' Counts the how many rows are in a ticker in order to get the row of the opening price (which I named first day price) for each ticker symbol
Dim firstDayPriceRow As Double

' Gets the row of where the closing price (which I named last Day Price) is located
Dim lastDayPrice As Double

' stores the actual value of the opening price (which I called first day price)
Dim firstDayPrice As Double

' loops through each worksheet in the excel workbook
For Each ws In Worksheets

    ' Counts the number of rows
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Sets the title column for Ticker, Yearly Change, Percent Change, and Total Stock Volume
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'set this to 0
    firstDayPriceRow = 0
    
    'set this to 0
    tickerTotal = 0
    
    ' set this to zero, so the ticker list with other info. can start from row 2 instead of 1
    listOfTickers = 2

    ' loops through the <ticker> column
    For i = 2 To lastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'firstDayPrice = ws.Cells(i - firstDayPriceRow, 3).Value
            
            tickerSymbolName = ws.Cells(i, 1).Value
            tickerTotal = tickerTotal + ws.Cells(i, 7).Value
            lastDayPrice = ws.Cells(i, 6).Value
            firstDayPrice = ws.Cells(i - firstDayPriceRow, 3).Value
            
            ws.Range("I" & listOfTickers).Value = tickerSymbolName
            ws.Range("L" & listOfTickers).Value = tickerTotal
            ws.Range("J" & listOfTickers) = (firstDayPrice - lastDayPrice) * -1
            ws.Range("K" & listOfTickers).Value = FormatPercent((((firstDayPrice - lastDayPrice) / lastDayPrice) / 100) * -1)

            listOfTickers = listOfTickers + 1
            tickerTotal = 0
            firstDayPriceRow = 0
                
        Else
            tickerTotal = tickerTotal + ws.Cells(i, 7).Value
            firstDayPriceRow = firstDayPriceRow + 1
        
        End If
    
    
    Next i

Next ws

End Sub


Sub conditionalFormatting()

For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    For i = 2 To lastRow
    
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        
        If ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
    
        
    Next i
    
Next ws


End Sub


Sub summaryTable()

For Each ws In Worksheets

    Dim storeTicker As String
    Dim storeLargestNum As Double
    storeLargestNum = 0
    
    Dim storeTickerGreatestDec As String
    Dim storeGreatestDec As Double
    storeGreatestDec = 0
    
    Dim storeTickerGreatestInc As String
    Dim storeGreatestInc As Double
    storeGreatestInc = 0
    
    
    
        lastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        For i = 2 To lastRow
        
             If ws.Cells(i, 11).Value > storeGreatestInc Then
                storeGreatestInc = ws.Cells(i, 11).Value
                storeTickerGreatestInc = ws.Cells(i, 9).Value
                
                ws.Cells(2, 16).Value = storeTickerGreatestInc
                ws.Cells(2, 17).Value = FormatPercent(storeGreatestInc)
                
            End If
            
            
            If ws.Cells(i, 11).Value < storeGreatestDec Then
                storeGreatestDec = ws.Cells(i, 11).Value
                storeTickerGreatestDec = ws.Cells(i, 9).Value
                
                
                ws.Cells(3, 16).Value = storeTickerGreatestDec
                ws.Cells(3, 17).Value = FormatPercent(storeGreatestDec)
                
            End If
            
            
            If ws.Cells(i, 12).Value > storeLargestNum Then
                storeLargestNum = ws.Cells(i, 12).Value
                storeTicker = ws.Cells(i, 9).Value
                
                ws.Cells(4, 17).Value = storeLargestNum
                ws.Cells(4, 16).Value = storeTicker
            End If
            
        Next i

Next ws


End Sub
