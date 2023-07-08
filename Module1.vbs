Attribute VB_Name = "Module1"
Sub tickerList()

' Counts the number of rows
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim tickerSymbolName As String
Dim listOfTickers As String
Dim tickerTotal As Double
Dim firstDayPrice As Double
Dim lastDayPrice As Double

Dim firstDayPriceHolder As Double



tickerTotal = 0
listOfTickers = 2
For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'firstDayPriceHolder = firstDayPrice
        'MsgBox (firstDayPriceHolder)
        'MsgBox ("This is i " & i & " and this is " & firstDayPrice & " First day Price.")
        Cells(i - 252, 3).Interior.ColorIndex = 45 ' first day number, orange
        'MsgBox ("This is the number I want, " & i - 252 & ". Is this what I am getting " & firstDayPrice)
        
        tickerSymbolName = Cells(i, 1).Value
        Cells(i, 1).Interior.ColorIndex = 37 ' blue
        
        tickerTotal = tickerTotal + Cells(i, 7).Value
        
        lastDayPrice = Cells(i, 6)
        Cells(i, 6).Interior.ColorIndex = 38 ' pink
        
        Range("I" & listOfTickers).Value = tickerSymbolName
        Range("L" & listOfTickers).Value = tickerTotal
        Range("J" & listOfTickers).Value = lastDayPrice
        
        
            
            'For j = 1 To lastRow
             '   If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
              '      firstDayPrice = Cells(j + 1, 3)
               '     Cells(j + 1, 3).Interior.ColorIndex = 50
                'Else
                    'firstDayPrice = firstDayPrice
                    'MsgBox ("Else " & firstDayPrice)
             '   End If
            'Next j
            
            
            listOfTickers = listOfTickers + 1
            tickerTotal = 0
            
    Else
        tickerTotal = tickerTotal + Cells(i, 7).Value
        Cells(i, 7).Interior.ColorIndex = 43 ' green
        firstDayPrice = firstDayPrice + 1
        'MsgBox (firstDayPrice)
    
    End If


Next i

MsgBox ("Done")

End Sub


Sub greatestVolume()

lastRow = Cells(Rows.Count, 12).End(xlUp).Row

Dim storeTicker As String
Dim storeLargestNum As Double
storeLargestNum = 0

Dim storeTickerGreatestDec As String
Dim storeGreatestDec As Double
storeGreatestDec = 10

Dim storeTickerGreatestInc As String
Dim storeGreatestInc As Double
storeGreatestInc = 0

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For i = 2 To lastRow

    If Cells(i, 12).Value > storeLargestNum Then
        storeLargestNum = Cells(i, 12).Value
        storeTicker = Cells(i, 9).Value
        
        Cells(4, 17).Value = storeLargestNum
        Cells(4, 16).Value = storeTicker
    End If
    
    
    If Cells(i, 10).Value < storeGreatestDec Then
        storeGreatestDec = Cells(i, 10).Value
        storeTickerGreatestDec = Cells(i, 9).Value
        
        Cells(3, 17).Value = storeGreatestDec
        Cells(3, 16).Value = storeTickerGreatestDec
    End If
    
    
     If Cells(i, 10).Value > storeGreatestInc Then
        storeGreatestInc = Cells(i, 10).Value
        storeTickerGreatestInc = Cells(i, 9).Value
        
        Cells(2, 16).Value = storeTickerGreatestInc
        'Range("").ClearFormats
        Cells(2, 17).Value = storeGreatestInc
        
    End If
    

Next i





End Sub


Sub test()





End Sub

'Sub tickerList()
'
'' Counts the number of rows
'lastRow = Cells(Rows.Count, 1).End(xlUp).Row
'
'Dim tickerSymbolName As String
'Dim listOfTickers As String
'Dim tickerTotal As Double
'Dim firstDayPrice As Double
'Dim lastDayPrice As Double
'
'
' For Each ws In Worksheets
'
'ws.Cells(1, 9).Value = "Ticker"
'ws.Cells(1, 10).Value = "Yearly Change"
'ws.Cells(1, 11).Value = "Percent Change"
'ws.Cells(1, 12).Value = "Total Stock Volume"
'
'tickerTotal = 0
'listOfTickers = 2
'For i = 2 To lastRow
'
 '   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 '
  '      tickerSymbolName = ws.Cells(i, 1).Value
   '
    '    tickerTotal = tickerTotal + ws.Cells(i, 7).Value
     '
      '  lastDayPrice = ws.Cells(i, 6)
       '
        'ws.Range("I" & listOfTickers).Value = tickerSymbolName
        'ws.Range("L" & listOfTickers).Value = tickerTotal
        'ws.Range("J" & listOfTickers).Value = lastDayPrice
        '
        'listOfTickers = listOfTickers + 1
        'tickerTotal = 0
        '
    'Else
     '   tickerTotal = tickerTotal + ws.Cells(i, 7).Value
    '
    '
    'End If
'
'
'Next i
'
' Next ws
'
'MsgBox ("Done")

'End Sub

'23.43 - 23.26 = 0.17 (but should show -0.17)
'C1 - F253
