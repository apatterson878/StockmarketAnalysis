Sub stock()
    
    'Easy variables
    Dim LRow As Long
    Dim LCol As Long

    Dim ticker As String
    Dim tickerticker As String
    Dim volume As Double
    Dim tickertow As Integer
    tickerrow = 2
    
    
    Dim bestchange As Double, biggestvolume As Double, worstchange As Double
    Dim tickerone As String, tickertwo As String, tickerthree As String
    
    
    'Moderate variables
    Dim stockdate As String, stockyear As String, monthday As String
    Dim openprice As Double, closeprice As Double
    Dim yearlychange As Double, percentchange As Double, ogpricerow As Integer
    ogpricerow = 2

    volume = 0
    'Find the last non-blank cell in column A(1)
    LRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Find the last non-blank cell in row 1
    LCol = Cells(1, Columns.Count).End(xlToLeft).Column

    'monthday = Right(Cells(2, 2), 4)
    openprice = Cells(2, 3)

    For Row = 2 To LRow
    ticker = Cells(Row, 1)
    'Tickerticker is being set to the next ticker symbol to be compared in IF
    tickerticker = Cells((Row + 1), 1)

    stockyear = Left(Cells(Row, 2), 4)
    stockdate = Left(Cells((Row + 1), 2), 4)
    closeprice = Cells((Row + 1), 6)
    
    If ticker = tickerticker Then
    volume = (Cells(Row, 7) + volume)
    
    
    If stockdate > stockyear Then
    closeprice = Cells((Row), 6)
    End If
    
    If openprice <> 0 Then
    percentchange = ((closeprice / openprice) - 1)
    Cells(tickerrow, 11) = FormatPercent(percentchange, 1)
    Else
    'Post an error when diving by zero
    Cells(tickerrow, 11) = "Error"
    End If
    
    
    'percentchange = ((closeprice / openprice) - 1)
    yearlychange = (closeprice - openprice)
    ' Test
    
    If yearlychange > 0 Then
    
    Cells(tickerrow, 12).Interior.ColorIndex = 4
    Else
    Cells(tickerrow, 12).Interior.ColorIndex = 3
    End If
    

    
    'Cells(tickerrow, 11) = FormatPercent(percentchange, 1)
    Cells(tickerrow, 12) = yearlychange
    Cells(tickerrow, 9) = ticker
    Cells(tickerrow, 10) = volume
            
        
        ElseIf ticker <> tickerticker Then
            openprice = Cells((Row + 1), 3)
            Cells(tickerrow, 10) = (Cells(Row, 7) + volume)
            tickerrow = (tickerrow + 1)
            volume = 0
        
    End If

    Next Row
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Total Stock Volume"
    Cells(1, 11) = "% Change"
    Cells(1, 12) = "Yearly Change"

'Hard part

LRow = Cells(Rows.Count, 10).End(xlUp).Row
biggestvolume = Cells(2, 10)
bestchange = Cells(2, 11)
bestyearlychange = Cells(2, 12)
tickerone = Cells(2, 9)
tickertwo = Cells(2, 9)
tickerthree = Cells(2, 9)


For i = 2 To LRow
volume = Cells(i, 10)
ticker = Cells(i, 9)
If Cells(i, 11) <> "Error" Then
percentchange = Cells(i, 11)
End If
yearlychange = Cells(i, 12)

    If biggestvolume < volume Then
    biggestvolume = volume
    tickerone = ticker
    End If
    
    If worstchange > percentchange Then
    worstchange = percentchange
    tickertwo = ticker
    End If
    
    If bestchange < percentchange Then
    bestchange = percentchange
    tickerthree = ticker
    End If

    Next i

Cells(1, 17) = "Ticker"
Cells(1, 18) = "Value"

Cells(2, 16) = "Largest Volume"
Cells(3, 16) = "Greatest % increase"
Cells(4, 16) = "Greatest % decrease"

Cells(2, 18) = biggestvolume
Cells(3, 18) = FormatPercent(bestchange, 1)
Cells(4, 18) = FormatPercent(worstchange, 1)

Cells(2, 17) = tickerone
Cells(3, 17) = tickerthree
Cells(4, 17) = tickertwo

End Sub