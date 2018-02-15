Sub flipBook()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Call analyze
        Call label
        Call fit
    Next
End Sub


Sub analyze()

'count rows of sheet
    Dim numRows As Double
    numRows = ActiveSheet.UsedRange.Rows.Count


'//initialize variables; these will be reused for each ticker
'Record ticker and temp for checking if new ticker
    Dim currentTicker, tickerAtIndex As String
    
'Running total of volume
    Dim totalVolume As Double
    
'Record open and close price for the year
    Dim yearOpen, yearClose As Double
    
'Variable for saving the difference of yearOpen and yearClose
    Dim yearChange As Double
'//



'//these will only be changed when records are broken
'eg. new greatest total volume

'Record record holding values; 0 by default when initialized
    Dim greatestIncrease_Value, greatestDecrease_Value, greatestVolume_Value As Double
    
'save ticker for record setters
    Dim greatestIncrease_Ticker, greatestDecrease_Ticker, greatestVolume_Ticker As String


'count unique tickers
'set at 2 so that tickers are output to correct rows later
    Dim tickerCount As Double
    tickerCount = 2


'Get first row data for comparison
    currentTicker = Cells(2, 1).Value
    yearOpen = Cells(2, 3).Value
    totalVolume = Cells(2, 7).Value

'Loop through every row starting at the second row of data to one past last row
    For i = 3 To numRows + 1
'First save the ticker at row i
        tickerAtIndex = Cells(i, 1)
'if the ticker at i is the same as the last then add that day's volume to total
        If (tickerAtIndex = currentTicker) Then
            totalVolume = totalVolume + Cells(i, 7).Value
'//else meanse we've encountered a new ticker which is saved in tickerAtIndex

        Else
            Cells(tickerCount, 9).Value = currentTicker
        
'Get the yearClose value from row i -1 since that is the last trading day for currentTicker
            yearClose = Cells(i - 1, 6).Value
'Calculate yearChange
            yearChange = yearClose - yearOpen
'Output yearChange
            Cells(tickerCount, 10).Value = yearChange

                If (yearOpen > 0) Then
                
                    changePercent = yearChange / yearOpen
                    Cells(tickerCount, 11).Value = changePercent
                    If (changePercent > greatestIncrease_Value) Then
                        greatestIncrease_Value = changePercent
                        greatestIncrease_Ticker = currentTicker
                    ElseIf (changePercent < greatestDecrease_Value) Then
                        greatestDecrease_Value = changePercent
                        greatestDecrease_Ticker = currentTicker
                    End If
            
                End If

            Cells(tickerCount, 12).Value = totalVolume
                If (totalVolume > greatestVolume_Value) Then
                    greatestVolume_Value = totalVolume
                    greatestVolume_Ticker = currentTicker
                End If
    
        
'Set the fill color to red if yearChange is negative and green if positive
            If (yearChange < 0) Then
                Cells(tickerCount, 10).Interior.ColorIndex = 3
            ElseIf (yearChange > 0) Then
                Cells(tickerCount, 10).Interior.ColorIndex = 4
            End If
            
'Save the new ticker as currentTicker; record yearOpen and volume from row i; increment count
            currentTicker = tickerAtIndex
            yearOpen = Cells(i, 3).Value
            totalVolume = Cells(i, 7).Value
            tickerCount = tickerCount + 1
    
        End If
    Next i
    
'Output Greatest Change etc. for sheet
    
    Cells(2, 16).Value = greatestIncrease_Ticker
    Cells(3, 16).Value = greatestDecrease_Ticker
    Cells(4, 16).Value = greatestVolume_Ticker
    
    Cells(2, 17).Value = greatestIncrease_Value
    Cells(3, 17).Value = greatestDecrease_Value
    Cells(4, 17).Value = greatestVolume_Value

End Sub

'Sub for labelling stuff
Sub label()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

End Sub


'Sub for formatting
Sub fit()
    With ActiveSheet
        .Range("J:O").EntireColumn.AutoFit
        .Range("K:K").EntireColumn.NumberFormat = "0.00%"
        .Range("N:N").ColumnWidth = 6
        .Cells(2, 17).NumberFormat = "0.00%"
        .Cells(3, 17).NumberFormat = "0.00%"
        .Cells(4, 17).NumberFormat = "0"
    End With
End Sub

'Sub for resetting sheets for testing
Sub reset()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Range("I:Q").Delete
    Next
End Sub