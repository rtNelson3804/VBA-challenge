Attribute VB_Name = "Module1"
Sub stockData()

    Application.ScreenUpdating = False

    Dim ticker, tickerMaxVolume, tickerMaxPerc, tickerMinPerc As String
    Dim i, j As Integer
    Dim maxVolume, openPrice, closePrice, volume As Double
    Dim maxPerc, minPerc As Single
    Dim ws As Worksheet

    ' Initialize global variables

    j = 2
    globalMaxVolume = 0
    globalMaxPerc = 0
    globalMinPerc = 1 ' Set to a high number since percentages are usually between 0 and 1

    ' Process each worksheet
    For Each ws In Worksheets
        ' Fully qualify all references with ws to avoid errors
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Define last row
        volume = 0 ' Opening Value for volume
        j = 2 ' opening value for the rows to display output
        ticker = ws.Cells(2, 1).Value ' Setting the initial ticker
        maxVolume = 0 ' Reset sheet-specific max volume
        maxPerc = 0
        minPerc = 1 ' Reset to high value for sheet

        ' Loop through rows in the worksheet
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = ticker And ws.Cells(i - 1, 1).Value <> ticker Then
                openPrice = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 1).Value = ticker Then
                volume = volume + ws.Cells(i, 7).Value
                closePrice = ws.Cells(i, 6).Value
            ElseIf ws.Cells(i, 1).Value <> ticker Or i = lastRow Then
                If closePrice - openPrice <> 0 Then
                    Dim percChange As Single
                    percChange = ((closePrice - openPrice) / openPrice)
                Else
                    percChange = 0
                End If
                
            
                    ws.Cells(j, 9).Value = ticker
                    ws.Cells(j, 10).Value = (openPrice - closePrice) * -1
                    ws.Cells(j, 10).NumberFormat = "0.00"
                    If ws.Cells(j, 10).Value > 0 Then
                        ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0) ' Green > 0
                    ElseIf Cells(j, 10).Value < 0 Then
                        ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0) ' Red < 0
                    End If
                    ws.Cells(j, 11).Value = percChange
                    ws.Cells(j, 11).NumberFormat = "0.00%"
                    If Cells(j, 11).Value > 0 Then
                        ws.Cells(j, 11).Interior.Color = RGB(0, 255, 0) ' Green > 0
                    ElseIf Cells(j, 11).Value < 0 Then
                        ws.Cells(j, 11).Interior.Color = RGB(255, 0, 0) ' Red for < 0
                    End If
                    ws.Cells(j, 12).Value = volume
                    ws.Cells(j, 12).NumberFormat = "0"
                    
                    j = j + 1
                    
                    

                ' Update sheet-specific max/min and volume
                If maxVolume < volume Then
                    maxVolume = volume
                    tickerMaxVolume = ticker
                End If

                If maxPerc < percChange Then
                    maxPerc = percChange
                    tickerMaxPerc = ticker
                End If

                If minPerc > percChange Then
                    minPerc = percChange
                    tickerMinPerc = ticker
                End If

                ' Reset for the next ticker
                volume = 0
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
            End If

        Next i
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = tickerMaxPerc
        ws.Cells(2, 17).Value = maxPerc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).Interior.Color = RGB(0, 255, 0) ' Green > 0
        ws.Cells(3, 16).Value = tickerMinPerc
        ws.Cells(3, 17).Value = minPerc
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).Interior.Color = RGB(255, 0, 0) ' Red < 0
        ws.Cells(4, 16).Value = tickerMaxVolume
        ws.Cells(4, 17).Value = maxVolume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        
        ws.Columns.AutoFit
        
        

    Next ws

    Application.ScreenUpdating = True

End Sub



