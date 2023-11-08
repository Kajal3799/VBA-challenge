Sub MyWork()
    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
        
        ' Initialize variables for stock data
        Dim ticker As String
        ticker = ""
        Dim startdate As Long
        Dim openvalue As Double
        Dim closevalue As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim volume As LongLong
        volume = 0
        Dim tickercount As Integer
        tickercount = 1
        Dim i As Long
        
        ' Find the last row with data in column 1 (ticker column)
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through the data rows in the worksheet
        For i = 2 To lastrow
            ' Check if the current row's ticker matches the previous one
            If ws.Cells(i, 1).Value = ticker Then
                ' Accumulate the stock's volume
                volume = volume + ws.Cells(i, 7).Value
            Else
                ' A new stock is encountered, update variables and ticker count
                tickercount = tickercount + 1
                ticker = ws.Cells(i, 1).Value
                ws.Cells(tickercount, 9).Value = ticker
                volume = 0 + ws.Cells(i, 7).Value
                startdate = ws.Cells(i, 2).Value
                openvalue = ws.Cells(i, 3).Value
            End If
            
            ' Check if the next row's ticker is different (end of the stock data)
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ' Calculate yearly change and percent change
                closevalue = ws.Cells(i, 6).Value
                yearlychange = closevalue - openvalue
                percentchange = yearlychange / openvalue
                percentchange = Application.WorksheetFunction.Round(percentchange, 4)
                
                ' Write the calculated values to the worksheet
                ws.Cells(tickercount, 10).Value = yearlychange
                ws.Cells(tickercount, 11).Value = percentchange
                ws.Cells(tickercount, 12).Value = volume
                
                ' Apply conditional formatting based on yearly change
                If yearlychange < 0 Then
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 3 ' Red
                ElseIf yearlychange > 0 Then
                    ws.Cells(tickercount, 10).Interior.ColorIndex = 4 ' Green
                End If
            End If
        Next i
        
        ' Format columns K and L as percentages and whole numbers
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "0"
        
        ' Format header row with bold text
        With ws.Range("I1:L1")
            .NumberFormat = "Text"
            .Font.Bold = True
        End With
        
        ' Add headers to the columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Auto-fit columns I to L
        ws.Columns("I:L").AutoFit
        
        ' Initialize variables for tracking the greatest increase, decrease, and volume
        Dim inc As Double
        inc = 0
        Dim dec As Double
        dec = 0
        Dim maxVol As LongLong
        maxVol = 0
        Dim incTic As String
        Dim decTic As String
        Dim maxVolTic As String
        
        ' Find the last row with data in column 9 (ticker column)
        dataEnd = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Loop through the data to find the greatest values
        For i = 2 To dataEnd
            If inc < ws.Cells(i, 11).Value Then
                inc = ws.Cells(i, 11).Value
                incTic = ws.Cells(i, 9).Value
            End If
            If dec > ws.Cells(i, 11).Value Then
                dec = ws.Cells(i, 11).Value
                decTic = ws.Cells(i, 9).Value
            End If
            If maxVol < ws.Cells(i, 12).Value Then
                maxVol = ws.Cells(i, 12).Value
                maxVolTic = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' Write the greatest increase, decrease, and volume to specific cells
        ws.Range("P2").Value = incTic
        ws.Range("Q2").Value = inc
        ws.Range("P3").Value = decTic
        ws.Range("Q3").Value = dec
        ws.Range("P4").Value = maxVolTic
        ws.Range("Q4").Value = maxVol
        
        ' Format percentage values in cells Q2 and Q3, and volume in cell Q4
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0,0000"
        
        ' Add labels for the greatest values
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Set headers for columns P and Q and make them bold
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O1:Q4").Font.Bold = True
        
        ' Auto-fit columns O to Q
        ws.Columns("O:Q").AutoFit
    Next
End Sub