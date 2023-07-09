Sub MultiYearStockAnalysis()

'looping multiple worksheets
For Each ws In Worksheets

    'Setting headers for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
     'Setting headers for summary table 2
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

    'Declaring Variables for data
        Dim ticker As String
        Dim totalStock As Double
            totalStock = 0
        Dim yearlyChange As Double
            yearlyChange = 0
        Dim percentChange As Double
        'Declaring yearly starting price
        Dim openPrice As Long
            openPrice = 2
        Dim YearlyOpenPrice As Double
            YearlyOpenPrice = ws.Cells(openPrice, 3).Value
    'Declaring starting row for output values
        Dim summary As Integer
            summary = 2
        
    'For loopiing through lastRow in the dataset
        Dim lastRow As Long
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Looping through the worksheet
        Dim i As Long
        For i = 2 To lastRow
        
            'Pulling individual ticker values if current ticker is not same as perious ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Ticker value
                ticker = ws.Cells(i, 1).Value
                'Adding stock volume
                totalStock = totalStock + ws.Cells(i, 7).Value
                'Calculating yearly price change
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = (closePrice - YearlyOpenPrice)
                'Calculating yearly percentage change
                percentChange = (yearlyChange / YearlyOpenPrice)
                
            'Establising Values in summary table
            ws.Range("I" & summary).Value = ticker
            ws.Range("L" & summary).Value = totalStock
            ws.Range("J" & summary).Value = yearlyChange
            ws.Range("K" & summary).Value = percentChange
            ws.Range("K" & summary).NumberFormat = "0.00%"
            
            'Resetting values for next row
            summary = summary + 1
            
            'Resetting values
            totalStock = 0
            yearlyChange = 0
            percentChange = 0
            YearlyOpenPrice = ws.Cells(i + 1, 3).Value
            
            Else
                'Adding total stock volume
                totalStock = totalStock + ws.Cells(i, 7).Value
            End If

        Next i
        
            'Declaring variable for summary table 2
            Dim greatestIncrease As Double
                greatestIncrease = ws.Cells(2, 11).Value
            Dim greatestDecrease As Double
                greatestDecrease = ws.Cells(2, 11).Value
            Dim highestValue As Double
                highestValue = ws.Cells(2, 12).Value
    
            'Calculating values for summary table 2
            For i = 2 To lastRow
            
                'For obtaining Greatest % Increase
                If ws.Cells(i, 11).Value > greatestIncrease Then
                    greatestIncrease = ws.Cells(i, 11).Value
                    ws.Range("Q2").Value = greatestIncrease
                    ws.Range("Q2").NumberFormat = "0.00%"
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                End If
                
                'For obtaining Greatest % Decrease
                If ws.Cells(i, 11).Value < greatestDecrease Then
                    greatestDecrease = ws.Cells(i, 11).Value
                    ws.Range("Q3").Value = greatestDecrease
                    ws.Range("Q3").NumberFormat = "0.00%"
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                End If
                
                'For obtaining Greatest Total stock Volume
                If ws.Cells(i, 12).Value > highestValue Then
                    highestValue = ws.Cells(i, 12).Value
                    ws.Range("Q4").Value = highestValue
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
                
                    'Formatting positive yearly change in green and negative yearly change in red
                    If ws.Cells(i, 10).Value >= 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                    End If

            Next i
        'For better visualization
        ws.Columns("I:Q").AutoFit

Next ws
    
End Sub
