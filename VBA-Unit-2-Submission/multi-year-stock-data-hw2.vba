Sub StockData()

    Dim currentWS As Worksheet
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim maxLimit As Long

    'Loop through all sheets in workbook
    For Each currentWS In Worksheets
        
        'Add calculated headers to every sheet
        currentWS.Cells(1, 9) = "Ticker"
        currentWS.Cells(1, 10) = "Total Volume"
        currentWS.Cells(1, 11) = "Yearly Change"
        currentWS.Cells(1, 12) = "Percent Change"
        currentWS.Cells(2, 15) = "Greatest % increase"
        currentWS.Cells(3, 15) = "Greatest % decrease"
        currentWS.Cells(4, 15) = "Greatest Total Volume"
        currentWS.Cells(1, 16) = "Ticker"
        currentWS.Cells(1, 17) = "Value"

        'To find the last position of the data cell of the row
        maxLimit = currentWS.Cells(Rows.Count, 1).End(xlUp).Row

        'Set an Integer for Loop
        TotalVolume = 0
        j = 1
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        greatestTotalVolume = 0
            
        OpenValue = currentWS.Cells(2, 3).Value

        'Loop
        For i = 2 To maxLimit

            'Calculate the Total Volume
            If currentWS.Cells(i, 1).Value = currentWS.Cells(i + 1, 1).Value Then
                TotalVolume = TotalVolume + currentWS.Cells(i, 7).Value
            Else
                TotalVolume = TotalVolume + currentWS.Cells(i, 7).Value
                currentWS.Cells(j + 1, 9).Value = currentWS.Cells(i, 1).Value
                currentWS.Cells(j + 1, 10).Value = TotalVolume
                
                'Check greatest volume from the total Volume
                If TotalVolume >= greatestTotalVolume Then
                    greatestTotalVolume = TotalVolume
                    greatestTotalVolumeTicker = currentWS.Cells(j + 1, 9).Value
                End If

                TotalVolume = 0
            End If


            If currentWS.Cells(i + 1, 1).Value <> currentWS.Cells(i, 1).Value Then
                CloseValue = currentWS.Cells(i, 6).Value
            
            'Calculate the yearly change
                YearlyChange = CloseValue - OpenValue
                currentWS.Cells(j + 1, 11).Value = YearlyChange
                
                'Cpnditional Formatting
                If currentWS.Cells(j + 1, 11).Value < 0 Then

                'Red color for Negative Values
                currentWS.Cells(j + 1, 11).Interior.ColorIndex = 3
            Else
                'Green color for Positive Values
                currentWS.Cells(j + 1, 11).Interior.ColorIndex = 4

            End If

            'Calculate the percent change
            If OpenValue <> 0 Then
                PercentChange = YearlyChange / OpenValue
            Else
                PercentChange = YearlyChange
            End If
                    
            OpenValue = currentWS.Cells(i + 1, 3).Value
            currentWS.Cells(j + 1, 12).Value = PercentChange
            currentWS.Cells(j + 1, 12).NumberFormat = "0.00%"
            j = j + 1

            'Check the greatest percent increase from the percent change
            If PercentChange >= greatestPercentIncrease Then
                greatestPercentIncrease = PercentChange
                greatestPercentIncreaseTicker = currentWS.Cells(i, 1).Value
            ElseIf PercentChange < greatestPercentDecrease Then
                greatestPercentDecrease = PercentChange
                greatestPercentDecreaseTicker = currentWS.Cells(i, 1).Value
            End If
        End If

        Next i
        
        'Set the values to the sheet
        currentWS.Cells(2, 17).Value = greatestPercentIncrease
        currentWS.Cells(2, 17).NumberFormat = "0.00%"
        currentWS.Cells(2, 16).Value = greatestPercentIncreaseTicker
            
        currentWS.Cells(3, 17).Value = greatestPercentDecrease
        currentWS.Cells(3, 17).NumberFormat = "0.00%"
        currentWS.Cells(3, 16).Value = greatestPercentDecreaseTicker
        
        currentWS.Cells(4, 17).Value = greatestTotalVolume
        currentWS.Cells(4, 16).Value = greatestTotalVolumeTicker
    Next
    
End Sub
