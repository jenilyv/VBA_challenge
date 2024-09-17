Sub stock_analysis()

    ' Set dimensions
    Dim ws1 As Worksheet
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double

    ' Loop through each worksheet in the workbook
    For Each ws1 In ThisWorkbook.Worksheets
        ' You don't need to activate the sheet, just refer to it directly.
        
        ' Set title row in the current worksheet
        ws1.Range("I1").Value = "Ticker"
        ws1.Range("J1").Value = "Quarterly Change"
        ws1.Range("K1").Value = "Percent Change"
        ws1.Range("L1").Value = "Total Stock Volume"
        ws1.Range("P1").Value = "Ticker"
        ws1.Range("Q1").Value = "Value"
        ws1.Range("O2").Value = "Greatest % Increase"
        ws1.Range("O3").Value = "Greatest % Decrease"
        ws1.Range("O4").Value = "Greatest Total Volume"

        ' Set initial values
        j = 0
        total = 0
        change = 0
        start = 2

        ' Get the row number of the last row with data
        rowCount = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker changes then print results
            If ws1.Cells(i + 1, 1).Value <> ws1.Cells(i, 1).Value Then

                ' Stores results in variables
                total = total + ws1.Cells(i, 7).Value

                ' Handle zero total volume
                If total = 0 Then
                    ' Print the results
                    ws1.Range("I" & 2 + j).Value = ws1.Cells(i, 1).Value
                    ws1.Range("J" & 2 + j).Value = 0
                    ws1.Range("K" & 2 + j).Value = "%" & 0
                    ws1.Range("L" & 2 + j).Value = 0

                Else
                    ' Find first non-zero starting value
                    If ws1.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws1.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    ' Calculate change
                    change = (ws1.Cells(i, 6) - ws1.Cells(start, 3))
                    percentChange = change / ws1.Cells(start, 3)

                    ' Start of the next stock ticker
                    start = i + 1

                    ' Print the results
                    ws1.Range("I" & 2 + j).Value = ws1.Cells(i, 1).Value
                    ws1.Range("J" & 2 + j).Value = change
                    ws1.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws1.Range("K" & 2 + j).Value = percentChange
                    ws1.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws1.Range("L" & 2 + j).Value = total

                    ' Colors positives green and negatives red
                    Select Case change
                        Case Is > 0
                            ws1.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws1.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws1.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If

                ' Reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0

            ' If ticker is still the same add results
            Else
                total = total + ws1.Cells(i, 7).Value
            End If

        Next i

        ' Take the max and min and place them in a separate part in the worksheet
        ws1.Range("Q2").Value = "%" & WorksheetFunction.Max(ws1.Range("K2:K" & rowCount)) * 100
        ws1.Range("Q3").Value = "%" & WorksheetFunction.Min(ws1.Range("K2:K" & rowCount)) * 100
        ws1.Range("Q4").Value = WorksheetFunction.Max(ws1.Range("L2:L" & rowCount))

        ' Returns one less because header row is not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws1.Range("K2:K" & rowCount)), ws1.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws1.Range("K2:K" & rowCount)), ws1.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws1.Range("L2:L" & rowCount)), ws1.Range("L2:L" & rowCount), 0)

        ' Final ticker symbol for total, greatest % of increase and decrease, and average
        ws1.Range("P2").Value = ws1.Cells(increase_number + 1, 9).Value
        ws1.Range("P3").Value = ws1.Cells(decrease_number + 1, 9).Value
        ws1.Range("P4").Value = ws1.Cells(volume_number + 1, 9).Value

    Next ws1

End Sub
