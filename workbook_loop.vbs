
Sub Q_create_summary_and_get_greatest()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim quarterStart As Double
    Dim quarterEnd As Double
    Dim quarterlyChange As Double
    Dim percentChange As String
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim tickerRange As Range
    Dim changeRange As Range
    Dim volumeRange As Range
    Dim maxPercentage As Double
    Dim minPercentage As Double
    Dim maxVolume As Double
    Dim maxPercentageTicker As String
    Dim minPercentageTicker As String
    Dim maxVolumeTicker As String

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the worksheet name starts with "Q"
        If Left(ws.Name, 1) = "Q" Then
            ' Find the last used row in column A
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

      ' Assign summary table headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

     ' Define variables for tracking ticker, quarter start, and total stock volume
            ticker = ""
            summaryRow = 2 ' Start writing data from row 2

    ' Define the ranges for columns I (Ticker), K (Percentage Change), and L (Total Volume)
            Set tickerRange = ws.Range("I2:I" & lastRow)
            Set changeRange = ws.Range("K2:K" & lastRow)
            Set volumeRange = ws.Range("L2:L" & lastRow)

     ' Initialize values based on the first row data
            maxPercentage = changeRange.Cells(1).Value
            minPercentage = changeRange.Cells(1).Value
            maxVolume = volumeRange.Cells(1).Value

    ' Loop through the data
    For i = 2 To lastRow
                If ws.Cells(i, 1).Value <> ticker Then
                    ' New ticker, update ticker, quarterStart, and reset totalVolume
                    ticker = ws.Cells(i, 1).Value
                    quarterStart = ws.Cells(i, 3).Value
                    totalVolume = 0 ' Reset totalVolume for the new ticker
                End If

        ' Update quarterEnd and totalVolume for each row
                quarterEnd = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value ' Accumulate total volume

    ' Calculate quarterly change and percentage change when quarterEnd is available
                If quarterEnd <> 0 Then
                    quarterlyChange = quarterEnd - quarterStart
                    If quarterStart <> 0 Then
                        percentChange = (quarterlyChange / quarterStart) * 100
                    Else
                        percentChange = 0
                    End If

                    ' Format percentChange as a string with % symbol
                    percentChange = Format(percentChange, "0.00") & "%"

                    ' Write results to summary table for each new ticker or at the end of data
                    If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                        ws.Cells(summaryRow, 9).Value = ticker
                        ws.Cells(summaryRow, 10).Value = quarterlyChange
                        ws.Cells(summaryRow, 11).Value = percentChange
                        ws.Cells(summaryRow, 12).Value = totalVolume

                        ' Move to the next row in the summary table
                        summaryRow = summaryRow + 1

                        ' Reset quarterStart for the next ticker
                        quarterStart = 0
                    End If
                End If
            Next i

            ' Apply conditional formatting based on quarterly change values
            lastRow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
            Set rng = ws.Range("J2:J" & lastRow)

            ' Clear existing conditional formatting
            rng.FormatConditions.Delete

                ' Apply conditional formatting based on quarterly change values
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value < 0 Then
                cell.Interior.ColorIndex = 3 ' Red
            ElseIf cell.Value > 0 Then
                cell.Interior.ColorIndex = 4 ' Green
            Else
                cell.Interior.ColorIndex = 0 ' No fill
            End If
        End If
Next cell

            ' Find the last row with data
            lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

            ' Initialize values based on the first row data
            maxPercentage = changeRange.Cells(1).Value
            minPercentage = changeRange.Cells(1).Value
            maxVolume = volumeRange.Cells(1).Value

            ' Loop through the data
            For i = 1 To lastRow
                Dim currentPercentage As Double
                Dim currentVolume As Double

                ' Get the percentage change and volume for the current row
                currentPercentage = changeRange.Cells(i).Value
                currentVolume = volumeRange.Cells(i).Value

                ' Update highest and lowest percentages
                If currentPercentage > maxPercentage Then
                    maxPercentage = currentPercentage
                    maxPercentageTicker = tickerRange.Cells(i).Value
                ElseIf currentPercentage < minPercentage Then
                    minPercentage = currentPercentage
                    minPercentageTicker = tickerRange.Cells(i).Value
                End If

                ' Update maximum volume
                If currentVolume > maxVolume Then
                    maxVolume = currentVolume
                    maxVolumeTicker = tickerRange.Cells(i).Value
                End If
            Next i

            ' Write the results to the specified cells
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("P2").Value = maxPercentageTicker
            ws.Range("Q2").NumberFormat = "0.00%" 
            ws.Range("Q2").Value = maxPercentage

            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("P3").Value = minPercentageTicker
            ws.Range("Q3").NumberFormat = "0.00%" 
            ws.Range("Q3").Value = minPercentage

            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P4").Value = maxVolumeTicker
            ws.Range("Q4").Value = maxVolume
        End If
    Next ws
End Sub