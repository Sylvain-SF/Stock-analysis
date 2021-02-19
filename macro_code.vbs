Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    'Choose which year we want to perform analysis on
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks(" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    'Activate data worksheet
     Worksheets(yearValue).Activate

    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'Initialyze variables
    Dim TickersVolumes(12) As Long
    Dim TickersStartingPrices(12) As Single
    Dim TickersEndingPrices(12) As Single

    'Loop over the index to set every volume to 0
    For i = 0 To 11
         TickersVolumes(i) = 0
    Next i




    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

        'Get stock index in tickers
        MyTickersIndex = Application.Match(Cells(i, 1).Value, tickers, False) - 1
        TickersVolumes(MyTickersIndex) = TickersVolumes(MyTickersIndex) + Cells(i, 8).Value

        'Check if previous line cells is different
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            TickersStartingPrices(MyTickersIndex) = Cells(i, 6).Value
        End If

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            TickersEndingPrices(MyTickersIndex) = Cells(i, 6).Value
        End If
    Next i

    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11

       Worksheets("All Stocks Analysis").Activate
       Cells(i + 4, 2).Value = TickersVolumes(i)
       Cells(i + 4, 1).Value = tickers(i)
       Cells(i + 4, 3).Value = (TickersEndingPrices(i) - TickersStartingPrices(i)) / TickersStartingPrices(i)
    Next i

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            Cells(i, 3).Interior.Color = vbGreen

        Else

            Cells(i, 3).Interior.Color = vbRed

        End If

    Next i

    'Finish and display timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
Sub ClearWorksheet()

Cells.Clear

End Sub

Sub yearValueAnalysis()

yearValue = InputBox("What year would you like to run the analysis on?")
End Sub
