Attribute VB_Name = "THE_code"
Sub stock_summary_table()
    Dim ws As Worksheet
        For Each ws In Worksheets
            ws.Activate
    Dim ticker As String
    Dim stock_volume As Double
        stock_volume = 0
    Dim summary_table_row As Double
        summary_table_row = 2
    Dim lastrow As Long
    Dim lastcol As Long
    Dim i As Double
    Dim yr_open As Double
        yr_open = 0
    Dim yr_close As Double
        yr_close = 0
    Dim percent As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
    For i = 2 To lastrow
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            yr_open = Cells(i, 3).Value
        End If
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            stock_volume = stock_volume + Cells(i, 7).Value

            yr_close = Cells(i, 6).Value
            Change = yr_close - yr_open
            percent = (yr_close - yr_open) / yr_open
            Range("I" & summary_table_row).Value = ticker
            Range("L" & summary_table_row).Value = stock_volume
            Range("J" & summary_table_row).Value = Change
            Range("K" & summary_table_row).Value = percent
            Range("K" & summary_table_row).NumberFormat = "0.00%"
           If Range("J" & summary_table_row).Value >= 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
            summary_table_row = summary_table_row + 1
            stock_volume = 0
            Change = 0
            percent = 0
        Else
            stock_volume = stock_volume + Cells(i, 7).Value
        End If
    Next i
Next ws
End Sub
