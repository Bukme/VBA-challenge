Sub stock()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Summary_Table_Row As Long
    Dim Opening_Price As Double
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxTicker As String
    Dim MinTicker As String
    Dim VolTicker As String
    Dim ResultRow As Integer
    ResultRow = 2

    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Opening_Price = ws.Cells(2, 3).Value
        Ticker = ws.Cells(2, 1).Value
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
        MaxTicker = ""
        MinTicker = ""
        VolTicker = ""
        Total_Volume = 0

        ' Add headers to the Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Loop through all stock and calculate
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> Ticker Then
                Yearly_Change = ws.Cells(i, 6).Value - Opening_Price
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                If Opening_Price <> 0 Then
                    Percent_Change = Round((Yearly_Change / Opening_Price) * 100, 2)
                Else
                    Percent_Change = 0
                End If

                ws.Cells(Summary_Table_Row, 9).Value = Ticker
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change & "%"
                ws.Cells(Summary_Table_Row, 12).Value = Total_Volume

                ' Apply conditional formatting
                If Yearly_Change > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                ElseIf Yearly_Change < 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End If

                If Percent_Change > MaxIncrease Then
                    MaxIncrease = Percent_Change
                    MaxTicker = Ticker
                ElseIf Percent_Change < MaxDecrease Then
                    MaxDecrease = Percent_Change
                    MinTicker = Ticker
                End If

                If Total_Volume > MaxVolume Then
                    MaxVolume = Total_Volume
                    VolTicker = Ticker
                End If

                Summary_Table_Row = Summary_Table_Row + 1
                Opening_Price = ws.Cells(i + 1, 3).Value
                Ticker = ws.Cells(i + 1, 1).Value
                Total_Volume = 0
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        
        ws.Cells(ResultRow, 15).Value = "Greatest % Increase"
        ws.Cells(ResultRow, 16).Value = MaxTicker
        ws.Cells(ResultRow, 17).Value = MaxIncrease & "%"

        ws.Cells(ResultRow + 1, 15).Value = "Greatest % Decrease"
        ws.Cells(ResultRow + 1, 16).Value = MinTicker
        ws.Cells(ResultRow + 1, 17).Value = MaxDecrease & "%"

        ws.Cells(ResultRow + 2, 15).Value = "Greatest Total Volume"
        ws.Cells(ResultRow + 2, 16).Value = VolTicker
        ws.Cells(ResultRow + 2, 17).Value = MaxVolume

        
    Next ws
End Sub
