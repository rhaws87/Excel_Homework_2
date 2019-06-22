Attribute VB_Name = "Module2"
Sub test_script_moderate()
Dim ticker As String
Dim ticker_volume As Double
ticker_volume = 0
Dim ticker_min As Double
ticker_end = 0

Dim ticker_max As Double
ticker_start = 0

Dim ticker_spread As Double
ticker_spread = 0


Dim summary_ticker_table As Integer
summary_ticker_table = 2

For i = 2 To 70926

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i, 1).Value
ticker_volume = ticker_volume + Cells(i, 7).Value
ticker_end = WorksheetFunction.Min(Cells(i, 6).Value)
ticker_start = Cells(i, 3).Value

Range("I" & summary_ticker_table).Value = ticker
Range("J" & summary_ticker_table).Value = ticker_volume
Range("K" & summary_ticker_table).Value = ticker_start
Range("L" & summary_ticker_table).Value = ticker_end
Range("M" & summary_ticker_table).Value = ticker_spread

summary_ticker_table = summary_ticker_table + 1
ticker_volume = 0

ticker_spread = (ticker_start - ticker_end) / (ticker_start)

Else

ticker_volume = ticker_volume + Cells(i, 7).Value

If Range("M" & summary_ticker_table).Value > 0 Then
Range("M" & summary_ticker_table).Interior.ColorIndex = 4

If Range("M" & summary_ticker_table).Value <= 0 Then
Range("M" & summary_ticker_table).Interior.ColorIndex = 8

End If


End If

End If




Next i




End Sub


