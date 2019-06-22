Attribute VB_Name = "Module1"
Sub test_script_Easy_Final()
Dim ticker As String
Dim ticker_volume As Double
ticker_volume = 0
Dim summary_ticker_table As Integer
summary_ticker_table = 2

For i = 2 To 760192

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

ticker = Cells(i, 1).Value
ticker_volume = ticker_volume + Cells(i, 7).Value


Range("I" & summary_ticker_table).Value = ticker
Range("J" & summary_ticker_table).Value = ticker_volume

summary_ticker_table = summary_ticker_table + 1
ticker_volume = 0

Else

ticker_volume = ticker_volume + Cells(i, 7).Value

End If

Next i

End Sub

