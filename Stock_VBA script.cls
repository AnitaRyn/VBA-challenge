VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

Cells(1, 8).Value = "ticker"
Cells(1, 9).Value = "yearly_change"
Cells(1, 10).Value = "percent_change"
Cells(1, 11).Value = "total_volume"

Dim ticker As String

Dim total_volume As Double

total_volume = 0

Dim open_price As Double
open_price = 0

Dim close_price As Double
close_price = 0

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim summary_table As Long

summary_table = 2

j = 0
x = 2


For I = 2 To 705714

If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

ticker = Cells(I, 1).Value

total_volume = total_volume + Cells(I, 7).Value

If total_volume = 0 Then

Range("H" & summary_table + j).Value = ticker
Range("K" & summary_table + j).Value = total_volume

summary_table = summary_table + 1

Else
close_price = Cells(I, 6).Value

open_price = Cells(x + 1, 3).Value

x = I + 1

yearly_change = close_price - open_price

Range("I" & summary_table + j).Value = yearly_change

percent_change = ((close_price - open_price) / open_price) * 100

Range("J" & summary_table + j).Value = percent_change

Range("J" & summary_table + j).Value = Round(Range("J" & summary_table + j).Value, 2)


End If

j = j + 1

total_volume = 0
yearly_change = 0
percent_change = 0

Else

total_volume = total_volume + Cells(I, 7).Value

End If

If Range("I" & summary_table + j).Value < 0 Then

Range("I" & summary_table + j).Interior.ColorIndex = 3

ElseIf Range("I" & summary_table + j).Value > 0 Then

Range("I" & summary_table + j).Interior.ColorIndex = 4

End If


Next I



End Sub
