# VBA-challenge
This is my homework

I have managed to run the code through one year and generate information on yearly change, percent change and total volume of each ticker.
I have also applied color coding for yearly change and modified the number of decimals for percent change.
Unfortunately I did not manage to run the code successfully through different sheets in the workbook.

This was the code I tried to apply:

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    'your code here
End Sub


I would appreciate feedback on how to make this work.
