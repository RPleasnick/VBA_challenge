Attribute VB_Name = "Module2"
'runs stocks sub routine for all years
Sub run_workbook()

Dim ws As Worksheet

Application.ScreenUpdating = False

For Each ws In Worksheets

    ws.Select

    Call Stocks

Next

Application.ScreenUpdating = True

 

End Sub
