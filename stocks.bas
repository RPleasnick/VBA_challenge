Attribute VB_Name = "Module1"
'calculates yearly change, percent change, and total stock volume on each sheet
Sub Stocks()

    Dim ticker As String
    Dim open_price As Double
    Dim end_price As Double
    Dim stock_v As Double
    Dim rowNum As Double
    Dim percent_change As Double
    Dim i As Double
    Dim yearly_change As Double
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_vol As Double
    Dim ws As Worksheet
    Dim lastRow As Double
    
    'counts the number of rows in the active sheet
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    
    i = 2
    ticker = Cells(2, 1)
    open_price = Cells(2, 3)
    
    'column headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stocks Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    
    'row discriptions
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    'formatting
    Columns("J").ColumnWidth = 12
    Columns("K").ColumnWidth = 13
    Columns("L").ColumnWidth = 15
    Columns("O").ColumnWidth = 25
    Columns("Q").ColumnWidth = 15
    
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0"
    Columns("L").NumberFormat = "0"
    
    'goes through each row in the dataset
    For rowNum = 2 To lastRow
    
    'adds the accumulative volume for each ticker
    If Cells(rowNum, 1) = Cells(rowNum + 1, 1) Then
        stock_v = stock_v + Cells(rowNum, 7)
        
    'when ticker changes
    Else
        stock_v = stock_v + Cells(rowNum, 7)
        end_price = Cells(rowNum, 6)
        yearly_change = end_price - open_price
        
        'formats cell as percentage
        Cells(i, 11).NumberFormat = "0.00%"
        
        'highlighs cell
        If yearly_change < 0 Then
            Cells(i, 10).Interior.Color = RGB(250, 0, 0)
        Else
            Cells(i, 10).Interior.Color = RGB(0, 250, 0)
        End If
        
        Cells(i, 9).Value = ticker
        Cells(i, 10).Value = yearly_change
        
        
        percent_change = yearly_change / open_price
        
        'highlighs cell
        If percent_change < 0 Then
            Cells(i, 11).Interior.Color = RGB(250, 0, 0)
        Else
            Cells(i, 11).Interior.Color = RGB(0, 250, 0)
        End If
        
        Cells(i, 11) = percent_change
        Cells(i, 12) = stock_v
        
        'calculates the % max increase
        If percent_change > max_increase Then
            max_increase = percent_change
            Cells(2, 16) = ticker
            Cells(2, 17) = max_increase
        End If
        
        'calculates % max decrease
        If percent_change < max_decrease Then
            max_decrease = percent_change
            Cells(3, 16) = ticker
            Cells(3, 17) = max_decrease
        End If
        
        'calculates max total stock volume
        If stock_v > max_vol Then
            max_vol = stock_v
            Cells(4, 16) = ticker
            Cells(4, 17) = max_vol
        End If
        
        ticker = Cells(rowNum + 1, 1)
        i = i + 1
        open_price = Cells(rowNum + 1, 3)
               
        stock_v = 0
    
    End If
      
    Next rowNum
    

End Sub


