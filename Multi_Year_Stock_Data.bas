Attribute VB_Name = "Module1"
Sub stock_data():

'Loop for worksheets
For Each ws In Worksheets

'Cell labels
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Create Summary table cell labels
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    Dim ticker_name As String

    Dim total_vol As Double
    total_vol = 0

    Dim table_row As Integer
    table_row = 2
    
    Dim yearOpen, yearClose, yearChange, percentChange As Double
    yearOpen = ws.Cells(2, 3).Value
              
    'Last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop all tickers
    For i = 2 To lastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        total_vol = total_vol + ws.Cells(i, 7).Value
                                   
        ws.Range("I" & table_row).Value = ticker
        ws.Range("L" & table_row).Value = total_vol
            
        yearClose = ws.Cells(i, 6).Value
        yearChange = yearClose - yearOpen
        ws.Range("J" & table_row).Value = yearChange
        
            If yearChange < 0 Then
            ws.Range("J" & table_row).Interior.ColorIndex = 3
                
            Else
            ws.Range("J" & table_row).Interior.ColorIndex = 4
                
            End If
                
            'Percent Change
            If yearOpen = 0 Then
            percentChange = yearClose - yearOpen
            Else
            percentChange = (yearChange / yearOpen)
            End If
            ws.Range("K" & table_row).Value = percentChange
                                
            'Move down a row & reset total volume
            table_row = table_row + 1
            total_vol = 0
            
            'Reset yearOpen
            yearOpen = ws.Cells(i + 1, 3).Value
                             
        'If the ticker below is the same..
        Else
            'Add the total volume of like tickers
            total_vol = total_vol + ws.Cells(i, 7).Value
        
        End If
            
    Next i
    
    'Change Percent Change column to percent
    ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
    
    'Greatest % Increase and Match Ticker
    ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & table_row))
    ws.Range("P2").NumberFormat = "0.00%"
    Dim increase_Number As Integer
    increase_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'Greatest % Decrease and Match Ticker
    ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & table_row))
    ws.Range("P3").NumberFormat = "0.00%"
    Dim decrease_Number  As Integer
    decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'Greatest Volume and Match Ticker
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & table_row))
    Dim great_volNumber As Integer
    great_volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & table_row)), ws.Range("L2:L" & table_row), 0)
    
    ws.Range("O2") = ws.Cells(increase_Number + 1, 9).Value
    ws.Range("O3") = ws.Cells(decrease_Number + 1, 9).Value
    ws.Range("O4") = ws.Cells(great_volNumber + 1, 9).Value
            
Next ws
        
End Sub
