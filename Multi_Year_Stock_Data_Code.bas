Attribute VB_Name = "Module1"
Sub Stock_Data()

Dim ws As Worksheet

For Each ws In Worksheets

      Dim Ticker As String
      Dim Open_Price As Double
      Dim Close_Price As Double
      Dim Yearly_Change As Double
      Dim Percent_Change As Double
      Dim Greatest_Increase As Double
      Dim Greatest_Decrease As Double
      Dim Greatest_Total As Currency
      
      Dim Price_Row As Long
      Price_Row = 2
      
      Total = 0
      Greatest_Increase = 0
      Greatest_Decrease = 0
      Greatest_Total = 0
      
      Dim Table_Row As Integer
      Table_Row = 2
      
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      For i = 2 To Last_row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               Ticker = ws.Cells(i, 1).Value
               
               Total = Total + ws.Range("G" & i).Value
               
               Ticker = ws.Range("I" & Table_Row).Value
               
               Total = ws.Range("L" & Table_Row).Value
               
               Open_Price = ws.Range("C" & Price_Row).Value
               Close_Price = ws.Range("F" & i).Value
               Yearly_Change = Close_Price - Open_Price
               
                  If Open_Price = 0 Then
                      Percent_Change = 0
                     Else
                         Percent_Change = Yearly_Change / Open_Price
                  End If
                 
                  ws.Range("J" & Table_Row).Value = Yearly_Change
                  ws.Range("K" & Table_Row).Value = Percent_Change
                  ws.Range("K" & Table_Row).NumberFormat = "0.00%"
                  
                        If ws.Range("J" & Table_Row).Value > 0 Then
                            ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                        End If
                  
                  Table_Row = Table_Row + 1
                  Price_Row = i + 1
               
                  Total = 0
            Else
              Total = Total + ws.Range("G" & i).Value
                 
            End If
                      
       Next i
         
        Dim r As Range
        
        Set r = ws.Range("K2:K3001")
        Greatest_Decrease = Application.WorksheetFunction.Min(r)
        Greatest_Increase = Application.WorksheetFunction.Max(r)
       
        Set r2 = ws.Range("L2:L3001")
        Greatest_Total = Application.WorksheetFunction.Max(r2)
       
        For i = 2 To Last_row_col_I
        
        If ws.Cells(i, 11).Value = Greatest_Increase Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    End If
    
    If ws.Cells(i, 11).Value = Greatest_Decrease Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    End If
    
    If Cells(i, 12).Value = Greatest_Total Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If
      
 Next i
 
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("P4").Value = Greatest_Total_Ticker
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q4").Value = Greatest_Total
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
   
   Next ws
   
End Sub
