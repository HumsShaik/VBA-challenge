Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

'for loop need to create for the worksheets to loop in and it picks current sheet
Dim ws As Worksheet


    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate
        
        'creating headers for 4 columns - ticker, yearly change, percent change, and total stock volume
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("A1:L1").Font.FontStyle = "Bold"
        
        
        ' creating variables
        Dim ticker As String
        Dim stock_volume As Double
        stock_volume = 0
        Dim j As Integer
        Dim last_row As Double
        
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        j = 2
        'For loop for calculating the total stock volume
        
        For i = 2 To last_row
        
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                stock_volume = stock_volume + Cells(i, 7).Value
                
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                stock_volume = stock_volume + Cells(i, 7).Value
                
                Cells(j, 9).Value = Cells(i, 1).Value
                
                Cells(j, 12).Value = stock_volume
                
                j = j + 1
                stock_volume = 0
                
            End If
            
        Next i



        Dim opening_stock As Double
        Dim closing_stock As Double
        Dim yearly_change As Double
        Dim percentage_change
        Dim k As Integer
        
        k = 2
 
        'For loop to calculate yearly change and percentage change
        For i = 2 To last_row
        
        If Right(ws.Cells(i, 2), 4) = "0101" Then
            opening_stock = ws.Cells(i, 3)
            
            ' for the cells value ending with dec 31
        ElseIf Right(ws.Cells(i, 2), 4) = "1231" Then
            closing_stock = ws.Cells(i, 6)
            
            'formula for calculating yearly change
            
            yearly_change = closing_stock - opening_stock
            
            ' to calculate percentage change by avoiding division by zero error
            
            If opening_stock <> 0 Then
              percentage_change = yearly_change / opening_stock
            
            ws.Cells(k, 10).Value = yearly_change
            End If
            
            
        ElseIf Right(Cells(i, 2), 4) = "1230" Then
            closing_stock = ws.Cells(i, 6)
                     
            yearly_change = closing_stock - opening_stock
            
            ' to avoid division by zero error
            If opening_stock <> 0 Then
                percentage_change = yearly_change / opening_stock
          
                     
            ws.Cells(k, 10).Value = yearly_change
            
            End If
            
            ' for applying conditional formatting of reg and green color ( positve number then green or else red)
            
        If ws.Cells(k, 10).Value > 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 4
                
        ElseIf ws.Cells(k, 10).Value < 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 3
                
        End If
        ws.Cells(k, 11).Value = percentage_change
        
        k = k + 1
        
        
        
        End If
      Next i

      ' to display the percentage value by rounding off to 2 decimal places
      
      Range("K1").EntireColumn.NumberFormat = "0.00%"
      
       

        'To resize the cells using autofit.
      Dim ws_active As String
      ws_active = ActiveSheet.Name
      Worksheets(ws_active).Columns("I:L").AutoFit
      

    Next ws



End Sub
