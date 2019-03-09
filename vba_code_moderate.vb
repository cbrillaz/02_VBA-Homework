Sub Homework()

Dim ws As Worksheet
    
    For Each ws In Sheets
        ws.Activate
        

        Dim Ticket As String
        Dim Total_stock_volume As Double

        Total_stock_volume = 0

        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2

        Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

            For i = 2 To lastrow

                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                Ticket = Cells(i, 1).Value
        
                Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
        
                ws.Range("I" & Summary_Table_Row).Value = Ticket
        
                ws.Range("J" & Summary_Table_Row).Value = Total_stock_volume
        
                Summary_Table_Row = Summary_Table_Row + 1
        
                Total_stock_volume = 0
        
             Else
    
                Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
        
             End If
        
            Next i
            
        Cells(1, 9).Value = "Ticket"
        Cells(1, 10).Value = "Total Stock Volume"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
    
Dim Open_price As Double
Dim Close_price As Double
Dim year_change As Double
Dim percent_change As Double

Summary_Table_Row = 2

For i = 2 To lastrow

    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Open_price = ws.Cells(i, 3).Value
        
        ElseIf ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Close_price = ws.Cells(i, 6).Value
    
        year_change = Close_price - Open_price
        
           If Open_price = 0 Then
            percent_change = 0
            Else: percent_change = (Close_price - Open_price) / Open_price
            End If
    
        ws.Range("K" & Summary_Table_Row).Value = year_change
        ws.Range("L" & Summary_Table_Row).Value = percent_change
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
    
        Summary_Table_Row = Summary_Table_Row + 1
        
        year_change = 0
        
    End If
    
  Next i

Dim lastrow2 As Long
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

Summary_Table_Row = 2

For i = 2 To lastrow2

    If Cells(i, 11) < 0 Then
    
    Cells(i, 11).Interior.ColorIndex = 3
    
    Else: Cells(i, 11).Interior.ColorIndex = 4
    
    End If
    
  Next i
  
Next ws
    
End Sub