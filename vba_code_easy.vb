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

Next ws

End Sub