Sub alphabetical_testng()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
    Dim Ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Variant
    Dim Greatest_percent_inc As Double
    Dim Greatest_percent_decrease As Double
    
    Dim Rowcount As Long
    Dim greatest_total_volume As Integer
    Dim summary_table As Integer
    
    Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        summary_table = 2
     
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    
    ws.Range("I:Q").EntireColumn.AutoFit
    
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                total_volume = total_volume + ws.Cells(i, 7).Value
                    ws.Range("I" & summary_table).Value = Ticker
                    ws.Range("J" & summary_table).Value = yearly_change
                summary_table = summary_table + 1
                
                total_volume = 0
           
            
            Else
                total_volume = total_volume + Cells(i, 7).Value
    
    
            End If
                
                    
    
            open_price = ws.Cells(i, 3).Value
            close_price = ws.Cells(i , 6).Value

               
                ws.Range("L" & summary_table).Value = total_volume

            yearly_change = close_price - open_price

            If (open_price = 0 And close_price = 0) Then
            percent_change = 0

            ElseIf open_price = 0 And close_price <> 0 Then
            percent_change = -1
            
            Else: percent_change = (yearly_change / open_price)
            
            ws.Range("K" & summary_table).Value = percent_change
            ws.Range("K" & summary_table).NumberFormat = "0.00%"

            
            
            
            End If
            
            If ws.Range("J" & summary_table).Value >= 0 Then
            ws.Range("J" & summary_table).Interior.ColorIndex = 4
            
            Else:
            ws.Range("J" & summary_table).Interior.ColorIndex = 3
            
            End If
        
        Next i
        
    Next ws
    
End Sub



