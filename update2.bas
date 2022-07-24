Attribute VB_Name = "Module1"
Sub Stock_Data()

For Each ws In Worksheets
    Dim sheets As String
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
Dim Ticker As String
Dim Volume As Double
   Volume = 0
    
Dim row As Integer
    row = 2
    
For i = 2 To Lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        
        Volume = Volume + Cells(i, 7).Value
        
        Range("I" & row).Value = Ticker
        
        Range("L" & row).Value = Volume
        
        row = row + 1
        
        Volume = 0
    
Else

Volume = Volume + Cells(i, 7).Value

End If

Next i

    
Next ws

End Sub
