Attribute VB_Name = "Module1"
Sub Stock_Data()

For Each ws In Worksheets
    Dim sheets As String
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
Next ws

End Sub
