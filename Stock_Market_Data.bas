Attribute VB_Name = "Module1"
Sub Stock_Mark_Data()


    Dim ws As Worksheet
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As String
    Dim Total_Stock_Volume As Double
    Dim Close_Y As Double
    Dim Open_Y As Double
    Dim Summary_Table_Row As Long

 
    For Each ws In Worksheets

    ws.Activate

    Total_Stock_Volume = 0
    Summary_Table_Row = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
    Ticker = Cells(i, 1).Value
    Open_Y = ws.Cells(i, 3).Value
    Close_Y = ws.Cells(i, 6).Value
    Yearly_Change = Close_Y - Open_Y
    Percent_Change = Close_Y - Open_Y
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                

Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

Summary_Table_Row = Summary_Table_Row + 1

  Total_Stock_Volume = 0

    Else

Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        End If
    
    Next i

ws.Columns("K").NumberFormat = "0.00%"

    Next ws
 
End Sub
