Attribute VB_Name = "Module4"
Sub Color_Added()

    Dim rg As Range
    Dim lg As Long
    Dim c As Long
    Dim color_cell As Range
    
     
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For lg = 1 To c
    Set color_cell = rg(lg)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.ColorIndex = 4
            End With
        Case Is < 0
            With color_cell
                .Interior.ColorIndex = 3
            End With
       End Select
    Next lg

Next ws

End Sub
