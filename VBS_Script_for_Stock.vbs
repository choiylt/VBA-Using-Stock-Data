Sub alpha()


    Dim Stock_Name As String
    
    Dim Stock_Total As Double

    Dim Percent_Change As Double

    
    Dim Table As Integer
        Table = 2
     

 For I = 2 To 70926
 
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        Stock_Name = Cells(I, 1).Value
        Stock_Total = Stock_Total + Cells(I, 7).Value
        
        Range("I" & Table).Value = Stock_Name
        Range("J" & Table).Value = Stock_Total
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Total Stock Volume"
        Columns("J:J").EntireColumn.AutoFit
        
        
        
        Table = Table + 1
        
        Stock_Total = 0
        
        Else
        
            Stock_Total = Stock_Total + Cells(I, 7).Value
        End If
    Next I
            

End Sub


