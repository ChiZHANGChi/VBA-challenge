Attribute VB_Name = "Module1"
Sub stock():

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

    If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then
    
    [i1] = "ticker"

    [j1] = "yearly Change"

    [k1] = "precent Change"

    [l1] = "total Stock Volume"

    [O2] = "Greatest % Increase"

    [o3] = "Greatest % Decrease"

    [o4] = "Greatest Total Volume"

    [p1] = "ticker"

    [Q1] = "volume"
    

        Total = 0

        Index = 2

        firstOpen = 0
        

            lastRow = Cells(Rows.Count, "A").End(xlUp).Row

   

        For i = 2 To lastRow

   

        Total = Total + Cells(i, "G")

   

        If firstOpen = 0 Then

       

            firstOpen = Cells(i, "c")

             

        End If

      

        If Cells(i, "A") <> Cells(i + 1, "A") Then

           
            YearlyCh = Cells(i, "F") - firstOpen

           

            Cells(Index, "i") = Cells(i, "A")

           

            Cells(Index, "J") = YearlyCh

           

            Cells(Index, "K") = YearlyCh / firstOpen

           

            Cells(Index, "L") = Total

                       

            Total = 0

           

            firstOpen = 0

       

            Index = Index + 1

           
           End If
    

    Next i

   

End If

Next ws

 

End Sub

Sub Formatting()
    Dim ws As Worksheet
    
    Dim rng As Range
    
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("2018")
     

  
    Set rng = ws.Range("J2:J" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row)

   
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    .Interior.Color = RGB(255, 0, 0)
    
    End With

    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    .Interior.Color = RGB(0, 255, 0)
    
    End With
    
End Sub
