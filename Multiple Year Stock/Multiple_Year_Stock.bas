Attribute VB_Name = "Multiple_Year_Stock"
Sub Multiple_Year_Stock()
  
  
  For Each Sheet In Worksheets
    
    
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim i As Long
    Dim RowEnd As Long
    
    Dim RowSum As Long
    Dim OpenRow As Long
    
    
        
    
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0
    YearlyOpen = 0
    YearlyClose = 0
    
    RowSum = 2
    OpenRow = 2
    

    Sheet.Range("I1").Value = "Ticker"
    Sheet.Range("J1").Value = "Yearly Change"
    Sheet.Range("K1").Value = "Percent Change"
    Sheet.Range("L1").Value = "Total Stock Volume"
   
    

    RowEnd = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To RowEnd

                TotalStockVolume = TotalStockVolume + Sheet.Cells(i, 7).Value
           
            If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
                Ticker = Sheet.Cells(i, 1).Value
                Sheet.Range("I" & RowSum).Value = Ticker
                
                Sheet.Range("L" & RowSum).Value = TotalStockVolume
                TotalStockVolume = 0
                
                YearlyOpen = Sheet.Range("C" & OpenRow)
                YearlyClose = Sheet.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                Sheet.Range("J" & RowSum).Value = YearlyChange

            
            If YearlyChange < 0 Then
                 Sheet.Range("J" & RowSum).Interior.Color = RGB(255, 0, 0)
              Else
                 Sheet.Range("J" & RowSum).Interior.Color = RGB(124, 252, 0)
              End If
                  
                  
             If YearlyOpen <> 0 And YearlyChange <> 0 Then
                PercentChange = YearlyChange / YearlyOpen
             Else
                PercentChange = 0
             End If
             
                Sheet.Range("K" & RowSum).Value = PercentChange
            
                Sheet.Range("K" & RowSum).NumberFormat = "0.00%"
                
                  
                  
                RowSum = RowSum + 1
                OpenRow = i + 1
                
                End If
                
            Next i
            
            
            Next Sheet
            
            
    End Sub


Sub Max_Min()

    Dim Sheet As Worksheet
    Dim Max As Double
    Dim Min As Double
    
For Each Sheet In Worksheets
    Sheet.Activate
    
    Sheet.Range("O2").Value = "Greatest % Increase"
    Sheet.Range("O3").Value = "Greatest % Decrease"
    Sheet.Range("O4").Value = "Greatest Total Volume"
    Sheet.Range("P1").Value = "Ticker"
    Sheet.Range("Q1").Value = "Value"
    
    
    Max = 0
    Min = 0
    
    
RowEnd = Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To RowEnd
    If Cells(i, 11).Value > Max Then
    Max = Cells(i, 11).Value
    Cells(2, 16) = Cells(i, 9).Value
    Cells(2, 17) = Max * 100 & "%"
    End If
Next i
    

For i = 2 To RowEnd
    If Cells(i, 11).Value < Min Then
    Min = Cells(i, 11).Value
    Cells(3, 16) = Cells(i, 9).Value
    Cells(3, 17) = Min * 100 & "%"
    End If
Next i


For i = 2 To RowEnd
    If Cells(i, 12).Value > Max Then
    Max = Cells(i, 12).Value
    Cells(4, 16) = Cells(i, 9).Value
    Cells(4, 17) = Max
    End If
Next i
   

Next Sheet


End Sub



