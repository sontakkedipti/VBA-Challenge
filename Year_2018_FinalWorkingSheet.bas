Attribute VB_Name = "Year2018"
Sub Year2018():

Dim colj As String
Dim nextrow As Long
nextrow = 2

Dim StPrice As Double
Dim endPrice As Double
Dim totVolume As LongLong
totVolume = 0

Dim maxchange As Double
maxchange = 0
Dim minchange As Double
minchange = 0
Dim maxticker As String
Dim minticker As String

Dim maxtotalVol As LongLong
maxtotalVol = 0

'Counting total Rows for Col A
totalrows = Cells(Rows.Count, "A").End(xlUp).Row
       
'logic to run for end of the row for Col A
For i = 2 To totalrows
       
         'Total volume logic to add volume for each row
          totVolume = Cells(i, 7).Value + totVolume
       
         'Logic to find uniue ticker and saving to colJ
         If (Cells(i, 1) <> Cells(i - 1, 1)) Then
             
            'Copy the value to colj
             Cells(nextrow, 10).Value = Cells(i, 1).Value
           
            'Logic to add start price and end price
             StPrice = Cells(i, 3)
           
           
          ElseIf (Cells(i, 1) <> Cells(i + 1, 1)) Then
           
            endPrice = Cells(i, 6)
            'total volume for unique ticker
            Cells(nextrow, 13).Value = totVolume
           
            Change = endPrice - StPrice
            'Yearly chnage for unique ticker
           
            Cells(nextrow, 11).Value = Change
                   
                    'Highlight diff in Red and Green
                    If (Change < 0) Then
                    Cells(nextrow, 11).Interior.ColorIndex = 3
                   
                    Else
                    Cells(nextrow, 11).Interior.ColorIndex = 10
                   
                    End If
           
           percentChange = (Change / StPrice)
           
           '%chnage for unique ticker
           Cells(nextrow, 12).Value = percentChange
           
                  'Logic for Greatest % Increase
                  If (percentChange > maxchange) Then
                  maxchange = percentChange
                  maxticker = Cells(i, 1)
                                                   
                  ElseIf (percentChange < minchange) Then
                  'Logic for Greatest % Descrease
                  minchange = percentChange
                  minticker = Cells(i, 1)
                                                   
                  End If
               
                  If (totVolume > maxtotalVol) Then
                  maxtotalVol = totVolume
                  totticker = Cells(i, 1)
                 
                  End If
                                         
            nextrow = nextrow + 1
            totVolume = 0
                         
        Else
         'do nothing
        End If
               
        Next i
           
         Cells(2, 15).Value = "Greatest % Increase"
         Cells(2, 16).Value = maxticker
         Cells(2, 17).Value = maxchange
         
         
         Cells(3, 15).Value = "Greatest % Decrease"
         Cells(3, 16).Value = minticker
         Cells(3, 17).Value = minchange
         
         Cells(4, 15).Value = "Greatest Total Stock"
         Cells(4, 16).Value = totticker
         Cells(4, 17).Value = maxtotalVol
         
         Cells(1, 10).Value = "Ticker"
         Cells(1, 11).Value = "Yearly Change"
         Cells(1, 12).Value = "Percent Change"
         Cells(1, 13).Value = "Total Stock Volume"
        
End Sub


