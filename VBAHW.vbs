Sub Stocksloop()

For Each ws In Worksheets
    

Dim Ticker As String
Dim TotalVolume As Double
Dim Yearly As Double
Dim Startvalue As Double
Dim Percentage As Double
Dim Max As Double

Max = 0
Min = 0
Vol = 0

Startvalue = 2

TotalVolume = 0

Dim Summary As Integer

Summary = 2

Dim lr As Double


lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To lr

If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then

Ticker = ws.Cells(I, 1).Value

TotalVolume = TotalVolume + ws.Cells(I, 7).Value

Yearly = ws.Cells(I, 6).Value - ws.Cells(Startvalue, 3).Value

If ws.Cells(Startvalue, 3) <> 0 Then


Percentage = (Yearly / ws.Cells(Startvalue, 3).Value)

End If

ws.Range("I" & Summary).Value = Ticker

ws.Range("J" & Summary).Value = TotalVolume

ws.Range("K" & Summary).Value = Yearly

ws.Range("L" & Summary).Value = Percentage

     If Yearly > 0 Then
     
     ws.Range("K" & Summary).Interior.Color = vbGreen
     
     Else
     
    ws.Range("K" & Summary).Interior.Color = vbRed

   End If
   

Summary = Summary + 1

TotalVolume = 0

Startvalue = I + 1


Else

TotalVolume = TotalVolume + ws.Cells(I, 7).Value


End If

Next I

For k = 2 To lr

If ws.Range("L" & k).Value > Max Then
Max = ws.Range("L" & k).Value
T = ws.Range("I" & k).Value

End If

If ws.Range("L" & k).Value < Min Then
Min = ws.Range("L" & k).Value
Ti = ws.Range("I" & k).Value

End If

If ws.Range("J" & k).Value > Vol Then
Vol = ws.Range("J" & k).Value
Tic = ws.Range("I" & k).Value

End If

 
Next k

ws.Range("P2").Value = Max
ws.Range("O2").Value = T
ws.Range("P3").Value = Min
ws.Range("O3").Value = Ti
ws.Range("P4").Value = Vol
ws.Range("O4").Value = Tic

Next ws


End Sub


Sub Clear():

For Each ws In Worksheets
'Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I2:I" & LastRow).ClearContents
ws.Range("J2:I" & LastRow).ClearContents
ws.Range("K2:I" & LastRow).ClearContents
ws.Range("L2:I" & LastRow).ClearContents
ws.Range("O2:I" & LastRow).ClearContents
ws.Range("P2:I" & LastRow).ClearContents

Next ws
End Sub



