Attribute VB_Name = "Module2"
Option Explicit

Sub Module2Challenge()

Dim WkSht As Object

For Each WkSht In Worksheets

Dim i, j, OpeningPrice As Integer

Dim TotalVolume, LastRow, LastRowK, LastRowL As Long

Dim PriceChange, YearlyChange, PercentChange As Double

Dim MostGrowth, LeastGrowth, MostVolume As Integer


WkSht.Range("I1").Value = "Ticker"
WkSht.Range("J1").Value = "Yearly Change"
WkSht.Range("K1").Value = "Percent Change"
WkSht.Range("L1").Value = "Total Stock Volume"
WkSht.Range("P1").Value = "Ticker"
WkSht.Range("Q1").Value = "Value"
WkSht.Range("O2").Value = "Greatest % Increase"
WkSht.Range("O3").Value = "Greatest % Decrease"
WkSht.Range("O4").Value = "Greatest Total Volume"

TotalVolume = 0

j = 0

OpeningPrice = 2

LastRow = WkSht.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow

    
    If WkSht.Cells(i + 1, 1).Value <> WkSht.Cells(i, 1).Value Then

        TotalVolume = TotalVolume + WkSht.Cells(i, 7).Value
        
        WkSht.Range("I" & 2 + j).Value = WkSht.Cells(i, 1).Value
        
        WkSht.Range("L" & 2 + j).Value = TotalVolume
        
        PriceChange = WkSht.Cells(i, 6).Value - WkSht.Cells(OpeningPrice, 3).Value
        
        WkSht.Range("J" & 2 + j).Value = PriceChange
        WkSht.Range("J" & 2 + j).NumberFormat = "0.00"
        
            If PriceChange > 0 Then
                WkSht.Range("J" & 2 + j).Interior.ColorIndex = 4
        
            ElseIf PriceChange < 0 Then
                WkSht.Range("J" & 2 + j).Interior.ColorIndex = 3
            
            Else
                WkSht.Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
            
        PercentChange = PriceChange / WkSht.Cells(OpeningPrice, 3).Value
        WkSht.Range("K" & 2 + j).Value = PercentChange
        WkSht.Range("K" & 2 + j).NumberFormat = "0.00%"
        
        TotalVolume = 0
        PriceChange = 0
        PercentChange = 0
        j = j + 1
        OpeningPrice = OpeningPrice + 1
        
    Else
        TotalVolume = TotalVolume + WkSht.Cells(i, 7).Value
        
    End If
    
Next i

LastRowK = WkSht.Cells(Rows.Count, "K").End(xlUp).Row
LastRowL = WkSht.Cells(Rows.Count, "L").End(xlUp).Row

WkSht.Range("Q2").Value = WorksheetFunction.Max(WkSht.Range("K2:K" & LastRowK).Value)
WkSht.Range("Q2").NumberFormat = "0.00%"

WkSht.Range("Q3").Value = WorksheetFunction.Min(WkSht.Range("K2:K" & LastRowK).Value)
WkSht.Range("Q3").NumberFormat = "0.00%"

WkSht.Range("Q2").Value = "%" & WorksheetFunction.Max(WkSht.Range("K2:K" & LastRowK).Value)
WkSht.Range("Q3").Value = "%" & WorksheetFunction.Min(WkSht.Range("K2:K" & LastRowK).Value)
WkSht.Range("Q4").Value = WorksheetFunction.Max(WkSht.Range("L2:L" & LastRowL).Value)

MostGrowth = WorksheetFunction.Match(WorksheetFunction.Max(WkSht.Range("K2:K" & LastRow).Value), WkSht.Range("K2:K" & LastRow).Value, 0)
LeastGrowth = WorksheetFunction.Match(WorksheetFunction.Min(WkSht.Range("K2:K" & LastRow).Value), WkSht.Range("K2:K" & LastRowK).Value, 0)
MostVolume = WorksheetFunction.Match(WorksheetFunction.Max(WkSht.Range("L2:L" & LastRowL).Value), WkSht.Range("L2:L" & LastRowL).Value, 0)

WkSht.Range("K2:K" & LastRowK).NumberFormat = "0.00%"

WkSht.Range("P2").Value = WkSht.Cells(MostGrowth + 1, 9).Value

WkSht.Range("P3").Value = WkSht.Cells(LeastGrowth + 1, 9).Value

WkSht.Range("P4").Value = WkSht.Cells(MostVolume + 1, 9).Value


Next WkSht

End Sub
