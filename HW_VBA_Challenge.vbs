Sub Every_Sheey()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stuck_Market
    Next
    Application.ScreenUpdating = True
End Sub

Sub Stuck_Market()
    'For Each ws In Worksheets
        Dim i As LongLong
        Dim ticker_counter As Integer
        Dim last_row As LongLong
        Dim rng_vol_1 As Range
        Dim rng_vol_2 As Range
        On Error Resume Next
        
        last_row = Cells(Rows.Count, 1).End(xlUp).Row '- 50000
        'MsgBox (last_row)
        ticker_counter = 1
    
        'last_row = Range("K" & 3).Value
        'last_row = 10000
       ' Range("M" & ticker_counter + 1).Value = Cells(2, 1).Value
       ' Range("L" & ticker_counter + 1).Value = ticker_counter
       ' ticker_counter = ticker_counter + 1
        
        opening_row = 2
        
        For i = 2 To last_row
        
        'Makes the list of the different Ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'Range("L" & ticker_counter + 1).Value = ticker_counter
                Range("M" & ticker_counter + 1).Value = Cells(i, 1).Value
                'Opening Row
                'Range("N" & ticker_counter + 1).Value = opening_row
                'Closing Row
                'Range("O" & ticker_counter + 1).Value = i
                'Yearl Change Column F - C
                closing_row = i
                o_value = Range("C" & opening_row).Value
                c_value = Range("F" & closing_row).Value
                Year_Change = c_value - o_value
                Range("N" & ticker_counter + 1).Value = Year_Change
                
                'Percentage Change
                P_Change = c_value / o_value - 1
                Range("O" & ticker_counter + 1).Value = P_Change
                
                'Volume
                rng_vol_1 = Range("G" & opening_row) '"G"&closing_row)
                rng_vol_2 = Range("G" & closing_row)
                'Range("R" & ticker_counter + 1).Value =
                ' G = 7
                'ActiveCell.Value = Application.Sum(Range(Cells(opening_row, 7), Cells(closing_row, 7)))
                'Range("P" & ticker_counter + 1).Value = ActiveCell.Value
                
                Sum_Volume = Application.Sum(Range(Cells(opening_row, 7), Cells(closing_row, 7)))
                Range("P" & ticker_counter + 1).Value = Sum_Volume
                
                
                
                ticker_counter = ticker_counter + 1
                opening_row = i + 1
            End If
            
            
        Next i
        
        'Format Percentage
        Dim rng As Range
        Dim cond1 As FormatCondition
        Dim cond2 As FormatCondition
        Set rng = Range("O:O")
        Set rng2 = Range("N:N")
       
        Set rng3 = Range("P:P")
        rng.NumberFormat = "0.00%"
        rng3.NumberFormat = "0,000"
        Set cond1 = rng2.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set cond2 = rng2.FormatConditions.Add(xlCellValue, xlLess, "=0")
        With cond1
            .Interior.Color = vbGreen
        End With
        With cond2
            .Interior.Color = vbRed
        End With
        
        'Header
        Range("N1").FormatConditions.Delete
        Range("M1").Value = "Ticker"
        Range("N1").Value = "Yearly Change"
        Range("O1").Value = "Percent Change"
        Range("P1").Value = "Total Stock Volume"
        
        'Bonus
        last_row_ticker = Cells(Rows.Count, 13).End(xlUp).Row
        Dim great_increas As Double
        Dim tricker_great_increase As String
        
        ' Greates Change
        great_increase = Cells(2, 15).Value
        great_decrease = Cells(2, 15).Value
        greater_volume = Cells(2, 16).Value
        ticker_great_increase = Cells(2, 13).Value
        ticker_great_decrease = Cells(2, 13).Value
        ticker_great_volume = Cells(2, 13).Value
        
        For i = 2 To last_row_ticker
            If Cells(i, 15).Value > great_increase Then
                great_increase = Cells(i, 15).Value
                ticker_great_increase = Cells(i, 13).Value
            End If
            If Cells(i, 15).Value < great_decrease Then
                great_decrease = Cells(i, 15).Value
                ticker_great_decrease = Cells(i, 13).Value
            End If
            If Cells(i, 16).Value > greater_volume Then
                greater_volume = Cells(i, 16).Value
                ticker_greater_volume = Cells(i, 13).Value
            End If
                
            
        Next i
        Range("R2").Value = "Greatest % Increase"
        Range("S2").Value = ticker_great_increase
        Range("T2").Value = great_increase
        
    
        Range("R3").Value = "Greatest % decrease"
        Range("S3").Value = ticker_great_decrease
        Range("T3").Value = great_decrease
        Range("T2:T3").NumberFormat = "0.00%"
        
        
        Range("R4").Value = "Greatest total volume"
        Range("S4").Value = ticker_greater_volume
        Range("T4").Value = greater_volume
        Range("T4").NumberFormat = "0,000"
    'Next ws
End Sub


