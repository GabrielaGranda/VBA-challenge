Attribute VB_Name = "Module1"
Sub Module2()

Dim last_row As Long
Dim ticker As String
Dim Summary_Table_Row As Integer
Dim match1 As Long
Dim match2 As Long

    For i = 1 To 4
        
        Worksheets("Q" & i).Select
        
        Summary_Table_Row = 2
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        Range("I1").Value = "Ticker"
        Range("j1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        For j = 2 To last_row
        
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
    
        ticker = Cells(j, 1).Value
        Range("I" & Summary_Table_Row).Value = ticker
        
        match1 = WorksheetFunction.Match(Range("I" & Summary_Table_Row), Range("A:A"), 0)
        match2 = WorksheetFunction.Match(Range("I" & Summary_Table_Row), Range("A:A"), 1)
        
        Range("J" & Summary_Table_Row).Value = WorksheetFunction.Index(Range("F:F"), match2) - WorksheetFunction.Index(Range("C:C"), match1)
        
        Range("K" & Summary_Table_Row).Value = (WorksheetFunction.Index(Range("F:F"), match2) / WorksheetFunction.Index(Range("C:C"), match1)) - 1
        
        Range("L" & Summary_Table_Row).Value = WorksheetFunction.SumIf(Range("A:A"), Range("I" & Summary_Table_Row), Range("G:G"))
        
        If Range("K" & Summary_Table_Row).Value > 0 Then
        
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf Range("K" & Summary_Table_Row).Value < 0 Then
            
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        End If
    Next j
        
    Range("P2").Value = WorksheetFunction.Max(Range("K:K"))
    Range("O2").Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("P2").Value, Range("K:K"), 0))
    Range("P3").Value = WorksheetFunction.Min(Range("K:K"))
    Range("O3").Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("P3").Value, Range("K:K"), 0))
    Range("P4").Value = WorksheetFunction.Max(Range("L:L"))
    Range("O4").Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Range("P4").Value, Range("L:L"), 0))

    Range("K:K").NumberFormat = "0.00%"
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"

    Next i
    
End Sub

