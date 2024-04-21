Attribute VB_Name = "Module1"
Sub Module2Challenge():
  
    'Marks the columns with appropriate value identifiers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    'Making labels for the columns and rows I will use for the 'Bonus' section of this assignment
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volumn"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Dim'ing a bunch of stuff I will use for stuff later
    Dim rowCount As Double
    Dim VolTotal As Double
    Dim PercentChange As Double
    Dim OpenBeggining As Double
    Dim CloseEnd As Double
    Dim summaryRow As Long
    Dim Row As Long
    Dim RepeatCounter As Long
    summaryRow = 2
    
    Dim Ticker As String
    
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'I don't even know how I made this work tbh
    For Row = 2 To rowCount
        If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
        
            Ticker = Cells(Row, 1).Value
            
            VolTotal = VolTotal + Cells(Row, 7).Value
            
            Cells(summaryRow, 9).Value = Ticker
            
            Cells(summaryRow, 12).Value = VolTotal
            
            VolTotal = 0
            
            OpenBeggining = Cells((Row - RepeatCounter), 3)
            CloseEnd = Cells(Row, 6)
            Cells(summaryRow, 10) = CloseEnd - OpenBeggining
            Cells(summaryRow, 11) = CloseEnd / OpenBeggining - 1 'there's probably a better way to do this than the way I did.
                'some snazzy conditional formating for yearly change -- thanks https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
                If Cells(summaryRow, 10) > 0 Then
                    Cells(summaryRow, 10).Interior.ColorIndex = 4
                    ElseIf Cells(summaryRow, 10) = 0 Then
                    Cells(summaryRow, 10).Interior.ColorIndex = 6 'I hope this is orange for 0 okay
                    Else
                    Cells(summaryRow, 10).Interior.ColorIndex = 3
                    
                    
                End If
            
            summaryRow = summaryRow + 1
            
            RepeatCounter = 0 'probably not the best solution
        Else
            RepeatCounter = RepeatCounter + 1 'neither is this
            VolTotal = VolTotal + Cells(Row, 7).Value
            
            
        End If
    Next Row
        
    Dim PercentChangeSummaryRange As Range
    Dim CellLocSummary As CellFormat
        
    PercentChangeSummaryRange = Range("K1:K999")
        'max ticker & value
    HighestPercent = Application.WorksheetFunction.Max(Range("PercentChangeSummaryRange").Value)
    CellLocSummary = Application.WorksheetFunction.Max(Range("PercentChangeSummaryRange").Cell)
    Cells("P2").Value = Cells("CellLocSummary" - ",2")
        'min ticker & value
    LowestPercent = Application.WorksheetFunction.Max("PercentChangeSummaryRange")
    CellLocSummary = Application.WorksheetFunction.Max(Range("PercentChangeSummaryRange").Cell)
    Cells("P2").Value = Cells("CellLocSummary" - ",2")
        'max ticker & value for stock vol

    
    'Autofit & formating
    Columns("A:G").AutoFit
    Columns("I:L").AutoFit
    Columns("O:Q").AutoFit
    Columns("K").NumberFormat = "0.00%"
    
End Sub

