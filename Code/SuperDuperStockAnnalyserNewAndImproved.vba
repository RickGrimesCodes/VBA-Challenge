Attribute VB_Name = "SuperDuperStockAnnalyserNewandImproved"
Sub Module2Challenge():
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
With ws
    'haha, nuclear bombs go brrrrrr
    ws.Cells.ClearFormats
    ws.Range("H1", "Q:Q") = ""
    'Marks the columns with appropriate value identifiers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Making labels for the columns and rows I will use for the 'Bonus' section of this assignment
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volumn"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Dim'ing a bunch of stuff I will use for stuff later
    Dim rowCount As Double
    Dim VolTotal As Double
    Dim PercentChange As Double
    Dim OpenBeggining As Double
    Dim CloseEnd As Double
    Dim summaryRow As Long
    Dim Row As Double
    Dim RepeatCounter As Long
    summaryRow = 2
    
    Dim Ticker As String
    
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'I don't even know how I made this work tbh
    For Row = 2 To rowCount
        If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
        
            Ticker = ws.Cells(Row, 1).Value
            
            VolTotal = VolTotal + ws.Cells(Row, 7).Value
            
            ws.Cells(summaryRow, 9).Value = Ticker
            
            ws.Cells(summaryRow, 12).Value = VolTotal
            
            VolTotal = 0
            
            OpenBeggining = ws.Cells((Row - RepeatCounter), 3)
            CloseEnd = ws.Cells(Row, 6)
            ws.Cells(summaryRow, 10) = CloseEnd - OpenBeggining
            ws.Cells(summaryRow, 11) = CloseEnd / OpenBeggining - 1 'there's probably a better way to do this than the way I did.
                'some snazzy conditional formating for yearly change -- thanks https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
                If ws.Cells(summaryRow, 10) > 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(summaryRow, 10) = 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 6 'I hope this is orange for 0 okay
                    Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    
                    
                End If
            
            summaryRow = summaryRow + 1
            
            RepeatCounter = 0 'probably not the best solution
        Else
            RepeatCounter = RepeatCounter + 1 'neither is this
            VolTotal = VolTotal + ws.Cells(Row, 7).Value
            
            
        End If
    Next Row
    Dim MultiVariable As Double
    Dim CellAddress As Range
    Dim CellAddress2 As Range
        'hope this is correct, used allot of googling to figure this out, couln't figure out how to make summaryRow variable to work how I wanted with .Find, it works but I'm not happy with it
        'max for Percent Change with Ticker
    MultiVariable = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & summaryRow).Value)
    ws.Range("Q2").Value = MultiVariable
    Set CellAddress = ws.Range("K:K").Find(What:=MultiVariable, After:=ws.Range("K1"))
    Set CellAddress2 = CellAddress.Offset(0, -2)
    ws.Range("P2").Value = ws.Range(CellAddress2.Address)
        'min for Percent Change with Ticker
    MultiVariable = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & summaryRow).Value)
    ws.Range("Q3").Value = MultiVariable
    Set CellAddress = ws.Range("K:K").Find(What:=MultiVariable, After:=ws.Range("K1"))
    Set CellAddress2 = CellAddress.Offset(0, -2)
    ws.Range("P3").Value = ws.Range(CellAddress2.Address)
        'max Total Stock Volume with Ticker
    MultiVariable = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & summaryRow).Value)
    ws.Range("Q4").Value = MultiVariable
    Set CellAddress = ws.Range("L:L").Find(What:=MultiVariable, After:=ws.Range("L1"))
    Set CellAddress2 = CellAddress.Offset(0, -3)
    ws.Range("P4").Value = ws.Range(CellAddress2.Address)
    
    'Autofit & final snazzy formating
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#,###"
    ws.Range("L2:L" & summaryRow).NumberFormat = "#,###"
    ws.Columns("A:G").AutoFit
    ws.Columns("I:L").AutoFit
    ws.Columns("O:Q").AutoFit

End With
Next ws
End Sub

