Attribute VB_Name = "vbachallenge"
Sub vbachallenge()

'This counter is used to create our new ticker table
Dim j As Long

'These counters are required to calculate total stock volume
Dim Opening As Double
Dim Closing As Double

'This is used to define the end point of the raw datasheet
Dim NumRows As Long

'Loop to run through each worksheet in workbook
Dim ws As Worksheet
For Each ws In Worksheets

'House cleanup first- clears formatting and old row data just in case macro is rerun several times
ws.Range("J:Q").FormatConditions.Delete
ws.Range("J:Q").Value = ""

'Reset of certain variables for each new worksheet
Opening = ws.Cells(2, 3).Value
NumRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
j = 2

'Create Required Headers
ws.Range("J1:M1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'Creating new ticker table. Counter i is used to search the Raw dataset row by row
For i = 2 To NumRows
'total volume calculation
    ws.Cells(j, 13).Value = ws.Cells(j, 13).Value + ws.Cells(i, 7)

'Using nextcells principle
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
'Ticker Name and Closing value
            ws.Cells(j, 10).Value = ws.Cells(i, 1).Value
            Closing = ws.Cells(i, 6).Value
'Yearly Change and Percent Change
            ws.Cells(j, 11).Value = Closing - Opening
            ws.Cells(j, 12).Value = FormatPercent((Closing - Opening) / Opening, 2)
            Opening = ws.Cells(i + 1, 3)
            j = j + 1
        End If

Next i

'Conditional Formatting for Yearly and Percent Change
With ws.Range("K:L").FormatConditions.Add(xlCellValue, xlLess, "=0")
ws.Range("K:L").FormatConditions(1).Interior.Color = vbRed
End With
With ws.Range("K:L").FormatConditions.Add(xlCellValue, xlGreater, "=0")
ws.Range("K:L").FormatConditions(2).Interior.Color = vbGreen
End With
ws.Range("K1:L1").FormatConditions.Delete

'Greatest % Increase, % decrease, total volume
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("Q2").Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("L:L")))
ws.Range("P2").Value = ws.Cells(Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("L:L"), 0), 10)

ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("Q3").Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("L:L")))
ws.Range("P3").Value = ws.Cells(Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("L:L"), 0), 10)

ws.Range("O4").Value = "Greatest total Volume"
ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("M:M"))
ws.Range("P4").Value = ws.Cells(Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("M:M"), 0), 10)

'Cleanup tasks- Automatically change the column widths to complete the table
ws.Columns("J:Q").AutoFit

Next ws
End Sub

