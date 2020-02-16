

Sub Stock_Ticker()

'Define variables to allocate values
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim vol As Double
vol = 0
Dim row As Double
row = 2
'dim column as double ; column = 1

    For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Determine the last row'
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

    'Fix initial open price
    year_open = ws.Cells(2, 3).Value

    'Loop through the ticker column to check if the we are in same ticker name'
    Dim i As Double
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'set ticker name value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(row, 9).Value = ticker
            
            'Set year_close price
            year_close = ws.Cells(i, 6).Value

            'Set yearly change
            yearly_change = year_close - year_open
            ws.Cells(row, 10).Value = yearly_change

            'Set %age change
            If (year_open = 0 And year_close = 0) Then
                percentage_change = 0
            ElseIf (year_open = 0 And year_close <> 0) Then
                 percentage_change = 1
            Else: percentage_change = yearly_change / year_open
                ws.Cells(row, 11).Value = percentage_change
                ws.Cells(row, 11).NumberFormat = "0.00%"
            End If
            'Add volume of ticker
            vol = vol + ws.Cells(i, 7).Value
            ws.Cells(row, 12).Value = vol

            'add 1 row to the row and reset year_open value and volume
            row = row + 1
            year_open = ws.Cells(i + 1, 3).Value
            vol = 0

        'if ticker is the same
        Else: vol = vol + ws.Cells(i, 7).Value
        End If
    Next i

    'Determine last row of percentage change column to apply conditional formatting
    changelastrow = ws.Cells(Rows.Count, 10).End(xlUp).row
    'Set the colors
    Dim j As Integer
    For j = 2 To changelastrow
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 10

        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3

        Else: ws.Cells(j, 10).Interior.ColorIndex = 2
        End If
    Next j

    'Set the greatest % increase, decrease and volume
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest %age Increase"
    ws.Range("O3").Value = "Greatest %age Decrease"
    ws.Range("O4").Value = "Greatest total volume"

    'Find greatest increase, decrease values and volume
    Dim k As Integer
    For k = 2 To changelastrow
        If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & changelastrow)) Then
            ws.Range("P2") = ws.Cells(k, 9).Value
            ws.Range("Q2") = ws.Cells(k, 11).Value
            ws.Range("Q2").NumberFormat = "0.00%"

        ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & changelastrow)) Then
            ws.Range("P3") = ws.Cells(k, 9).Value
            ws.Range("Q3") = ws.Cells(k, 11).Value
            ws.Range("Q3").NumberFormat = "0.00%"
        
        ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & changelastrow)) Then
            ws.Range("P4") = ws.Cells(k, 9).Value
            ws.Range("Q4") = ws.Cells(k, 12).Value
        
        End If
    Next k


 Next ws

End Sub
