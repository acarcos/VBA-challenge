Sub alpha()

Dim ws As Worksheet
Dim ticker As String

'To read every worksheet
For Each ws In ActiveWorkbook.Sheets

    'Array to save calculations
    Dim data(0 To 3) As Integer

    'Rows to display calculatios
    Dim k As Integer
    k = 2 'For rows inside loop

    'Headers
    ws.Range("I1") = "Ticker"
    ws.Range("O1") = "Ticker"
    ws.Range("J1") = "Yearly_Change"
    ws.Range("K1") = "Percentage_Change"
    ws.Range("L1") = "Total_stock_volume"
    ws.Range("P1") = "Value"
    ws.Range("N2") = "Greatest_increase"
    ws.Range("N3") = "Greatest_decrease"
    ws.Range("N4") = "Greatest_volume"

    'Count to last row of all data
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            'Read data
            fecha = ws.Cells(i, 2).Value
            'deit = Split(fecha, "2016")
            'Instead of using split function, module is used to analyze data from different years
            deit_v2 = fecha Mod 10000

            'Values are saved depending of initial and final date
            'If deit(1) = 101 Then
            If deit_v2 = 101 Then
                open_year = ws.Cells(i, 3).Value
            End If
    
            'If deit(1) = 1230 Then
            If deit_v2 = 1230 Then
                close_year = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1).Value
    
                'Ticker value
                ws.Cells(k, 9).Value = ticker
                'Yearly change
                ws.Cells(k, 10).Value = close_year - open_year
                'Percentage change
                If close_year = 0 Then
                    ws.Cells(k, 11).Value = "0"
                Else
                    ws.Cells(k, 11).Value = Cells(k, 10).Value / close_year
                End If
                'Total volume (sum of all volume values in a ticker)
                ws.Cells(k, 12).Value = Application.Sum(Range(Cells(1, 7), Cells(i, 7)))
    
                k = k + 1 'To display calculation in next row
    
            End If
    
            Next i

    'Length of data calculations
    lastrowdata = ws.Cells(Rows.Count, 11).End(xlUp).Row

    'Conditional formatting in yearly change
    Dim rng As Range
    Dim condition1 As FormatCondition, condition2 As FormatCondition

    'Dataset
    Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(lastrowdata, 10))
    'Clear any conditional formatting
    rng.FormatConditions.Delete

    'Conditions to color green (positive values) and red (negative values)
    Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "0")
    Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "0")

    With condition1
        .Interior.ColorIndex = 43 'Green color
        End With
    With condition2
        .Interior.ColorIndex = 22 'Red color
    End With

    ' FOR MAX/MIIN VALUE

    'Maximum value of percentage
    ws.Cells(2, 16).Value = Application.Max(Range(ws.Cells(2, 11), ws.Cells(lastrowdata, 11)))
    ws.Cells(2, 16).NumberFormat = "0.00%"
    'Respective ticker
    ws.Cells(2, 15).Value = Application.Index(Range(ws.Cells(2, 9), ws.Cells(lastrowdata, 9)), Application.Match(ws.Cells(2, 16).Value, Range(ws.Cells(2, 11), ws.Cells(lastrowdata, 11)), 0))

    'Minimum value of percentage
    ws.Cells(3, 16).Value = Application.Min(Range(ws.Cells(2, 11), ws.Cells(lastrowdata, 11)))
    ws.Cells(3, 16).NumberFormat = "0.00%"
    'Respective ticker
    ws.Cells(3, 15).Value = Application.Index(Range(ws.Cells(2, 9), ws.Cells(lastrowdata, 9)), Application.Match(ws.Cells(3, 16).Value, Range(ws.Cells(2, 11), ws.Cells(lastrowdata, 11)), 0))

    'Greatest volume
    ws.Cells(4, 16).Value = Application.Max(Range(ws.Cells(2, 12), ws.Cells(lastrowdata, 12)))
    'Respective ticker
    ws.Cells(4, 15).Value = Application.Index(Range(ws.Cells(2, 9), ws.Cells(lastrowdata, 9)), Application.Match(ws.Cells(4, 16).Value, Range(ws.Cells(2, 12), ws.Cells(lastrowdata, 12)), 0))

Next ws

End Sub


