Sub Summary_Tables(ws As Worksheet):

    'Add Column Headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
        
    'Declaring Variables
    Dim TickerName As String
    Dim TotalSum As Double
    Dim Summary_Table_Row As Integer
    Dim i As Long
    Dim EndRow As Long

    'Setting initial values
    TotalSum = 0
    Summary_Table_Row = 2
    EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through data
    For i = 2 To EndRow
    
        'Check if cells are not equal and save values
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            TickerName = ws.Cells(i, 1).Value
            TotalSum = TotalSum + ws.Cells(i, 7).Value
        
            'Place values in summary table
            ws.Range("J" & Summary_Table_Row).Value = TickerName
            ws.Range("M" & Summary_Table_Row).Value = TotalSum
        
            'Next row in summary table
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Reset total sum for next ticker name
            TotalSum = 0
        Else
            'Adding up all the volumes for the equal ticker names
            TotalSum = TotalSum + ws.Cells(i, 7).Value
        End If
    Next i
    
    'Declaring Variables
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PriceArray(2) As Double
    Dim Yearly As Double
    Dim Percent As Double
    Dim PercentFormat As String

    'Setting initial values
    OpenPrice = 0
    ClosePrice = 0
    Yearly = 0
    Percent = 0
    Summary_Table_Row = 2

    'Loop through price data
    For i = 2 To EndRow
        'Store open price value
        If ws.Cells(i, 2).Value = "20180102" Or ws.Cells(i, 2).Value = "20190102" Or ws.Cells(i, 2).Value = "20200102" Then
            OpenPrice = ws.Cells(i, 3).Value
            PriceArray(0) = OpenPrice
    
        'Store close price value
        ElseIf ws.Cells(i, 2).Value = "20181231" Or ws.Cells(i, 2).Value = "20191231" Or ws.Cells(i, 2).Value = "20201231" Then
            ClosePrice = ws.Cells(i, 6).Value
            PriceArray(1) = ClosePrice
        
            'Calculate yearly change
            Yearly = PriceArray(1) - PriceArray(0)
        
            'Calculate percent change
            Percent = Yearly / PriceArray(0)
            PercentFormat = FormatPercent(Percent)
        
            'Place yearly and percent change values in summary table
            ws.Range("K" & Summary_Table_Row).Value = Yearly
            ws.Range("L" & Summary_Table_Row).Value = PercentFormat
        
            'Next row in summary table
            Summary_Table_Row = Summary_Table_Row + 1
        End If
    Next i
    
    'Format Summary Table
    With ws.Range("J1:M1")
    .NumberFormat = "Text"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Columns.AutoFit
    End With
    
    EndRowSummaryTable = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    With ws.Range("J1:M" & EndRowSummaryTable)
    .BorderAround xlContinous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    'Conditional Formatting for Yearly Change Column
    For i = 2 To EndRowSummaryTable
    
        If ws.Cells(i, 11) < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
    
        Else: ws.Cells(i, 11).Interior.ColorIndex = 4
    
        End If
    Next i
    
    'Add Headers to Increase_Decrease Summary Table
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    'Declaring variables
    Dim Increase As Double
    Dim IncreaseFormat As String
    Dim Decrease As Double
    Dim DecreaseFormat As String
    Dim Volume As Variant
    
    'Finding min and max for Increase_Decrease Summary Table
    Increase = Application.WorksheetFunction.Max(ws.Range("L1:L" & EndRowSummaryTable))
    IncreaseFormat = FormatPercent(Increase)
    Decrease = Application.WorksheetFunction.Min(ws.Range("L1:L" & EndRowSummaryTable))
    DecreaseFormat = FormatPercent(Decrease)
    Volume = Application.WorksheetFunction.Max(ws.Range("M1:M" & EndRowSummaryTable))
    
    'Adding the values to the Increase_Decrease Summary Table
    ws.Cells(2, 18).Value = IncreaseFormat
    ws.Cells(3, 18).Value = DecreaseFormat
    ws.Cells(4, 18).Value = Volume
    
    'Adding appropriate Ticker name to Increase_Decrease Summary Table
    For i = 2 To EndRowSummaryTable
        If ws.Cells(i, 12).Value = Increase Then
            ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
        
        ElseIf ws.Cells(i, 12).Value = Decrease Then
            ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
        
        ElseIf ws.Cells(i, 13).Value = Volume Then
            ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
        End If
    Next i
        
    'Format Increase_Decrease Summary Table
    EndRowIDST = ws.Cells(Rows.Count, 18).End(xlUp).Row
    
    With ws.Range("P1:R1")
    .NumberFormat = "Text"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("P1:R" & EndRowIDST)
    .BorderAround xlContinous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Columns.AutoFit
    End With

End Sub