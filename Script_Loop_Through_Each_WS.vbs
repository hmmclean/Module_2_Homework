Sub ForEachws():
    
    Dim ws As Worksheet
    
    'Loops through each worksheet and runs the Summary_Tables macro
    For Each ws In ThisWorkbook.Worksheets
        Call Summary_Tables(ws)

    Next ws

End Sub