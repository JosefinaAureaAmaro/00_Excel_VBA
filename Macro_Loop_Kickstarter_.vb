Sub UniqueTicker()
    For Each ws In Worksheets
    'Variables
    Dim Column As Integer
    Dim i As Long
    Dim j As Long



    TotalVol = 0

    'header
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Volume"

    'Assign Variables
    Column = 1
    j = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To last_row
 
    volume = Range("G" & i).Value
    TotalVol = TotalVol + volume

    'Loop through rows in column
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        TickerRow = Range("A" & i).Value
        
        Range("I" & j).Value = TickerRow
        Range("J" & j).Value = TotalVol
        
        TotalVol = 0
        j = j + 1
        
    End If

    Next i
    Next ws
    
End Sub

