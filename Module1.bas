Attribute VB_Name = "Module1"
Sub stockAnalysis()

For Each ws In Worksheets
    Dim worksheetName As String
    worksheetName = ws.Name

    'set the header of the summary table
    Range("I" & 1).Value = "ticker"
    Range("J" & 1).Value = "Yearly Change"
    Range("K" & 1).Value = "Percent Change"
    Range("L" & 1).Value = "Total Stock Volume"
    
    'start at the second row due to the header
    firstRow = 2
    ' Determine the Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' Determine the Last Column
    lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'variable for ticker name
    Dim tickerName As String
    
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    
    'variable for total volume stock for a particular ticker name
    Dim tickerVSTotal As Double
    tickerVSTotal = 0
    
    Dim summaryTableRow As Integer
    summaryTableRow = firstRow
    
    Dim startIndex As Double
    startIndex = firstRow
    
    'loop through the rows (ticker records)
    For i = firstRow To lastRow
    
        ' Check if we are still within the same ticker name, if it is not...
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

            ' Set the ticker name
            tickerName = Cells(i, 1).Value
            
            'Yearly change
            yearlyChange = Cells(i, 6).Value - Cells(startIndex, 3).Value
           
                
            
            'Percent change and handle division by zero
            If (Cells(startIndex, 3).Value) <> 0 Then
                percentChange = yearlyChange / (Cells(startIndex, 3).Value)
            End If
            
            ' Add to the total stock volume
            tickerVSTotal = tickerVSTotal + Cells(i, 7).Value

            ' Print the ticker name in the Summary Table
            Range("I" & summaryTableRow).Value = tickerName
            
            
            'Print the yearly change and percent change in the summary table
            Range("J" & summaryTableRow).Value = yearlyChange
            If yearlyChange >= 0 Then
                Range("J" & summaryTableRow).Interior.ColorIndex = 4
            Else
                Range("J" & summaryTableRow).Interior.ColorIndex = 3
            
            End If
            
            If (Cells(startIndex, 3).Value) <> 0 Then
                Range("K" & summaryTableRow).Value = percentChange
                'Range("K" & summaryTableRow).Style = "Percent"
                Range("K" & summaryTableRow).NumberFormat = "0.00%"
            Else
                Range("K" & summaryTableRow).Value = ""
            End If
                
            ' Print the total stock volume to the Summary Table
            Range("L" & summaryTableRow).Value = tickerVSTotal

            ' Add one to the summary table row
            summaryTableRow = summaryTableRow + 1
      
            ' Reset the total stock volume
            tickerVSTotal = 0
            
            '????? got an overflow error if I had startIndex as Integer
            If i < lastRow Then
                startIndex = i + 1
            End If

            ' If the cell immediately following a row is the same ticker name
        Else

            ' Add to the total stock volume
            tickerVSTotal = tickerVSTotal + Cells(i, 7).Value

        End If

    Next i

Next ws
End Sub



