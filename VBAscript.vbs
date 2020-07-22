Attribute VB_Name = "Module1"
Sub stocks()

'Loop through all sheets
For Each ws In Worksheets

'establish variables
Dim openStock As Double
Dim closeStock As Double
Dim Ticker As String
Dim OutputRow As String
Dim percentchange As Double
Dim stockvolume As Double
Dim max As Double
Dim min As Double
Dim tag As String
Dim volume As Double
Dim tag2 As String
Dim volume2 As Double
Dim volMax As Double
Dim tag3 As String


OutputRow = 2

'Determine the last active row
Dim lastRow As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Ticker = ws.Cells(2, 1).Value
openStock = ws.Cells(2, 3).Value

'Add headers for new columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'main for loop
For i = 2 To lastRow

    stockvolume = stockvolume + ws.Cells(i, 7).Value
    
If (Ticker <> ws.Cells(i, 1).Value) Then
        closeStock = ws.Cells(i - 1, 6).Value
        
        If openStock <> 0 Then
            percentchange = Round((closeStock - openStock) / (openStock) * 100, 2)
            Else: percentchange = 0
            End If
        
        ws.Cells(OutputRow, 10).Value = closeStock - openStock
        ws.Cells(OutputRow, 12).Value = stockvolume
        ws.Cells(OutputRow, 9).Value = Ticker
        ws.Cells(OutputRow, 11).Value = percentchange
        
        
        Ticker = ws.Cells(i, 1).Value
        openStock = ws.Cells(i, 3).Value
        
        OutputRow = OutputRow + 1
        stockvolume = 0


End If
 
'Conditional for coloring
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
        End If

'Challenges
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Total Stock Volume"

max = WorksheetFunction.max(ws.Range("K:K"))
    If ws.Cells(i, 11) = max Then
    tag = ws.Cells(i, 9).Value
    ws.Cells(2, 16) = tag
    volume = ws.Cells(i, 12).Value
    ws.Cells(2, 17) = volume
    
    
    End If

min = WorksheetFunction.min(ws.Range("K:K"))
    If ws.Cells(i, 11) = min Then
    tag2 = ws.Cells(i, 9).Value
    ws.Cells(3, 16) = tag2
    volume2 = ws.Cells(i, 12).Value
    ws.Cells(3, 17) = volume2
    
    End If
    
volMax = WorksheetFunction.max(ws.Range("L:L"))
    If ws.Cells(i, 12) = volMax Then
    tag3 = ws.Cells(i, 9).Value
    ws.Cells(4, 16) = tag3
    ws.Cells(4, 17) = volMax
    
    End If



Next i


Next ws
 
End Sub
