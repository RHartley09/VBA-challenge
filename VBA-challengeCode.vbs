Sub VBATest()
'Gather all ticker symbols and populate them on another column.
Dim i As Long
Dim ws As Worksheet
Dim lRow As Long
Dim Summary_Table_Row As Long
Dim ticker As String
Dim stock_total As Double
Dim openprice As Double
Dim closeprice As Double

For Each ws In Worksheets

Summary_Table_Row = 2
lRow = Cells(Rows.Count, "G").End(xlUp).Row
stock_total = 0
ticker = ("")

'Loop

openprice = ws.Cells(2, 3).Value

For i = 2 To lRow

'Check stocks

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ticker = ws.Cells(i, 1).Value

closeprice = ws.Cells(i, 6).Value

stock_total = stock_total + ws.Cells(i, 7).Value

ws.Range("M" & Summary_Table_Row).Value = closeprice - openprice

If openprice = 0 Then
ws.Range("N" & Summary_Table_Row).Value = 0

Else

ws.Range("N" & Summary_Table_Row).Value = Round((closeprice - openprice) / openprice * 100, 2) & "%"

End If

ws.Range("K" & Summary_Table_Row).Value = ticker

ws.Range("L" & Summary_Table_Row).Value = stock_total

If ws.Range("M" & Summary_Table_Row).Value > 0 Then

ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4

Else

ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3

End If

Summary_Table_Row = Summary_Table_Row + 1

openprice = ws.Cells(i + 1, 3).Value

stock_total = 0

Else: stock_total = stock_total + ws.Cells(i, 7).Value

'Print

End If


Next i

Next ws

End Sub

