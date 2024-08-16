Sub Homework()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate


Cells(1, 9) = "Ticker"
Cells(1, 10) = "Quarterly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stack Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greates Total Volume"
    


Dim i As Long

Dim j As Long

Dim total As Double

total = 0

Dim ticker_name As String

Dim ticker_begin As Double

Dim ticker_end As Double

Dim ticker_diff As Double

Dim lastRow As Long

Dim maxval As Double

Dim minval As Double

Dim maxtotal As Double

Dim maxRow As Double

Dim minRow As Double

Dim totalrow As Double



lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

lastrowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

lastrowL = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

 j = i - (i - 2)

For i = 2 To lastRow


   
    
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        
        
        ticker_name = Cells(i, 1).Value
        
        ticker_begin = Cells(j, 3).Value
                
        ticker_end = Cells(i, 6).Value
        
        ticker_diff = ticker_end - ticker_begin
        
        total = total + Cells(i, 7)
                
        Cells(Summary_Table_Row, 10).Value = ticker_diff
        
        Cells(Summary_Table_Row, 9).Value = ticker_name
        
        Cells(Summary_Table_Row, 11).Value = FormatPercent(((ticker_end - ticker_begin) / ticker_begin))
        
        
        Cells(Summary_Table_Row, 12).Value = total
                
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        total = 0
        
        j = i + 1
        
        
        Else
        
        ticker_diff = ticker_end - ticker_begin
        
        total = total + Cells(i, 7)
                 
        
        End If
        
     If Cells(i, 10).Value > 0 Then
     
     Cells(i, 10).Interior.ColorIndex = 4
     
     ElseIf Cells(i, 10).Value < 0 Then
     
     Cells(i, 10).Interior.ColorIndex = 3
     
     Else
     
     End If
     
     
        
    Next i
        
        maxval = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowK))
        minval = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowK))
        maxtotal = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowL))
        
        maxRow = Application.WorksheetFunction.Match(maxval, ws.Columns(11), 0)
        minRow = Application.WorksheetFunction.Match(minval, ws.Columns(11), 0)
        totalrow = Application.WorksheetFunction.Match(maxtotal, ws.Columns(12), 0)
                
        
        Cells(2, 16).Value = Cells(maxRow, 9)
        Cells(3, 16).Value = Cells(minRow, 9)
        Cells(4, 16).Value = Cells(totalrow, 9)
  
        Cells(2, 17).Value = FormatPercent(maxval)
        Cells(3, 17).Value = FormatPercent(minval)
        Cells(4, 17).Value = maxtotal
        
     
  
  Next ws
  
        
        
End Sub

