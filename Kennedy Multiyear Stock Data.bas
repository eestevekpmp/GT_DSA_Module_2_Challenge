Attribute VB_Name = "Module2"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
End Sub
Sub RunCode()
 
lastRow = Cells(Rows.Count, 1).End(xlUp).row

Dim ticName As String

Dim volTotal As Double
volTotal = 0

Dim ticRows As Long
ticRows = 2

Dim tickerStart As Long
tickerStart = 2


Dim yChange, pChange, firstOpen, lastClose As Double
Dim max_pChange, min_pChange As Single
Dim max_pChangeticker, min_pChangeticker, max_stockvolticker As String
Dim max_stockvol As Double

Dim row As Long

    For row = 2 To lastRow
    
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        
        ticName = Cells(row, 1).Value
        Cells(ticRows, 9).Value = ticName
        
        firstOpen = Cells(tickerStart, 3).Value
        tickerStart = row + 1
        lastClose = Cells(row, 6).Value
      
        
        yChange = lastClose - firstOpen
        pChange = yChange / firstOpen
        Cells(ticRows, 10).Value = yChange
        Cells(ticRows, 11).Value = pChange
        Cells(ticRows, 11).Style = "Percent"
        
        If yChange > 0 Then
            Cells(ticRows, 10).Interior.ColorIndex = 3
        Else
            Cells(ticRows, 10).Interior.ColorIndex = 4
            
        End If
        ticRows = ticRows + 1
        If Cells(tickerStart, 3).Value = 0 Then
          For findrow = tickerStart To row
           If Cells(findrow, 3).Value <> 0 Then
                tickerStart = findrow
                Exit For
            End If
        Next findrow
        

        'tickerStart = row + 1
        volTotal = 0
        
     
            End If
        'firstOpen = Cells(tickerStart, 3).Value
        'lastClose = Cells(row, 6).Value
      

       Else
       
        volTotal = volTotal + Cells(row, 7).Value 'grabs row 2 column 7 when IF is FALSE
        
        Cells(ticRows, 12).Value = volTotal
        
      
        
      End If
     
       
        ' Source for INDEX:MATCH function code is stack overflow at:
      'https://stackoverflow.com/questions/53261956/how-to-find-get-the-variable-name-having-largest-value-in-excel-vba/53262316#53262316
            
      max_pChange = WorksheetFunction.Max(Range("K:K"))
      Range("Q2").Value = max_pChange
      Range("Q2").Style = "Percent"
      
      max_pChangeticker = Evaluate("INDEX(I:I, MATCH(MAX(K:K), K:K, 0))")
      Range("P2").Value = max_pChangeticker
      
      min_pChange = WorksheetFunction.Min(Range("K:K"))
      Range("Q3").Value = min_pChange
      Range("Q3").Style = "Percent"
      
      min_pChangeticker = Evaluate("INDEX(I:I, MATCH(MIN(K:K), K:K, 0))")
      Range("P3").Value = min_pChangeticker
      
      max_stockvol = WorksheetFunction.Max(Range("L:L"))
      Range("Q4").Value = max_stockvol
      
      max_stockvolticker = Evaluate("INDEX(I:I, MATCH(MAX(L:L), L:L, 0))")
      Range("P4").Value = max_stockvolticker
       
        
    Next row
    
  
End Sub

