VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAHomework()

For Each ws In Worksheets
      
        Dim totalV As Double
        Dim lastRow As Long
        Dim writeRow As Long
        Dim placeHold As Long
        Dim yearOpen As Double
        Dim yearClose As Double
        Dim yearChange As Double
       
writeRow = 2
totalV = 0
placeHold = 2


lastRow = ws.Cells(Rows.count, 1).End(xlUp).row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = "Greatest % Decrease"
       ws.Cells(3, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
For i = 2 To lastRow
totalV = ws.Cells(i, 7).Value + totalV

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
    ws.Cells(writeRow, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(writeRow, 12).Value = totalV
    yearOpen = ws.Cells(placeHold, 3)
    yearClose = ws.Cells(i, 6)
    yearChange = yearClose - yearOpen
    ws.Cells(writeRow, 10).Value = yearChange
            If yearOpen = 0 Then
                ws.Cells(writeRow, 11).Value = 0
            Else
                ws.Cells(writeRow, 11).Value = yearChange / yearOpen
            End If
           
            ws.Cells(writeRow, 11).NumberFormat = "0.00%"
 If ws.Cells(writeRow, 10).Value <= 0 Then
 
ws.Cells(writeRow, 10).Interior.ColorIndex = 3


            Else
            
                ws.Cells(writeRow, 10).Interior.ColorIndex = 4
                
                
            End If
            
            placeHold = i + 1
            writeRow = writeRow + 1
            totalV = 0
            
            
        End If
        
        
    Next i
    'Bonus stuff
        
    For i = 2 To lastRow
If ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then

ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
End If

            If ws.Cells(i, 11).Value <= ws.Cells(3, 17).Value Then
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                
            End If
            If ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                 ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                 
                 
                 
            End If
            
            
    Next i
    
    
Next ws


End Sub
