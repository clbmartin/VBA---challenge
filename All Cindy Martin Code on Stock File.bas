Attribute VB_Name = "Module1"

        
Sub AllCindyWrittenCode()
            
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        

Dim Ticker_Name As String
Dim Ticker_Total As Double
Ticker_Total = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
For i = 2 To 753500
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Name = Cells(i, 1).Value
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
     ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
     ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
Summary_Table_Row = Summary_Table_Row + 1
    Ticker_Total = 0
Else
Ticker_Total = Ticker_Total + Cells(i, 7).Value
    End If
        Next i
        
Worksheets(ws.Name).Range("I1:Q1").Columns.AutoFit
Worksheets(ws.Name).Range("O2:O4").Columns.AutoFit

      
    Next ws
    
     

    
End Sub
