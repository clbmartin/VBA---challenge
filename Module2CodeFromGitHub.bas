Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim select_index As Double
    Dim first_row As Double
    Dim select_row As Double
    Dim last_row As Double
    Dim year_opening As Single
    Dim year_closing As Single
    Dim volume As Double

    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        select_index = 2
        first_row = 2
        select_row = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        volume = 0
        
        'Assigns headers etc to columns and rows
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'Loop through all rows to find unique tickers, then place each unique ticker in 9th column
        For i = first_row To last_row
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers <> tickers2 Then
                Cells(select_row, 9).Value = tickers
                select_row = select_row + 1
            End If
         Next i
    
        'Loop through all rows and add to volume if the ticker hasn't changed. Once ticker has changed, reset volume and continue.
        For i = first_row To last_row + 1
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers = tickers2 And i > 2 Then
                volume = volume + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(select_index, 12).Value = volume
                select_index = select_index + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
            
        'Loop through all rows. If previous ticker is different, assign year_opening. If next ticker is different, assign year_closing.
        select_index = 2
        For i = first_row To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                year_closing = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_opening = Cells(i, 3).Value
            End If
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                Cells(select_index, 10).Value = increase
                Cells(select_index, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                select_index = select_index + 1
            End If
        Next i
        
        'Finds min and max values, then assigns each value to proper cell
        max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        Range("Q2").Value = FormatPercent(max_per)
        Range("Q3").Value = FormatPercent(min_per)
        Range("Q4").Value = max_vol
        
        
        'Loops through columns 11 & 12. If either column contains min or max values, apply corresponding ticker to corresponding cell
        For i = first_row To last_row
            If max_per = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_per = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf max_vol = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        'Loops through column 10 then applies either green or red interior
        For i = first_row To last_row
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    
    Worksheets(ws.Name).Range("I1:Q1").Columns.AutoFit
    Worksheets(ws.Name).Range("O2:O4").Columns.AutoFit
    Worksheets(ws.Name).Range("Q2:Q4").Columns.AutoFit
    Next ws
                
End Sub
        
        
        
        
        


