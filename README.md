# VBA---challenge
Cindy Martin VBA - challenge
•	Please Note I was able to source code for this project as listed below, but only to the point of getting the “Total Volume”, then I hit a wall of coding inability, and my code just did not work. I have rewatched all of the videos and was just not grasping correctly what to write or how to get this module finished on my own. I wrote code, but just got error messages. I searched online for help (at this time, it was the weekend and I felt the quickest option was to look online). I found the same project on GitHub and after clicking on whatever I could, it gave me code for the project. I then copied it and put it on a clean version of: Multiple Year Stock data and ran it, thinking it would give me an error message or just not work. To my shock, it did EVERYTHING for the project, except autofitting the columns.  I have attached the project for the Multiple Stock Year Data, which has my code. I have submitted my code for the following: Titles, Ticker Symbol, Total Stock Volume, AutoFitting and looping through the worksheets. My code was not working when I added it to the GitHub code. The remainder I used the GitHub code for the alphabet testing file. I could not incorporate other than my autofit code into the GitHub coding.

I have reviewed the GitHub code as much as I can to understand what it is doing. After reviewing the supplied code, I know I would not have been able to create code that actually ran correctly for the rest of this project without this code. I have supplied as much information about the source of this GitHub code as possible.  

As I’ve said, I used my own code for, creating: Titles, Ticker Symbol list, Total Volume list, AutoFitting and Looping through the worksheets in the Multiple Year Stock Data file.  The alphabet file was primarily done with the GitHub file.

VBA Module 2 Code
-	Code to go through Multiple Sheets from: VBA Wells Fargo Demo (Part 2)
Sub WellsFargo1()
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
         MsgBox (WorksheetName)
       
    Next ws
End Sub

-	Code to add Titles to Columns and Other area from: VBA Wells Fargo Demo (Part 2)

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"


-	Code to list unique stock Ticker from: VBA Scripting Unit/Lesson/03-04-2024…/06-Credit Card Checker
https://git.bootcampcontent.com/UNC-Charlotte/UNCC-VIRT-DATA-PT-03-2024-U-LOLC/-/commit/3a55639d20b03127d1e6e1e83b24d0bb214c1428?page=2#651ee509e4044ed10f8c6fae4216c70d86238e94
Dim Ticker_Name As String
DimTicker_Total As Double
Ticker_Total=0
Dim Summary_Table_Row As Integer
Summary_Table_Row=11
For i=2 To 75350
If Cells(i+1,1).Value<>Cells(I,1).Value Then
	Ticker_Name = Cells(i,1).Value
	Ticker_Total=Ticker_Total+Cells(i,3).Value
	Range(“I) & Summary_Table_Row).Value=Ticker_Total
Summary_Table_Row=Summary_Table_Row +1
	Ticker_Total=0
Else
Ticker_Total=Ticker_Total+Cells(i,12).Value
	End If
Next i	



Code to loop through worksheets: VBA Bonus Demo 01-Using sheet references as Variables in VBA Scripting (Updated)
Dim sheet1, sheet2, sheet3 As Worksheet
Set sheet1=Worksheets(“2018”)
Set sheet2=Worksheets(“2019”)
Set sheet3=Worksheets(“2020”)
Dim string As String
String1=”Say anything”
sheet1.Range(“I5”).Value=string1
sheet2.Range(“I5”).Value=string1
sheet3.Range(“I5”).Value=string1


Code to AutoFit Titles: VBA Bonus Demo 03 - Aggregates (Part 1) (Updated)
Worksheets(ws.Name).Range("I1:Q1").Columns.AutoFit
Worksheets(ws.Name).Range("O2:O4").Columns.AutoFit
(Code in the Attached Screenshots File)

•	Code that  I would have used for conditional formatting, if I didn’t find GitHub code:
https://git.bootcampcontent.com/UNC-Charlotte/UNCC-VIRT-DATA-PT-03-2024-U-LOLC/-/commit/3a55639d20b03127d1e6e1e83b24d0bb214c1428#f38ecb7c5f9a50c5e329b15f2cf6c886af28734d
Colors VBA Scripting Unit/Lessons/03-04-2024
Sub formatter()

  ' Set the Font color to Red
  Range("A1").Font.ColorIndex = 3

  ' Set the Cell Colors to Red
  Range("A2:A5").Interior.ColorIndex = 3

  ' Set the Font Color to Green
  Range("B1").Font.ColorIndex = 4

  ' Set the Cell Colors to Green
  Range("B2:B5").Interior.ColorIndex = 4

  ' Set the Color Index to Blue
  Range("C1").Font.ColorIndex = 5

  ' Set the Cell Colors to Blue
  Range("C2:C5").Interior.ColorIndex = 5

  ' Set the Color Index to Magenta
  Range("D1").Font.ColorIndex = 7

For Row =1 To lastRow
https://git.bootcampcontent.com/UNC-Charlotte/UNCC-VIRT-DATA-PT-03-2024-U-LOLC/-/commit/3a55639d20b03127d1e6e1e83b24d0bb214c1428#dfab4123d519c7451ceb0261b61b42365a374a8c
VBA Scripting Unit/Lessons/03-04-2024/Stu_Gradebook
  ' Check if the student's grade is greater than or equal to 90...
  If Cells(2, 2).Value >= 90 Then

      ' Establish that the grade is Passing
      Cells(2, 3).Value = "Pass"

      ' Color the Passing grade green
      Cells(2, 3).Interior.ColorIndex = 4

      ' Set the letter grade to "A"
      Cells(2, 4).Value = "A"

  ' Check if the student's grade is greater than or equal to 80...
  ElseIf Cells(2, 2).Value >= 80 Then

      ' Establish that the grade is Passing
      Cells(2, 3).Value = "Pass"

      ' Color the Passing grade green
      Cells(2, 3).Interior.ColorIndex = 4

      ' Set the letter grade to "B"
      Cells(2, 4).Value = "B"

  ' Check if the student's grade is greater than or equal to 70...
  ElseIf Cells(2, 2).Value >= 70 Then

      ' Establish that the grade is a Warning
      Cells(2, 3).Value = "Warning"

      ' Color the Warning grade yellow
      Cells(2, 3).Interior.ColorIndex = 6

      ' Set the letter grade to "C"
      Cells(2, 4).Value = "C"

  ' Check if the students' grade is failing
  Else

      ' Establish that the grade is Failing
       Cells(2, 3).Value = "Fail"

      ' Color the Failing grade red
      Cells(2, 3).Interior.ColorIndex = 3

      ' Set the letter grade to "F"
      Cells(2, 4).Value = "F"

  End If

End Sub
COPIED SOURCE CODE FROM GITHUB FOR THE ALPHABET FILE

GitHub source code information used for:
Yearly Change, % Change, Greatest % Change, Greatest % Decrease, Greatest Total Volume and Conditional Formatting 


Source Code for below areas: https://github.com/theodoremoreland/YearlyStocks.git
GitHub - Theodore Moreland – Yearly Stocks

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

‘Added my autofit code here
Code to AutoFit Titles: VBA Bonus Demo 03 - Aggregates (Part 1) (Updated)
Worksheets(ws.Name).Range("I1:Q1").Columns.AutoFit
Worksheets(ws.Name).Range("O2:O4").Columns.AutoFit

    Next ws
                
End Sub 
