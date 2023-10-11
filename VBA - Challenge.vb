Sub TickerOnAllSheets()
    Dim ws As Worksheet

    ' Loop through all the sheets in the base for the macro to run at once
    For Each ws In ThisWorkbook.Sheets
        ' Check if the sheet has data in column A
        If ws.Cells(2, 1).Value <> "" Then
            ' Call the Ticker macro for the current sheet
            Ticker ws
        End If
    Next ws
End Sub



Sub Ticker(ByVal Active_Sheet As Worksheet)

    Dim Total_Ticker_Volume As Double
    Dim Ticker_Symbol As String
    Dim Summary_Table_Row As Integer
    Dim ws As Worksheet
    Dim SummarySheet As Worksheet
    Dim LastRow As Long
    Dim FirstCValue As Double
    Dim LastFValue As Double
    Dim FirstCValueCalculated As Boolean
    
    'Adding the headers for all the calculations
   
    Active_Sheet.Range("I1").Value = "Ticker"
    Active_Sheet.Range("J1").Value = "Yearly Change"
    Active_Sheet.Range("K1").Value = "Percent Cange"
    Active_Sheet.Range("L1").Value = "Total Stock Value"
    Active_Sheet.Range("O2").Value = "Greatest % Increase"
    Active_Sheet.Range("O3").Value = "Greatest % Decrease"
    Active_Sheet.Range("O4").Value = "Greatest Total Volume"
    Active_Sheet.Range("P1").Value = "Ticker"
    Active_Sheet.Range("Q1").Value = "Value"
    
    ' Initialize the total ticker_volume in 0
    Total_Ticker_Volume = 0
    
    ' Inicialize the variable en False
    
    FirstCValueCalculated = False
    
    ' Initialize the summary table row in row 2 for results to start showing there
    Summary_Table_Row = 2
        
    
    ' Set Active_Sheet to the active sheet (Parameter)
    
    Set ws = Active_Sheet
    
    ' Set the summary sheet as the first worksheet in the workbook (Parameter)
    Set SummarySheet = ws
    
    ' Find the last row with data in column A for the active sheet becuase I dont know the last value
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all the ticker symbols until the first empty row in column A - starting in 2
    For i = 2 To LastRow
            
        ' Check if we are still within the same Ticker_Symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ' Set the Ticker_Symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
                
            ' Add to total ticker_volume from column 7
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
       
            ' Still in the loop save the first value of column C for the current Ticker_Symbol
            LastFValue = ws.Cells(i, 6).Value ' Column F is column 6
             
            ' Calculate and print the difference of the Last closing value and the first opening value.
            SummarySheet.Cells(Summary_Table_Row, 11).Value = (LastFValue - FirstCValue) / FirstCValue * 100
            
     
            ' Calculate and print the difference of the Last closing value and the first opening value.
            SummarySheet.Cells(Summary_Table_Row, 10).Value = (LastFValue - FirstCValue)
            
            ' Print the Ticker_Symbol in the Summary Table (Sheet 1, Column I)
            SummarySheet.Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
                
            ' Print the total ticker_volume to the Summary Table
            SummarySheet.Cells(Summary_Table_Row, 12).Value = Total_Ticker_Volume
    
            ' Adding one to the summary table row for it to start again
            Summary_Table_Row = Summary_Table_Row + 1
                
            ' Reset the total ticker_volume and reset FirstCValueCalculated (start sum in 0)
            
            Total_Ticker_Volume = 0
            FirstCValueCalculated = False

        ' If the cell immediately following a row is the same Ticker
        Else


            ' Calculate and print: the difference of the Last closing value and the first opening value,
            ' but only if it hasn't been calculated yet for this Ticker_Symbol
            
            If Not FirstCValueCalculated Then
                FirstCValue = ws.Cells(i, 3).Value
                ' MsgBox "FirstCValue " & FirstCValue
                
                ' Marca que se ha calculado FirstCValue
                
                FirstCValueCalculated = True
                 
            End If
            ' If the cell immediately following is the same Ticker_Symbol, add to Total_Ticker_Volume
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value

        End If

    
    Next i


   ' Call Conditional_Formatting
   Call Conditional_Formatting(Active_Sheet)

    ' Call Summary
   Call Summary2(Active_Sheet)


End Sub

Sub Conditional_Formatting(ByVal Active_Sheet As Worksheet)
    Dim YearlyChangeRange As Range
    Dim Cell As Range
    Dim LastRow As Long
    
    'Find the last row of the colum K
    
    LastRow = Active_Sheet.Cells(Active_Sheet.Rows.Count, "J").End(xlUp).Row
    
    ' Set the range from J2 to J26
    Set YearlyChangeRange = Active_Sheet.Range("J2:J" & LastRow)
    
    ' Loop For  Each to go thru all the cells in the range
    For Each Cell In YearlyChangeRange
        If Cell.Value >= 0 Then
            
            Cell.Interior.Color = RGB(0, 255, 0)
        
        Else
           
            Cell.Interior.Color = RGB(255, 0, 0)
        
        End If
    
    Next Cell


End Sub


Sub Summary2(ByVal Active_Sheet As Worksheet)
    
    Dim LastRow As Long
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim MaxValueColumnL As Double
    Dim ws As Worksheet
    Dim FoundRow As Long
    Dim FoundRow2 As Long
    Dim FoundRow3 As Long
    
    ' Set the worksheet you want to work with
    Set ws = Active_Sheet

    ' Find the last row with data in column K
    LastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row

    ' Using a function  (WorksheetFunction.Max) to find the maximum value in column K
    MaxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))

    '  Using a function  (WorksheetFunction.Min) to find the minimum value in column K
    MinValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))

    ' Using a function to find the maximum value in column L
    MaxValueColumnL = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))

   ' Print the maximum value found in cell Q2, the minimum value in cell Q3, and the maximum value in column L in cell Q4
    ws.Range("Q2").Value = MaxValue
    ws.Range("Q3").Value = MinValue
    ws.Range("Q4").Value = MaxValueColumnL
    
    ' Find the values in column I for MaxValue, MinValue, and MaxValueColumnL
    FoundRow = Application.Match(MaxValue, ws.Range("K2:K" & LastRow), 0)
    FoundRow2 = Application.Match(MinValue, ws.Range("K2:K" & LastRow), 0)
    FoundRow3 = Application.Match(MaxValueColumnL, ws.Range("L2:L" & LastRow), 0)

    ' In P2, P3 and P4, print the values from column I adding 1 becuase it starts from 0
    ws.Range("P2").Value = ws.Cells(FoundRow + 1, "I").Value
    ws.Range("P3").Value = ws.Cells(FoundRow2 + 1, "I").Value
    ws.Range("P4").Value = ws.Cells(FoundRow3 + 1, "I").Value

End Sub


