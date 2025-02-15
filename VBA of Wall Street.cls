Sub MultipleYear_Ticker()

'This will allow me to loop through all sheets in the workbook'
Dim sheet As Worksheet
For Each sheet In ActiveWorkbook.Worksheets
sheet.Activate

'Using Ctrl+ Shift + End to determine last row'
        LastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row

'Creating Variables to save values for required columns'
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Ticker_Name As String
Dim Percentage_Change As Double
Dim Total_Volume As Double
Dim Column As Integer
Dim Row As Double
Dim i As Long

'Reassigning friendly names to column headers'
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Volume"

'Setting initial values in Variables'
Volume = 0
Column = 1
Row = 2

'Setting the Opening_Price'
Opening_Price = Cells(2, Column + 2).Value

'For loop, to cycle through all tickers'
For i = 2 To LastRow

    'code that is executed 'If' the ticker changes by setting all the values for that ticker'
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
        Ticker_Name = Cells(i, Column).Value
        Cells(Row, Column + 8).Value = Ticker_Name
        Closing_Price = Cells(i, Column + 5).Value
        'CHECK' 'Calculating and adding Yearly_Change value'
        Yearly_Change = (Closing_Price - Opening_Price)
        Cells(Row, Column + 9).Value = Yearly_Change
    
        'Calcualting and adding % change'
        If (Opening_Price = 0 And Closing_Price = 0) Then
            Percentage_Change = 0
        ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
            Percentage_Change = 1
        Else
            Percentage_Change = Yearly_Change / Opening_Price
                Cells(Row, Column + 10).Value = Percentage_Change
                'formatting Percentage change output to 2-decimal places'
                Cells(Row, Column + 10).NumberFormat = "0.00%"
        End If
    
    'Adding the Total Volume, before then resetting'
    Total_Volume = Total_Volume + Cells(i, Column + 6).Value
    Cells(Row, Column + 11).Value = Total_Volume
    'formatting volume values'
    Cells(Row, Column + 11).NumberFormat = "0"
    Volume = 0
    
    'Counting '1' to the column headers'
        Row = Row + 1
    
    'Resetting the Opening_Price'
    Opening_Price = Cells(i + 1, Column + 2)
    
    'If the cells are from the same ticker'
    Else
        Total_Volume = Total_Volume + Cells(i, Column + 6).Value
    End If
'Moving onto the next Ticker'
Next i

'identifying last row of yearly change'
YearlyChange = sheet.Cells(Rows.Count, Column + 8).End(xlUp).Row

'Conditional formatting of Yearly_Change, based on the value'
For c = 2 To YearlyChange
    If (Cells(c, Column + 9).Value > 0 Or Cells(c, Column + 9).Value = 0) Then
       Cells(c, Column + 9).Interior.ColorIndex = 4
    ElseIf Cells(c, Column + 9).Value < 0 Then
        Cells(c, Column + 9).Interior.ColorIndex = 3
    End If
    
Next c

'Assigning labels for % In/Decrease and Total Volumne'
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest % Volume"
Range("O4").Value = "Greatest Total Volume"

'Assigning labels for columns Ticker and Value'
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'For loop to cycle each columns value and find the Greatest % Increase / Decrease and Total Vol.'
For j = 2 To YearlyChange
            'utilizing the ability to call Excel functions from within VBA i.e. Max & Min Function'
            If Cells(j, Column + 10).Value = Application.WorksheetFunction.Max(sheet.Range("K2:K" & YearlyChange)) Then
                Cells(2, Column + 15).Value = Cells(j, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(j, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(j, Column + 10).Value = Application.WorksheetFunction.Min(sheet.Range("K2:K" & YearlyChange)) Then
                Cells(3, Column + 15).Value = Cells(j, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(j, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(j, Column + 11).Value = Application.WorksheetFunction.Max(sheet.Range("L2:L" & YearlyChange)) Then
                Cells(4, Column + 15).Value = Cells(j, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(j, Column + 11).Value
                'I noticed the screenshot in the homework README.md didnt format the Greatest Total Volume, so I decided to'
                Cells(4, Column + 16).NumberFormat = "0"
        End If
    Next j
    
Next sheet


End Sub

