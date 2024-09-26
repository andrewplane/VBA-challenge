Attribute VB_Name = "Module1"
Sub all_sheets():
    Dim ws As Worksheet

    ' Loop through each worksheet in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        'Run stock analysis for each sheet
        Call stock_analysis(ws)
    Next ws

End Sub

Sub stock_analysis(ws As Worksheet):
    With ws
    
    'Delcare and initialize variables
    '---------------------------------------------------
    'ticker symbol
    Dim ticker As String
    ticker = ""
    
    'opening, closing and change in price / percent change
    Dim price1, price2, price_change, percent_change As Double
    price1 = 0
    price2 = 0
    price_change = 0
    percent_change = 0
    
    'Stock Volume, row counter, result row counter
    Dim volume, result_row As Integer
    Dim row As Long
    volume = 0
    row = 2
    result_row = 2
    
   'Ouput headers
   .Cells(1, 9).Value = "Ticker"
   .Cells(1, 10).Value = "Quarterly Change"
   .Cells(1, 11).Value = "Percent Change"
   .Cells(1, 12).Value = "Total Stock Volume"
      
   
   'Begin For Loop at row 2 due to header on row 1
   '---------------------------------------------------
   'Continue loop while ticker cell in not empty
   While Not IsEmpty(.Cells(row, 1))
   'While row < 1000
        
        'Check if row is mid or last stock data row
        'if current row ticker = previous row ticker
        If .Cells(row, 1).Value = .Cells(row - 1, 1) Then
            'condition that row is mid-quarter or last
            
           'check if row is last stock data row of the quarter
            'if current row ticker = next row ticker
            If .Cells(row, 1).Value = .Cells(row + 1, 1).Value Then
                'condition mid quarter row
            
                'add volume to running total
                volume = volume + .Cells(row, 7).Value
        
            Else
                'condition last row of stock data in quarter
                               
                'Read Closing Price at end of quarter
                price2 = .Cells(row, 6).Value
                
                'add volume to running total
                volume = volume + .Cells(row, 7).Value
        
                'Calculate Quarterly Change
                price_change = price2 - price1
        
                'Calculate Percentage Change and round to 2 decimal places
                percent_change = Round((price2 / price1) - 1, 2)
 
                'Read out quarterly stock information
                .Cells(result_row, 9).Value = ticker
                .Cells(result_row, 10).Value = price_change
                .Cells(result_row, 11).Value = percent_change
                .Cells(result_row, 12).Value = volume
                
                'color price change cell
                If price_change > 0 Then
                    .Cells(result_row, 10).Interior.ColorIndex = 4 'cell green for positive
                End If
                If price_change < 0 Then
                    .Cells(result_row, 10).Interior.ColorIndex = 3 'cell red for negative
                End If
                If price_change = 0 Then
                    .Cells(result_row, 10).Interior.ColorIndex = 2 'cell white for zero
                End If
                
                'advance the result row counter
                result_row = result_row + 1
                
            'end of if containing mid and last row of data conditions
            End If
        
        Else
            'condition that row is first row of quarter -previous row has different ticker
            
            'Read Ticker Symbol
            ticker = .Cells(row, 1).Value
            
            'Read opening price at begining of quarter
            price1 = .Cells(row, 3).Value
            
            'Read opening volume
            volume = .Cells(row, 7).Value
            
        'end of if that checks row condition
        End If
        
        
    'End While Loop
    row = row + 1
    Wend
    
    'Return Stock Greatest % Increase, Decrease and Total Volume
    
    'set Format for Stock Volume
    .Range("L2:L999999").NumberFormat = "0"
    .Range("Q4").NumberFormat = "0"
    
    'Labels
    .Cells(1, 16).Value = "Ticker"
    .Cells(1, 17).Value = "Value"
    .Cells(2, 15).Value = "Greatest % Increase"
    .Cells(3, 15).Value = "Greatest % Decrease"
    .Cells(4, 15).Value = "Greatest Total Volume"
    
    'Make Calculations
    Dim inc, dec, biggestVol As Double
    inc = Application.WorksheetFunction.Max(.Range("K2:K999999"))
    dec = Application.WorksheetFunction.Min(.Range("K2:K999999"))
    biggestVol = Application.WorksheetFunction.Max(.Range("L2:L999999"))
    
    '********Find Ticker for each calculated stock************
    Dim foundCell As Range
    Dim searchValue As Double
    Dim foundRow As Integer
    Dim rng As Range
    
    Set rng = ws.Columns("K:L")
    
    '% Increase
    searchValue = inc
    
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    foundRow = foundCell.row
    .Cells(2, 16).Value = .Cells(foundCell.row, 9).Value
    .Cells(2, 17).Value = inc
    
    '% Decrease
    searchValue = dec
    
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    foundRow = foundCell.row
    .Cells(3, 16).Value = .Cells(foundCell.row, 9).Value
    .Cells(3, 17).Value = dec
    
    'Greatest Total Volume
    searchValue = biggestVol
    
    Set foundCell = rng.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    foundRow = foundCell.row
    .Cells(4, 16).Value = .Cells(foundCell.row, 9).Value
    .Cells(4, 17).Value = biggestVol
    
    'Auto Fit all Cells
    .Cells.EntireColumn.AutoFit

    'Format percentages
    .Range("K2:K999999").NumberFormat = "0.00%"
    .Range("Q2:Q3").NumberFormat = "0.00%"


    End With

End Sub

Sub reset_file(): 'Resets all sheets to pre-analysis state
    Dim i As Integer
    
    'Loop to cycle through all workbook sheets and delete columns I through Q - This also resets formating
    For i = 1 To Sheets.Count
        With Sheets(i)
            .Columns("I:Q").Delete
        End With
    Next i
End Sub
