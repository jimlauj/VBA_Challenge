Sub FinalVers1():
For Each ws In Worksheets

'Columns
    'create new columns for ticker, yearly change ($), percent change, and total stock volume
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change ($)"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
    'Create column headers for summary box
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"

'Variables
            Dim WorksheetName As String
    'The ticker symbol
        'set initial variable for holding the ticker
            Dim ticker As String
    'The stocktotal
        'set initial value for holding the total stock volume
            Dim stocktotal As Double
            stocktotal = 0
    'Variable for full year change calculation
            Dim year_change As Double
            year_change = 0
    'Variable for percent change calculation
            Dim PerChange As Double
    'Variable for greatest increase calculation
            Dim GreatInc As Double
    'Variable for greatest decrease calculation
            Dim GreatDec As Double
    'Variable for greatest total volume
            Dim GreatVol As Double
    'Current row
            Dim i As Long
    'need j for "<open>" cell (a 2nd row value that stays at the start of the ticker change)
            Dim j As Long
            j = 2
    'keep track of location for each value in the summary table
            Dim summarytable_row As Integer
            summarytable_row = 2

'Get the WorksheetName
        WorksheetName = ws.Name
    
'Determine Last Row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all tickers
            For i = 2 To LastRow
    
    'check if still in the same ticker, if its not ..
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set ticker name
        ticker = ws.Cells(i, 1).Value
    'Add to stock total
        stocktotal = stocktotal + ws.Cells(i, 7).Value
    'Print ticker in the summary table
        ws.Range("I" & summarytable_row).Value = ticker
    'Print stock total to the summary table
        ws.Range("L" & summarytable_row).Value = stocktotal
    'Reset the stocktotal
             stocktotal = 0
    
    'Calculate and write Yearly Change in column J (#10)
        year_change = year_change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
        ws.Range("J" & summarytable_row).Value = year_change
    'Reset the year_change
            year_change = 0
    
    'Conditional formating for less than/more than 0
        If ws.Range("J" & summarytable_row).Value < 0 Then
    'Set cell background color to red
        ws.Range("J" & summarytable_row).Interior.ColorIndex = 3
        Else
    'Set cell background color to green
        ws.Range("J" & summarytable_row).Interior.ColorIndex = 4
        End If
    
    'Calculate and write percent change in column K (#11)
        If ws.Cells(j, 3).Value <> 0 Then
        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    'Percent formating
        ws.Range("K" & summarytable_row).Value = Format(PerChange, "Percent")
        Else
        ws.Range("K" & summarytable_row).Value = Format(0, "Percent")
        End If
    'add one to the summary table row
        summarytable_row = summarytable_row + 1
    'Set new start row of the ticker block
        j = i + 1

    'if the cell immediately following a row is the same ticker...
        Else
    'add to the stocktotal
        stocktotal = stocktotal + ws.Cells(i, 7).Value
    'add to the year_change
        year_change = year_change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
  
    End If
Next i

'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

 'Prepare for summary
        GreatVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value

For i = 2 To LastRowI

'For greatest increase-check down column if next value is larger-if yes take over a new value and populate Cells
        If ws.Cells(i, 11).Value > GreatInc Then
        GreatInc = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        Else
        GreatInc = GreatInc
        End If

'For greatest decrease-check down column if next value is smaller-if yes take over a new value and populate Cells
        If ws.Cells(i, 11).Value < GreatDec Then
        GreatDec = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        Else
        GreatDec = GreatDec
        End If

'For greatest total volume-check down column if next value is larger-if yes take over a new value and populate Cells
        If ws.Cells(i, 12).Value > GreatVol Then
        GreatVol = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        Else
        GreatVol = GreatVol
        End If


 'Write summary results in format
        ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
        ws.Cells(3, 17).Value = Format(GreatDec, "Percent")
        ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")

Next i

'Adjust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
Next ws
End Sub

