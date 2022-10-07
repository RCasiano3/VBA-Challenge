VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Current row identified
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim TickerCnt As Long
        'Last row of column A
        Dim LstRowA As Long
        'Last row of column I
        Dim LstRowI As Long
        'Variable for the percent change calculation
        Dim PercChg As Double
        'Variable for the greatest increase calculation
        Dim GreatIncrease As Double
        'Variable for the greatest decrease calculation
        Dim GreatDecrease As Double
        'Variable for the greatest total volume
        Dim GreatVolume As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Creating column headers for the WorkSheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        TickerCnt = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LstRowA = ws.Cells(Rows.Count, "A").End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
            'Loop through all rows
            For i = 2 To LstRowA
            
                'Check if the ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I
                ws.Cells(TickerCnt, "I").Value = ws.Cells(i, 1).Value
                
                'Calculate and place Yearly Change in column J
                ws.Cells(TickerCnt, "J").Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickerCnt, 10).Value < 0 Then
                
                    'Set the cell background color to Red
                    ws.Cells(TickerCnt, 10).Interior.Color = RGB(245, 0, 0)
                
                    Else
                
                    'Set the cell background color to Green
                    ws.Cells(TickerCnt, 10).Interior.Color = RGB(0, 250, 0)
                
                    End If
                    
                    'Calculate and write the percent change identified in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercChg = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickerCnt, "K").Value = Format(PercChg, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerCnt, "K").Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write the total volume identified in column L
                ws.Cells(TickerCnt, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increasing the TickCount by 1
                TickerCnt = TickerCnt + 1
                
                'Setting new start row for the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Finding the last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, "I").End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Prepare for summary
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For the greatest total volume - check if next value is larger, if yes take the new value and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                'For greatest increase -check if next value is larger, if yes take the new value and populate ws.Cells
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                'For greatest decrease - check if next value is smaller, if yes take the new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
            'Writing the summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next i
            
        'Automatic adjustment of column width
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
