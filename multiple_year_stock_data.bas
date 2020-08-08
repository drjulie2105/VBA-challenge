Attribute VB_Name = "Module1"
Sub TestMarket()

'Loop through all worksheets

For Each ws In Worksheets

'Determine the last row of sheet
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Insert Ticker, Yearly Change, Percent Change, Total Stock Volume heading in column i
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'Declare Ticker Symbol

Dim TickerSymbol As String

'Declare Summary Table Entries

Dim SummaryTableRow As Integer
        SummaryTableRow = 2

'Set first date of Year Open
    
Dim FirstOpen As Double
    FirstOpen = ws.Cells(2, 3).Value
    

'Set YearlyChange and Percent Change
  
    Dim YearlyChange As Double

    Dim PercentChange As Double
    
'Declare Total Volume
               
    Dim TotalVolume As Double
        TotalVolume = 0

      
'Loop through each row to get TickerSymbol in column i
For i = 2 To LastRow


   
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
        TickerSymbol = ws.Cells(i, 1).Value
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        Dim LastClose As Double
        LastClose = ws.Cells(i, 6).Value
        
        YearlyChange = LastClose - FirstOpen
        
        ws.Range("I" & SummaryTableRow).Value = TickerSymbol

        ws.Range("L" & SummaryTableRow).Value = TotalVolume

        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        
            If FirstOpen = 0 Then
                PercentChange = 0
            Else
                PercentChange = YearlyChange / FirstOpen
            End If
                
            ws.Range("K" & SummaryTableRow).Value = PercentChange
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
       
        
        SummaryTableRow = SummaryTableRow + 1
    
        TotalVolume = 0
        
        FirstOpen = ws.Cells(i + 1, 3)
        
        
    Else
        
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        
    End If
        
Next i

YearlyChangeLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

    For j = 2 To YearlyChangeLastRow
    
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3

        End If
        
    Next j

Next ws

End Sub














