Sub StockMarket()

    'Scrip for all worksheets
    Dim ws As Worksheet
    
    For Each ws In Worksheets

       'Set variables
        Dim TixSym As String
        Dim WorkshName As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearChg As Double
        Dim PercentChg As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatTotVol As Double
        
        'Label column headers on each worksheet
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
  
        'Label analysis cells
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        'Label column headers for analysis
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
    
        ' Last Row in worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initial variable for ticker symbol volume
        Dim TixSym_Total As Double
        TixSym_Total = 0

        ' Summary table tracking ticker symbol
        Dim TixSym_SummTbl_Row As Integer
        TixSym_SummTbl_Row = 2

        ' Identifying Worksheet   UNSURE IF NEEDED YET
        WorkshName = ws.Name
  
        'Identify 1st opening price of year
        OpenPrice = Cells(2, 3).Value
    
        'Loop through stocks
        For i = 2 To LastRow
             
            'Check if ticket symbol changed
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Identify ticker symbol
            TixSym = Cells(i, 1).Value
            
            'Identify closing price at end of year
            ClosePrice = Cells(i, 6).Value
            
            'Add volume
            TixSym_Total = TixSym_Total + Cells(i, 7).Value
        
            'Print ticker symbol in table
            ws.Range("I" & TixSym_SummTbl_Row).Value = TixSym
                     
            'Calculate Yearly Change
            YearChg = (ClosePrice - OpenPrice)
            
            'Print Yearly Change
            ws.Range("J" & TixSym_SummTbl_Row).Value = YearChg
                        
            'Calculate Percent Change
            PercentChg = ((ClosePrice - OpenPrice) / OpenPrice)
                    
            'Print Percent Change
            ws.Range("K" & TixSym_SummTbl_Row).Value = PercentChg
            
            'Print total volume of stock
            ws.Range("L" & TixSym_SummTbl_Row).Value = TixSym_Total
            
            'Add one to row to summary
            TixSym_SummTbl_Row = TixSym_SummTbl_Row + 1
            
            'Reset volume
            TixSym_Total = 0
            
            'Reset Opening Price
            OpenPrice = Cells(i + 1, 3).Value
            
            'If ticket symbol did NOT change
            Else
        
            'Add Volume total
            TixSym_Total = TixSym_Total + Cells(i, 7).Value
        
            End If
            
        Next i
        
            For i = 2 To LastRow
            
                If ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                End If
                
            Next i
                          
            'Format columns
            ws.Range("J:J").NumberFormat = "0.00"
            ws.Range("K:K").NumberFormat = "0.00%"
            
            ws.Columns("A:Q").AutoFit

            'Calculate Greatest Percent Increase
            GreatIncr = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
            
            'Calculate Greatest Percent Decrease
            GreatDecr = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
                     
            'Calculate Greatest total volume
            GreatTotVol = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
            
            For i = 2 To LastRow
            
                If GreatIncr = ws.Cells(i, 11).Value Then
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                End If
            Next i
            
            For i = 2 To LastRow
            
               If GreatDecr = ws.Cells(i, 11).Value Then
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
               End If
            Next i
            
            For i = 2 To LastRow
                
                If GreatTotVol = ws.Cells(i, 12).Value Then
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                End If
                
            Next i
            
            'Print Greatest values
            ws.Range("Q2").Value = GreatIncr
            ws.Range("Q3").Value = GreatDecr
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Value = GreatTotVol
            

    Next ws

    
End Sub