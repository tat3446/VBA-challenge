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
  
        'Identify opening price of year
        'OpenPrice = Cells(i, 3).Value
    
        'Loop through stocks
        For i = 2 To LastRow
         
            'Check if ticket symbol changed
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Identify ticker symbol
            TixSym = Cells(i, 1).Value
            
            'Identify opening price of year
            OpenPrice = Cells(2, 3).Value
            
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
                        
            'Calculate Yearly Change
            PercentChg = (ClosePrice - OpenPrice / OpenPrice)
                    
            'Print Percent Change
            'Range("K" & TixSym_SummTbl_Row).Value = ClosePrice
            ws.Range("K" & TixSym_SummTbl_Row).Value = PercentChg
                        
            'TEST PRINT OPEN CLOSE numbers
            'ws.Cells(i, "N").Value = OpenPrice
            'ws.Cells(i, "O").Value = ClosePrice
                        
            'Print total volume of stock
            ws.Range("L" & TixSym_SummTbl_Row).Value = TixSym_Total
            
            'Add one to row to summary
            TixSym_SummTbl_Row = TixSym_SummTbl_Row + 1
            
            'Reset volume
            TixSym_Total = 0
            
            
            
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
            
          ' Worksheet.Columns("A:R").AutoFit
                'Format columns
                'ws.Cells("J2:J").NumberFormat = "0.00"
                'ws.Cells("K2:K").NumberFormat = "0.00%"                               
                            
    Next ws

    
End Sub