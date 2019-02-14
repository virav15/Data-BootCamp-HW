Sub AlphabeticalTest()
        For Each Ws In Worksheets
            
            Ws.Range("I1").Value = "Ticker"
            Ws.Range("L1").Value = "Total Stock Volume"
            Ws.Range("J1").Value = "Yearly Change"
            Ws.Range("K1").Value = "Percent Change"
           
            
            Dim Ticker As String
            Dim LastRow As Long
            Dim TotalVol As Double
            Dim Count As Long
            Dim PreAmount As Long
            Dim YearlyChange As Double
            Dim YearlyOpen As Double
            Dim YearlyClose As Double
            Dim PercentChange As Double
           
                 
            TotalVol = 0
            Count = 2
            
            PreAmount = 2
            
            LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                
                TotalVol = TotalVol + Ws.Cells(i, 7).Value
                
                If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                Ticker = Ws.Cells(i, 1).Value
                        
                'Ticker symbol
                Ws.Range("I" & Count).Value = Ticker
                
                'Total Stock Volume
                Ws.Range("L" & Count).Value = TotalVol
                TotalVol = 0
                
                'Yearly Change
                YearlyOpen = Ws.Range("C" & PreAmount)
                YearlyClose = Ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                Ws.Range("J" & Count).Value = YearlyChange
                
                'Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = Ws.Range("C" & PreAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                Ws.Range("K" & Count).NumberFormat = "0.00%"
                Ws.Range("K" & Count).Value = PercentChange
                
                'Conditional highlight positive/negative
                    If Ws.Range("J" & Count).Value >= 0 Then
                        Ws.Range("J" & Count).Interior.ColorIndex = 4
                    Else
                        Ws.Range("J" & Count).Interior.ColorIndex = 3
                    End If
                
                Count = Count + 1
                PreAmount = i + 1
                End If
            Next i
            
           

        Next Ws
    End Sub
    


