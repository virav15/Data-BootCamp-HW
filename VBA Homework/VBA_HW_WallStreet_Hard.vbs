Sub AlphabeticalTest()
        For Each Ws In Worksheets
            
            Ws.Range("I1").Value = "Ticker"
            Ws.Range("L1").Value = "Total Stock Volume"
            Ws.Range("J1").Value = "Yearly Change"
            Ws.Range("K1").Value = "Percent Change"
            Ws.Range("P1").Value = "Ticker"
            Ws.Range("Q1").Value = "Value"
            Ws.Range("O2").Value = "Greatest % Increase"
            Ws.Range("O3").Value = "Greatest % Decrease"
            Ws.Range("O4").Value = "Greatest Total Volume"
            
            Dim Ticker As String
            Dim LastRow As Long
            Dim TotalVol As Double
            Dim Count As Long
            Dim PreAmount As Long
            Dim YearlyChange As Double
            Dim YearlyOpen As Double
            Dim YearlyClose As Double
            Dim PercentChange As Double
            Dim IncreaseVal As Double
            Dim DecreaseVal As Double
            Dim LastRowofVal As Long
            Dim GreatestVal As Double
                 
            TotalVol = 0
            Count = 2
            IncreaseVal = 0
            DecreaseVal = 0
            GreatestVal = 0
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
            
            'Find the greatest value
            LastRowofVal = Ws.Cells(Rows.Count, 11).End(xlUp).Row
            Ws.Range("Q2").NumberFormat = "0.00%"
            Ws.Range("Q3").NumberFormat = "0.00%"
            
            For j = 2 To LastRowofVal
                If Ws.Range("K" & j).Value > IncreaseVal Then
                    IncreaseVal = Ws.Range("K" & j).Value
                    Ws.Range("Q2").Value = IncreaseVal
                    Ws.Range("P2").Value = Ws.Range("I" & j).Value
                End If
                If Ws.Range("K" & j).Value < DecreaseVal Then
                    DecreaseVal = Ws.Range("K" & j).Value
                    Ws.Range("Q3").Value = DecreaseVal
                    Ws.Range("P3").Value = Ws.Range("I" & j).Value
                End If
                If Ws.Range("L" & j).Value > GreatestVal Then
                    GreatestVal = Ws.Range("L" & j).Value
                    Ws.Range("Q4").Value = GreatestVal
                    Ws.Range("P4").Value = Ws.Range("I" & j).Value
                End If
            Next j
    

        Next Ws
    End Sub
    


