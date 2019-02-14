Sub AlphabeticalTest()
        For Each Ws In Worksheets
            
            Ws.Range("I1").Value = "Ticker"
            Ws.Range("J1").Value = "Total Stock Volume"
            
            
            Dim Ticker As String
            Dim LastRow As Long
            Dim TotalVol As Double
            Dim Count As Long
            Dim PreAmount As Long
            
                 
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
                Ws.Range("J" & Count).Value = TotalVol
                TotalVol = 0
                Count = Count + 1
                PreAmount = i + 1
                End If
            Next i
       Next Ws
    End Sub
    