Attribute VB_Name = "Module2"
Sub Testing()

    For Each ws In Worksheets

        Dim WorksheetName As String

        Dim TableRow As Double
        TableRow = 2

        Dim Ticker As String

        Dim OpenPrice As Double

        Dim ClosePrice As Double

        Dim YearlyChange As Double

        Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        WorksheetName = ws.Name
    

        ' To display headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    

        ' To loop through our data set

        For i = 2 To LastRow
       
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & TableRow).Value = Ticker
            
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                ws.Range("L" & TableRow).Value = TotalVolume
            
                ' Sets close price
            
                ClosePrice = ws.Cells(i, 6).Value
            
       
                ' Fills Yearly Change column with correct value
            
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("J" & TableRow).Value = YearlyChange
            
            
                ' Colors negative values red and positive values green
            
                If YearlyChange < 0 Then
            
                    ws.Range("J" & TableRow).Interior.ColorIndex = 3
                
                ElseIf YearlyChange > 0 Then
            
                    ws.Range("J" & TableRow).Interior.ColorIndex = 4
                
                End If
            
                ' Fills the Percent Change Column
            
                If OpenPrice = 0 Then
            
                    PercentChange = 0
                
                Else
                
                    PercentChange = YearlyChange / OpenPrice ' Calculates Percentage
                
                    ws.Range("K" & TableRow).Value = PercentChange ' Fills the proper cell
                
                    ws.Range("K" & TableRow).NumberFormat = "0.00%" ' Changes to proper style
                
                End If
            
            Table = Tabe + 1 ' Increments Table
            
            TableRow = TableRow + 1 ' Increments Table Row
            
            TotalVolume = 0 ' Resets total volume
        
            Else
        
         ' Makes sure that the Opening price is only read on the first instance for a given ticker
    
                If TotalVolume = 0 Then
            
                    OpenPrice = ws.Cells(i, 3).Value
                
                End If
            
                ' Increments Total Volume
            
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        
            End If
           
        Next i
    
    
     'Bonus Material
 
        Dim MaxPercent As Double
        Dim MinPercent As Double
        Dim MaxVolume As Double
 
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
 
        MaxPercent = ws.Application.WorksheetFunction.Max(ws.Range("K2:K3001"))
        MinPercent = ws.Application.WorksheetFunction.Min(ws.Range("K2:K3001"))
        MaxVolume = ws.Application.WorksheetFunction.Max(ws.Range("L2:L3001"))
        ws.Cells(2, 17).Value = MaxPercent
        ws.Cells(3, 17).Value = MinPercent
        ws.Cells(4, 17).Value = MaxVolume
 
        For i = 2 To LastRow
 
            If ws.Cells(i, 11).Value = MaxPercent Then
            
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value = MinPercent Then
            
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
            End If
    
        Next i

        For i = 2 To LastRow

            If ws.Cells(i, 12).Value = MaxVolume Then
            
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            End If
    
        Next i
    
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
    Next ws


End Sub
