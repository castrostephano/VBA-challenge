Sub Stock_vol()

    Dim wbk As Workbook
    Dim i As Long
    Dim Ticker_symbol As String
    Dim Stock_vol As Integer
    Dim J As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Per_Change As Double
    Dim Tot_Vol As Double
    Dim Tot2_Vol As Double
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    J = 2
    i = 1
    
    Open_Price = Cells(2, 3).Value
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
     'Find the Open price for a new ticker
            
            Open_Price = Cells(i + 1, 6).Value
            
            Columns("K").NumberFormat = "0.00%"
   
     For i = 2 To RowCount
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Tot_Vol = Tot_Vol + Cells(i, 7).Value
        
           
            'Find the Close Price for the current ticker
            
            Close_Price = Cells(i, 6).Value
            
            'Calculating yearly range by subtracting close price by open price
            
            Yearly_Change = Close_Price - Open_Price
            
            'Populating yearly change Cells in the summary table
            
            Cells(J, 10).Value = Yearly_Change
                
                If Yearly_Change > 0 Then
                
                    Cells(J, 10).Interior.Color = vbGreen
                
                ElseIf Yearly_Change < 0 Then
                
                    Cells(J, 10).Interior.Color = vbRed
                    
                End If
            
                'Calculate percent change
                
                If Open_Price = 0 And Open_Price = 0 Then
                
                    Per_Change = 0
                
                Else
            
                    Per_Change = (Yearly_Change / Open_Price)
                
                End If
                
            'Populate the Percent Change Cells
            
             Cells(J, 11).Value = Per_Change
        
            'Populate the ticker Cells
        
            Cells(J, 9).Value = Cells(i, 1)
            
            'Populate the Total Volume Cells
            
            Cells(J, 12).Value = Tot_Vol
            
            J = J + 1
            
            Tot_Vol = 0
            
            Else
            
             'Find the Total Volume
            
            Tot_Vol = Tot_Vol + Cells(i, 7).Value
   
        End If
    
    Next i
    
    Next ws
    

End Sub
