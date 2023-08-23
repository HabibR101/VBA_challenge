Sub stock_Mrkt()

'Declare variables for the script below

    Dim ws As Worksheet
    Dim tickerSymbol As String
    Dim rowNumber As Integer
    Dim openingSt As Double
    Dim closingEn As Double
    Dim diff As Double
    Dim percentage As Double
    Dim totalStk As LongLong
    
    'Sets loop for worksheets
    
        For Each ws In Worksheets
        
        'Sets row Number to 2 to omit first row
        
        rowNumber = 2
        
        'Sets total Stock to zero
        totalStk = 0
        
        'Sets the opening stock to the value located in row two column 3 or B2 <open> column
        
        openingSt = ws.Cells(2, 3).Value
        
        'sets the Variabel as lastrow and enables return the total count of rows in the worksheet
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Names the ranges J1,K1,L1,M1 cells as the valeu in quotation marks
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percentage Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        'Initiates the For loop to loop through rows using the last row assigned figure
        
            For i = 2 To lastRow
                
                'Checks if the cells are not equal to each other and then complete the next action
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                
                    'assign value to variable tickersymbol
                    tickerSymbol = ws.Cells(i, 1).Value
                        
                    'Uses Column J and row number to increase row number in column j and change ticker symbol value
                    
                    ws.Range("J" & rowNumber).Value = tickerSymbol
                    
                    
                    'assign value to variable Closing end of year stock price
                    closingEn = ws.Cells(i, 6).Value
                    
                    'finds difference in the opening and closing stock price
                    diff = closingEn - openingSt
                    
                    'Uses loop to add value into yearly change found in Diff variable
                    
                    ws.Range("K" & rowNumber).Value = diff
                    
                    
                    'finds the precenatge by taking values in closing and opening compelteing the recenatge calcualtion
                    percentage = ((closingEn - openingSt) / (openingSt))
                    
                    'Uses loop to add value into yearly change found in precenatge variable and changes format
                    ws.Range("L" & rowNumber).Value = percentage
                    ws.Range("L" & rowNumber).NumberFormat = "0.00%"
                    
                    
                    'if loop Conditional formatting to change colour in yearly change
                        If diff > 0 Then
                        
                            ws.Range("K" & rowNumber).Interior.ColorIndex = 4
                        
                        Else
                        
                            ws.Range("K" & rowNumber).Interior.ColorIndex = 3
                            
                        End If
                        
                    'assign value to variable Opening stock of year stock price
                    openingSt = ws.Cells(i + 1, 3).Value
                    'assign value to totla stock amount
                    totalStk = totalStk + ws.Cells(i, 7).Value
                    'add total stock to columns  m by row
                    ws.Range("M" & rowNumber).Value = totalStk
                    
                    'rownumber increase
                    rowNumber = rowNumber + 1
                    
                    'total stock value
                    totalStk = 0
                    
                    
                Else
                
                
                    totalStk = totalStk + ws.Cells(i, 7).Value
                    
                    
                End If
                
            
            Next i
            
            'assigns values to the below cells and changes format to precenatge
            
                ws.Cells(2, 15).Value = "Greatest % Increase"
                ws.Cells(3, 15).Value = "Greatest % Decrease"
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 17).Value = "Value"
                ws.Range("Q2:Q3").NumberFormat = "0.00%"
                
                
                'assigns values to the below cells as max, min precenatge, Max volume and changes format to precenatge
                
                maxPercentage = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow + 1))
           
                minPercentage = Application.WorksheetFunction.Min(ws.Range("L2:L" & lastRow + 1))
    
                maxVolume = Application.WorksheetFunction.Max(ws.Range("M2:M" & lastRow + 1))
              
                ws.Cells(2, 17).Value = maxPercentage
                
                ws.Cells(3, 17).Value = minPercentage
                
                ws.Cells(4, 17).Value = maxVolume
                
                'Starts loop to match ticker symbol to the value found above for each max,min precentage and Max volume.
                'inserts into cells P1,P2 and P3 when found.
                For i = 2 To lastRow
                
                    If ws.Cells(i, 12).Value = maxPercentage Then
                        ws.Cells(2, 16).Value = ws.Cells(i, 10).Value
                    End If
                    If ws.Cells(i, 12).Value = minPercentage Then
                        ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
                    End If
                    If ws.Cells(i, 13).Value = maxVolume Then
                        ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
                    End If
                    
                Next i
                
                
                'formatting to autofit columns to cell lenght
                ws.Columns("I:Q").AutoFit
               
                 
                    
            
              'starts next worksheet loop.
            Next ws
            
                                 


End Sub