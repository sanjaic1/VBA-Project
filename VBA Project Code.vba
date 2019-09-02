Sub StockVol()

    ' Delcare the variables
    Dim ticker As String

    Dim current As Worksheet
    Dim lastrow, i, j As Long
    Dim volcnt As Double
    Dim tickcol, volcol, rowtowrite, coltowrite, highcol, lowcol As Integer
    
    Dim perchange, openprice, closeprice As Single
    Dim newstock As Boolean
    Dim yrlychange, yrlychangperc As Single
  
    'Hard challenge variables
    Dim highinc, highdec, highvol As Single
    Dim hightick, lowtick, voltick As String
    Dim pctCompl As Integer
    
    ticker = ""
    tickcol = 1                 ' Ticker is in column 1
    volcol = 7                  ' Volume data is in column 7
    highcol = 4
    lowcol = 5
    
    volcnt = 0
    openprice = 0
    closeprice = 0
    newstock = True
    yrlychange = 0
    yrlychangeperc = 0
    
        'Loop through all the worksheets in this file
        For Each current In Worksheets

            ' Reset the volume sum, low and high counters
            newstock = True
            closeprice = 0
            yrlychange = 0
            volcnt = 0
            
            ' Lastrow of data for this worksheet
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
             
            ' Where to start writing the data
            rowtowrite = 2    'after header
            coltowrite = 9
             
            'Activate the current worksheet
            current.Activate
             
            ' Write the headers for the new rows
            Cells(1, coltowrite) = "Ticker"
            Cells(1, coltowrite + 1) = "Year chg"
            Cells(1, coltowrite + 2) = "% chg"
            Cells(1, coltowrite + 3) = "Trading Vol"
             
            ' Loop through all the rows of data starting after the headers
            For i = 2 To lastrow
          
                ' Add the volume
                volcnt = volcnt + Cells(i, volcol)
                
                ' Get open price is this is first row in new stock
                If newstock Then
                    openprice = Cells(i, 3)
                    newstock = False
                End If
                
                ' If its a different ticker
                If Cells(i + 1, tickcol).Value <> Cells(i, tickcol).Value Then
                    ' Write the ticker
                    ticker = Cells(i, tickcol)
                    Cells(rowtowrite, coltowrite) = ticker
                    
                    'Write the yearly change (opening price - closeprice)
                    closeprice = Cells(i, 6)
                    yrlychange = closeprice - openprice
                    Cells(rowtowrite, coltowrite + 1) = yrlychange
                    Cells(rowtowrite, coltowrite + 1).NumberFormat = ".0000"
                    
                    If yrlychange > 0 Then
                        Cells(rowtowrite, coltowrite + 1).Font.Color = RGB(34, 139, 34)
                        Cells(rowtowrite, coltowrite + 2).Interior.ColorIndex = 4                   'Green
                    Else
                        Cells(rowtowrite, coltowrite + 1).Font.Color = RGB(255, 0, 0)
                        Cells(rowtowrite, coltowrite + 2).Interior.ColorIndex = 3                   'Red
                    End If
                
                    'Write the yearly % change
                    If yrlychange <> 0 Then
                        ' If the stock has openprice of 0 was causing divide by 0 error.  This is a temp solution but really should track the open from first non zero open price since it could be a new stock
                        If openprice <> 0 Then
                            yrlychangeperc = (yrlychange / openprice)
                            Cells(rowtowrite, coltowrite + 2).Value = yrlychangeperc
                            Cells(rowtowrite, coltowrite + 2).NumberFormat = "0.00%"
                        Else
                            Cells(rowtowrite, coltowrite + 2).Value = ""
                            Cells(rowtowrite, coltowrite + 2).Interior.ColorIndex = 6               ' Yellow
                        End If
                    End If
                                    
                    ' Write the volume
                    Cells(rowtowrite, coltowrite + 3) = volcnt
                    Cells(rowtowrite, coltowrite + 3).NumberFormat = "#,##0"
                    
                    ' Increment the next row to write to
                    rowtowrite = rowtowrite + 1
                    
                    ' Reset the volume sum, low and high counters
                    newstock = True
                    closeprice = 0
                    yrlychange = 0
                    volcnt = 0
                    
                    ' Display progress of processing the data
                    pctCompl = (i / lastrow) * 100
                    Application.StatusBar = "Processing data... " & pctCompl & "% Completed"
                End If
            Next i
          
            ' Now loop through the newly created rows to get the largest %  increase, decrease and volume
            ' Get last row for this my specific column
            lastrow = Cells(Rows.Count, 10).End(xlUp).Row
            
            ' Initialize variables
            highinc = 0
            lowdec = 0
            highvol = 0
          
            For j = 2 To lastrow
          
            ' Store highest increase, decrease and volume information
            If Cells(j, coltowrite + 2) > highinc Then
                highinc = Cells(j, coltowrite + 2)
                hightick = Cells(j, coltowrite)
            End If
            
            If Cells(j, coltowrite + 2) < highdec Then
                highdec = Cells(j, coltowrite + 2)
                lowtick = Cells(j, coltowrite)
            End If
           
            If Cells(j, coltowrite + 3) > highvol Then
                highvol = Cells(j, coltowrite + 3)
                voltick = Cells(j, coltowrite)
            End If
        Next j
        
        'Write out all the summary information
        Cells(2, coltowrite + 5) = "Greatest % Increase"
        Cells(3, coltowrite + 5) = "Greatest % decrease"
        Cells(4, coltowrite + 5) = "Highest trading volume"
        Columns(coltowrite + 5).EntireColumn.AutoFit
        
        Cells(1, coltowrite + 6) = "Ticker"
        Cells(1, coltowrite + 6).Font.Underline = True
        Cells(1, coltowrite + 6).Font.Bold = True
        
        Cells(2, coltowrite + 6) = hightick
        Cells(3, coltowrite + 6) = lowtick
        Cells(4, coltowrite + 6) = voltick
    
        Cells(1, coltowrite + 7) = "Value"
        Cells(1, coltowrite + 7).Font.Underline = True
        Cells(1, coltowrite + 7).Font.Bold = True
        
        Cells(2, coltowrite + 7) = highinc
        Cells(2, coltowrite + 7).NumberFormat = "0.00%"
        Cells(2, coltowrite + 7).Interior.ColorIndex = 4
        
        Cells(3, coltowrite + 7) = highdec
        Cells(3, coltowrite + 7).NumberFormat = "0.00%"
        Cells(3, coltowrite + 7).Interior.ColorIndex = 3
               
        Cells(4, coltowrite + 7) = highvol
        Cells(4, coltowrite + 7).NumberFormat = "000,000,000"
        
    Next                    'worksheet

    ' Clear the progress status bar
    Application.StatusBar = False

End Sub

' Done  Leading zeros
' Done Format finals with colors and/or lines
' Done Color code the hard vars red green

' Done Combine the Dims and As for less lines
' Done Comment the code well
' Done Fix the indenting
' Done Check totals manually by highlighting and against the images

' And need to run on big data set DONT FORGET
' Create deliverables - images etc.
' Create repo, upload and give link

