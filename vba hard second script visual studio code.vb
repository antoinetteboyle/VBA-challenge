Sub Stock_Count_Summary()

' Declare variables
Dim ws As Worksheet

Dim Ticker As String
Dim Volume_Total As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Movement As Double
Dim Percentage_Change As Double
Dim Summary_row_Counter As Double

Dim Max As Double
Dim Min As Double
Dim Vol As Double
Dim rwM As Double
Dim rwV As Double
Dim lastRow As Double

 For Each ws In ThisWorkbook.Worksheets
         
    ' Create headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    
    
    ' Initialise
    Open_Price = ws.Cells(2, 3).Value
    Summary_row_Counter = 2
    Volume_Total = 0
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
     
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
        Ticker = ws.Cells(i, 1).Value ' Grab the current ticker symbol in first column
        ws.Range("I" & Summary_row_Counter).Value = Ticker ' Print the current ticker symbol in summary column I
        
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value ' Grab the volume from 7th column
        ws.Range("L" & Summary_row_Counter).Value = Volume_Total  ' Print the current volume in summary column L
        
        
        Close_Price = ws.Cells(i, 6).Value ' Grab the current ticker symbol close price in F column
        Movement = Close_Price - Open_Price 'Calculate yearly movement in price
        ws.Range("J" & Summary_row_Counter).Value = Movement ' Print the current movement

          If Movement <> 0 And Open_Price <> 0 And Not IsNull(Open_Price) Then
            Percentage_Change = Movement / Open_Price 'Calculate percent movement in price
            ws.Range("K" & Summary_row_Counter).Value = FormatPercent(Percentage_Change, 2) ' Print the current percentage movement
          Else
            ws.Range("K" & Summary_row_Counter).Value = 0
          End If
          
          If Movement < 0 Then
           ws.Range("J" & Summary_row_Counter).Interior.Color = vbRed
           Else
           ws.Range("J" & Summary_row_Counter).Interior.Color = vbGreen
          End If
        
        Open_Price = ws.Cells(i + 1, 3).Value ' Reset the next ticker symbol open price in C column
        Summary_row_Counter = Summary_row_Counter + 1 ' increment summary count row by 1
        Volume_Total = 0 ' Reset the next ticker symbol volume
     Else
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
      
     
     End If
        
    Next i

        Min = 0
        Max = 0
        Vol = 1
        rwM = 2
        rwV = 2
    
     
        Do While ws.Cells(rwM, 11) <> ""
        
            If Max < ws.Cells(rwM, 11) Then
                Max = ws.Cells(rwM, 11)
                ws.Cells(2, 16) = FormatPercent(Max, 2) 'Max print
                ws.Cells(2, 15) = ws.Cells(rwM, 9) 'Ticker print
            End If
        
            If Min > ws.Cells(rwM, 11) Then
                Min = ws.Cells(rwM, 11)
                ws.Cells(3, 16) = FormatPercent(Min, 2)  'Min print
                ws.Cells(3, 15) = ws.Cells(rwM, 9) 'Ticker print
            End If
            
            rwM = rwM + 1
        Loop
        
        Do While ws.Cells(rwV, 12) <> ""
        
            If Vol < ws.Cells(rwV, 12) Then
                Vol = ws.Cells(rwV, 12)
                ws.Cells(4, 16) = Vol              'Max Vol print
                ws.Cells(4, 15) = ws.Cells(rwV, 9) 'Ticker print
            End If
           rwV = rwV + 1
        Loop
         

    
 Next ws

End Sub



