Attribute VB_Name = "Module1"
Sub stocks():


Range("j1") = "Ticker"
Range("k1") = "Year Change"
Range("l1") = "Percent Change"
Range("m1") = "Total Volume"
  
 'For Each ws In Worksheets
 Dim sheet As Worksheet
 
 
 'variable to hold summary table
 summaryTableRow = 2
 
 
 For Each sheet In ThisWorkbook.Worksheets
 
 sheetname = sheet.Name
    
    'create integer for last row
    Dim lastRow As Double
    
    
    'variable to hold ticker
 ticker = ""
 
 'variable to hold first opening price
 Dim FirstOpen As Double
 
 FirstOpen = sheet.Cells(2, 3).Value
 
 
 'variable to hold last closing price
 Dim LastClose As Double
 'LastClose = sheet.Cells(2, 6).Value
   
 
 'variable to hold yearly change
 Dim yearlyChange As Double
 
 'yearlyChange = 0
 
 'variable to hold percent change
 Dim PercentChange As Double
 
 
 'variable to hold total stk volume
 Dim totalStkVol As Double
 totalStockVol = 0
  
 'use function to find last row in sheeet
 'lastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
 
        'MsgBox (sheetname)
 
  
 ' Find the last row of all(sheets
  lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
                
 'loop from row 2 to last row
  For Row = 2 To lastRow
  
 
    'MsgBox (lastRow)
 
     'look for ticker chage
  If sheet.Cells(Row + 1, 1).Value <> sheet.Cells(Row, 1) Then
     
     'grab ticker
     ticker = sheet.Cells(Row, 1).Value
          
     'grab the last close to calculate
     LastClose = sheet.Cells(Row, 6).Value
     
     'calculate year change
     yearlyChange = LastClose - FirstOpen
     Cells(summaryTableRow, 11).Value = yearlyChange
     
     'calculate percent change
     PercentChange = yearlyChange / FirstOpen
          
    'test to make sure correct values were captured
     'Cells(summaryTableRow, 15).Value = FirstOpen
     'Cells(summaryTableRow, 16).Value = LastClose
     
     Cells(summaryTableRow, 12).Value = PercentChange
        'Cells(summaryTableRow, 12).Style = "percent"
        ' trying a differnt approach for percent to add 2 decimal places
     Cells(summaryTableRow, 12).NumberFormat = "0.00%"
    
     'add last volume from row to totalStkVol
     totalStockVol = totalStockVol + sheet.Cells(Row, 7).Value
     
     'add ticker to col j
     Cells(summaryTableRow, 10).Value = ticker
     
     
       'add total charges to col h
     Cells(summaryTableRow, 13).Value = totalStockVol
         
        'reset the brand total to zero
    totalStockVol = 0
          
    'go to next summary table row
    summaryTableRow = summaryTableRow + 1
    
    
     FirstOpen = sheet.Cells(Row + 1, 3).Value
     
  
  Else
    'if brand stays same...
    'add onto total charges from c column
    totalStkVol = totalStkVol + Cells(Row, 7).Value
    
 
 
 End If
 
 
 
        If Cells(Row, 11) > 0 Then
        Cells(Row, 11).Interior.ColorIndex = 4
        
        ElseIf Cells(Row, 11) < 0 Then
        Cells(Row, 11).Interior.ColorIndex = 3
        
        End If
        
        'variable for ADitional Ticker (ticker-Bonus)
        'variable for grtest increase
        'variable for greatest decrease
        'variable for greates total volume
        
        
       Dim TickerB_inc As String
       Dim TickerB_dec As String
       Dim TickerB__vol As String
       Dim GreatestIncrease As Double
       Dim GreatestDecrease As Double
       Dim GreatestTotVol As Double
       
       If sheet.Cells(Row, 11) > GreatestIncrease Then
       
       GreatestIncrease = sheet.Cells(Row, 11)
       TickerB_inc = sheet.Cells(Row, 10)
    
       Else
       End If
       
        If sheet.Cells(Row, 11) < GreatestDecrease Then
       
       GreatestDecrease = sheet.Cells(Row, 11)
       TickerB_dec = sheet.Cells(Row, 10)
                     
       Else
       End If
       
        If sheet.Cells(Row, 13) > GreatestTotVol Then
       
       GreatestTotVol = sheet.Cells(Row, 13)
       TickerB__vol = sheet.Cells(Row, 10)
    
       Else
       End If
       
           
        
        
Next Row



Next sheet

Range("o2:p2").Merge
Range("o3:p3").Merge
Range("o4:p4").Merge

Range("o2") = "Gretest % Increase"
Range("o3") = "Gretest % Decrease"
Range("o4") = "Greatest Total Volume"
Range("q1") = "Ticker"
Range("r1") = "Value"

Range("R2").Value = GreatestIncrease
Range("R3").Value = GreatestDecrease
Range("R4").Value = GreatestTotVol
Range("Q2").Value = TickerB_inc
Range("Q3").Value = TickerB_dec
Range("Q4").Value = TickerB__vol


  
 

End Sub





