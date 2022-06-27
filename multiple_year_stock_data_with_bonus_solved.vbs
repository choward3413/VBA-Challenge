Attribute VB_Name = "Module1"
Sub stocks():


Range("j1") = "Ticker"
Range("k1") = "Year Change"
Range("l1") = "Percent Change"
Range("m1") = "Total Volume"
  

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
 Dim totalStockVol As Double
 totalStockVol = 0
  
 'use function to find last row in sheeet
 'lastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
 
        MsgBox (sheetname)
 
  
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
     sheet.Cells(summaryTableRow, 11).Value = yearlyChange
     sheet.Cells(summaryTableRow, 11).NumberFormat = "$0.00"
     
     'calculate percent change
     PercentChange = yearlyChange / FirstOpen
          
    'test to make sure correct values were captured
     'Cells(summaryTableRow, 15).Value = FirstOpen
     'Cells(summaryTableRow, 16).Value = LastClose
     
     sheet.Cells(summaryTableRow, 12).Value = PercentChange
        'Cells(summaryTableRow, 12).Style = "percent"
        ' trying a differnt approach for percent to add 2 decimal places
     sheet.Cells(summaryTableRow, 12).NumberFormat = "0.00%"
    
     'add last volume from row to totalStkVol
     totalStockVol = totalStockVol + sheet.Cells(Row, 7).Value
     
     'add ticker to col j
     sheet.Cells(summaryTableRow, 10).Value = ticker
     
     
       'add total charges to col h
     sheet.Cells(summaryTableRow, 13).Value = totalStockVol
         
        'reset the brand total to zero
    totalStockVol = 0
          
    'go to next summary table row
   summaryTableRow = summaryTableRow + 1
        
     FirstOpen = sheet.Cells(Row + 1, 3).Value
     
  
  Else
    'if brand stays same...
    'add onto total charges from c column
    totalStockVol = totalStockVol + sheet.Cells(Row, 7).Value
    
 
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
       
       
       If Cells(Row, 12) > GreatestIncrease Then
       
       GreatestIncrease = Cells(Row, 12)
       TickerB_inc = Cells(Row, 10)
       
       End If
       
        If Cells(Row, 12) < GreatestDecrease Then
       
       GreatestDecrease = Cells(Row, 12)
       TickerB_dec = Cells(Row, 10)
                     
       Else
       End If
       
        If Cells(Row, 13) > GreatestTotVol Then
       
       GreatestTotVol = Cells(Row, 13)
       TickerB__vol = Cells(Row, 10)
    
       Else
        End If
         Next Row
        
Range("o2:p2").Merge
Range("o3:p3").Merge
Range("o4:p4").Merge

sheet.Cells(2, 15).Value = "Gretest % Increase"
sheet.Cells(3, 15).Value = "Gretest % Decrease"
sheet.Cells(4, 15).Value = "Greatest Total Volume"
sheet.Cells(1, 17) = "Ticker"
sheet.Cells(1, 18) = "Value"

sheet.Cells(2, 18).Value = GreatestIncrease
sheet.Cells(2, 18).NumberFormat = "0.00%"
sheet.Cells(3, 18).Value = GreatestDecrease
sheet.Cells(3, 18).NumberFormat = "0.00%"
sheet.Cells(4, 18).Value = GreatestTotVol
sheet.Cells(2, 17).Value = TickerB_inc
sheet.Cells(3, 17).Value = TickerB_dec
sheet.Cells(4, 17).Value = TickerB__vol


 
      Next sheet
 

End Sub





