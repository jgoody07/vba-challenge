Attribute VB_Name = "Module1"
Sub Stock_Tickers()

'Set Variables
Dim Ticker_Symbol As String
Dim Summary_Table_Row As Integer
Dim first_open, last_close As Long
    first_open = 0
    last_close = 0
Dim Ticker_Volume_Total As Single
    Ticker_Volume_Total = 0
Dim Yearly_Change As Double
    Yearly_Change = 0
Dim Percent_Change As String
    Percent_Change = 0
 ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
 ' --------------------------------------------
     
    For Each ws In Worksheets
     last_close = 2
     
       'Keep track of the location for each Ticker in the summary table
        Summary_Table_Row = 2
        
        'Loop through all Ticker data
           'Determine the last row
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
   '--------------------------------------------
   'Check if we are still within the same Ticker Symbol, if we are not...
   '--------------------------------------------
         'Store the Ticker Symbol and Print to Summary Table
          Ticker_Symbol = ws.Cells(i, 1).Value
          ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
          
           'Calculate total Stock Volume
           Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value
           'Adding Ticker_Volume_Total to summary chart
           ws.Range("M" & Summary_Table_Row).Value = Ticker_Volume_Total
         
         'Calculate year change
          Year_Change = ws.Range("F" & i).Value - ws.Range("C" & last_close).Value
          'Print in Summary Table
          ws.Range("K" & Summary_Table_Row).Value = Year_Change
          
          'Calculate Percentage Yearly Change and Print to Summary Table
           If ws.Range("C" & last_close).Value <> 0 Then
           Percent_Change = 100 * (Year_Change / ws.Range("C" & last_close).Value)
           ws.Range("L" & Summary_Table_Row).Value = Percent_Change
            End If
                          
          'Reset Data Calculations
             Yearly_Change = 0
             Percent_Change = 0
             Ticker_Volume_Total = 0
            
          'Add one to the Summary Table
             Summary_Table_Row = Summary_Table_Row + 1
             last_close = i + 1
        
        Else
        
        'Add total Stock Volume
        Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value
                        
        End If
     
     Next i
 '--------------------------------------------
   'FORMAT POSITIVE/NEGATIVE PERCENT CHANGE
 '--------------------------------------------
    lastrow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
    For i = 1 To lastrow2
        
        If ws.Cells(i, 11).Value < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
        Else
        ws.Cells(i, 11).Interior.ColorIndex = 4
        
        End If
    
    Next i
 ' --------------------------------------------
    'Calculate "Greatest % increase", "Greatest % decrease", and "Greatest Total Volume"
 ' --------------------------------------------
  'Calculating Greatest % increase and Greatest % Decrease
    
    'Set variables
    Dim xmax As Double
    Dim xmin As Double
    Dim r As Range
        
    'Run Function
     RowCount = ws.Cells(Rows.Count, 12).End(xlUp).Row
     Set r = ws.Range("L2:L" & RowCount)
     xmin = Application.WorksheetFunction.Min(r)
     xmax = Application.WorksheetFunction.Max(r)
        
    'Send to summary
     ws.Cells(3, 16).Value = xmin
     ws.Cells(2, 16).Value = xmax
 
 'Calculating Greatest Total Volume
    
    'Set Variables
    Dim Vmax As Double
    Dim z As Range
        
    'Run Function
    RowCount = ws.Cells(Rows.Count, 13).End(xlUp).Row
    Set z = ws.Range("M2:M" & Rows.Count)
    Vmax = Application.WorksheetFunction.Max(z)
        
     'Send to summary
     ws.Cells(4, 16).Value = Vmax
        
 '--------------------------------------------
    'Label Summary Table Headers
 '--------------------------------------------
    ws.Cells(1, 10).Value = "Ticker Symbol"
    ws.Cells(1, 11).Value = "Year Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
 
 Next ws

End Sub

