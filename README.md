# VBA-challenge
Peer Collaboration with Ryan Himes
Application.WorksheetFunction.Max(Range(Cells(m, 1),Cells(n, 1))) - Max and Min Formulas Cited from StackOverflow
  //stackoverflow.com/questions/31906571/excel-vba-find-maximum-value-in-range-on-specific-sheet
VBA Code Below  
Sub Ticker()
    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
    
    ' Set-up Worksheet
    Dim WorksheetName As String
    WorksheetName = ws.Name
    
    
    ' Set an initial variable for holding the Ticker, Open and Close Price, Stock Volume, Yearly Change, Percent Change, Greatest Increase/Decrease and Stock Volume
    Dim Ticker As String
    Dim Open_Price, Close_Price As Integer
    Dim Stock_Volume, GreatestVolume As LongLong
    Dim Yearly_Change, Percent_Change, GreatestIncrease, GreatestDecrease As Double
    
    
    ' Set Original Values for Open Price, Close Price, Stock Volume, and Yearly Change
    Open_Price = ws.Cells(2, 3).Value
    Close_Price = 6
    Stock_Volume = 0
    Yearly_Change = 0
    
    ' Create and keep track of Summary Table 1 for Tickers Yearly Change
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Columns("K").NumberFormat = "0.00%"
        
    ' Create Summary Table 2 for Greatest % Increase, Decrease and Stock Volume
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
     ' Determine the Last Row in the Sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all ticker symbols
    For i = 2 To LastRow

    ' Check if we are still within the ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value

      ' Calculate Yearly Change, Percent Change and Total Stock Volume
      Yearly_Change = ws.Cells(i, Close_Price).Value - Open_Price
      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
      Percent_Change = Yearly_Change / Open_Price
      
      ' Print the Ticker, Yearly Change, Percent Change And Total Stock Volume in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
      ' Conditional Format cell to Green for Positive Yearly Change Values and Red for Negative Yearly Change Values
        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Yearly Change, Percent Change andTotal Stock Volume
      Yearly_Change = 0
      Percent_Change = 0
      Stock_Volume = 0
      
      'Change Open Price to the next ticker opening price
      Open_Price = ws.Cells(i + 1, 3).Value
    
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Stock Volume
      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
      
    End If

  Next i
    
    'Define Greatest % Increase, % Decrease and Total Volume using Min and Max Formulas
    GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
    GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
    GreatestVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("Q2", "Q3").NumberFormat = "0.00%"
    
    'Loop through Summary Table 1 Ticker Yearly Change and Total Volume
    For i = 2 To LastRow
    
    'Find the Greatest % Increase and pull Ticker and Value into Summary Table 2
        If ws.Cells(i, 11).Value = GreatestIncrease Then
            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            
    'Find the Greatest % Decrease and pull Ticker and Value into Summary Table 2
        ElseIf ws.Cells(i, 11).Value = GreatestDecrease Then
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("P3").Value = ws.Cells(i, 9).Value
         End If
         
    'Find the Greatest Total Volume and pull Ticker and Value into Summary Table 2
         If ws.Cells(i, 12).Value = GreatestVolume Then
            ws.Range("Q4").Value = GreatestVolume
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
        
    Next i
    
    'Make Columns readable by auto fitting columns
    ws.Columns("J").AutoFit
    ws.Columns("K").AutoFit
    ws.Columns("L").AutoFit
    ws.Columns("I").AutoFit
    ws.Columns("O").AutoFit
    ws.Columns("P").AutoFit
    ws.Columns("Q").AutoFit
    
 Next ws
 
End Sub
