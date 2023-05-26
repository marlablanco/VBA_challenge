VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_Stock()

'Set up to loop all worksheets

'Define Variable
    Dim ws As Worksheet
    Dim Starting_WS As Worksheet
            Sheets(1).Select
        Set Starting_WS = ActiveSheet
        For Each ws In Sheets
            ws.Activate

'Define Variables

'Stock name
    Dim Ticker As String

'Variable for opening price
    Dim OpeningPrice As Variant

'Variable for closing price
    Dim ClosingPrice As Variant

'Variables for yearly change
    Dim YearlyChange As Double
    Dim PercentChange As Double

'Total stock volume
    Dim TotalStockVolume As Double
        TotalStockVolume = 0
    
'Track each stock volume summary
    Dim StockRow As Integer
        StockRow = 2
    
'Others
    Dim Final As String
         Final = Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Double
         i = 2

'Headers with titles
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
    
'Format headers
    Columns("A:Q").AutoFit
    
'Set up for loop
    OpeningPrice = Cells(i, 3).Value
        For i = 2 To Final
    
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        Else
                Ticker = Cells(i, 1).Value
        
'Print ticker name
        Range("I" & StockRow).Value = Ticker
        
'Add volume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        
'Closing price
        Cells(i + 1, 6).Select
        ActiveCell.Offset(-1, 0).Select
        ClosingPrice = ActiveCell.Value
        
'Yearly change
        YearlyChange = ClosingPrice - OpeningPrice
    
'To calculate percent change
        If OpeningPrice <> 0 Then
            PercentChange = YearlyChange / OpeningPrice
        Else
            PercentChange = 0
        End If
    
'Reset
        Cells(i + 1, 3).Select
        OpeningPrice = ActiveCell.Value
        
'Add to table
            Range("J" & StockRow).Value = YearlyChange
            Range("K" & StockRow).Value = PercentChange
            Range("K" & StockRow).NumberFormat = "0.00%"
            Range("L" & StockRow).Value = TotalStockVolume
            
'Move to next and reset
        StockRow = StockRow + 1
        TotalStockVolume = 0
        
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
    End If
        Next i
        
'Add percent changes and total vol to table

            Range("Q2").Value = WorksheetFunction.Max(Range("K:K"))
                Range("Q2").NumberFormat = "0.00%"
            Range("Q3").Value = WorksheetFunction.Min(Range("K:K"))
                FormatPercent (Range("Q2"))
                Range("Q3").NumberFormat = "0.00%"
            Range("Q4").Value = WorksheetFunction.Max(Range("L:L"))
                Range("Q4").NumberFormat = "##0.0E+0"
           
'Add format
     Columns("J").NumberFormat = "0.00000000"
     Columns("K").AutoFit
     Columns("L").AutoFit
     
'Add conditionals
        For i = 2 To Final
If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.Color = RGB(255, 0, 0)
End If
        Next i
        
'Ticker match
        For i = 2 To Final
    If Cells(i, 11).Value = Range("Q2").Value Then
            Range("P2").Value = Cells(i, 9).Value
    End If

    If Cells(i, 11).Value = Range("Q3").Value Then
            Range("P3").Value = Cells(i, 9).Value
    End If

    If Cells(i, 12).Value = Range("Q4").Value Then
        Range("P4").Value = Cells(i, 9).Value
End If
    Next i
    
'Format
    Columns("O").AutoFit
    Columns("Q").AutoFit
    
'Continue loop on next ws
    Next ws
        Starting_WS.Activate
    
End Sub

Sub Clear_WS_Data()
'Remove pre-filled in data to make room for data from current sheet

  For Each ws In Worksheets
  
    ws.Range("I:I").Clear
    ws.Range("J:J").Clear
    ws.Range("K:K").Clear
    ws.Range("L:L").Clear
    
    Next ws
    
End Sub

