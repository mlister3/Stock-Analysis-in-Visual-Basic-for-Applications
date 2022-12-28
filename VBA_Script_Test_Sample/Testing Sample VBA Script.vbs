Attribute VB_Name = "Module1"
Sub StockAnalyzer1():

Dim tickersym As String
Dim opening As Double
Dim closing As Double

'DeltaV will be the percent change
Dim DeltaV As Double

'TSV is total stock volume
Dim TSV As Variant

'establishing variables for logic
Dim startd As Integer
'sdlessone is first row for a stock
Dim sdlessone As Integer
Dim endrow As Integer
Dim lineindex As Integer
'establishing variables for greatest values outcome
Dim gincreaset As String
Dim gdecreaset As String
Dim GTV As String
Dim GIVal As Double
Dim GDVal As Double
Dim GTVVal As Variant

Dim sheetsplit As Worksheet

For Each sheetsplit In Worksheets

sheetsplit.Range("I1").Value = "Ticker"
sheetsplit.Range("J1").Value = "Yearly Change"
sheetsplit.Range("K1").Value = "Percent Change"
sheetsplit.Range("L1").Value = "Total Stock Volume"
sheetsplit.Range("P1").Value = "Ticker"
sheetsplit.Range("Q1").Value = "Value"
sheetsplit.Range("O2").Value = "Greatest % Increase"
sheetsplit.Range("O3").Value = "Greatest % Decrease"
sheetsplit.Range("O4").Value = "Greatest Total Volume"

'collects total number of rows in the dataset
endrow = sheetsplit.Range("A2").End(xlDown).Row
lineindex = 2
GIVal = 0
GDVal = 0
GTVVal = 0

For startd = 2 To endrow
    sdlessone = 2
    TSV = TSV + sheetsplit.Cells(startd, 7).Value
            
        'startd will be the last row for a stock
        If sheetsplit.Cells(startd + 1, 1).Value <> sheetsplit.Cells(startd, 1).Value Then
            
        'ticker collection and output
        tickersym = sheetsplit.Cells(startd, 1).Value
        sheetsplit.Cells(lineindex, 9).Value = tickersym
        
        'yearly change is calculated and output
        opening = sheetsplit.Cells(sdlessone, 3).Value
        closing = sheetsplit.Cells(startd, 6).Value
        sheetsplit.Cells(lineindex, 10).Value = closing - opening
            If sheetsplit.Cells(lineindex, 10).Value < "0" Then
                sheetsplit.Cells(lineindex, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf sheetsplit.Cells(lineindex, 10).Value > 0 Then
                sheetsplit.Cells(lineindex, 10).Interior.Color = RGB(0, 255, 0)
            End If
                      
        'percent change is calculated and output
        DeltaV = (closing - opening) / opening
        sheetsplit.Cells(lineindex, 11).Value = DeltaV
        sheetsplit.Cells(lineindex, 11).NumberFormat = "0.00%"
        
        'Total Stock Volume Dump
        sheetsplit.Cells(lineindex, 12).Value = TSV
        
            'stack of subfuctions to return greatest values and ticker
            If sheetsplit.Cells(lineindex, 11).Value > GIVal Then
            GIVal = sheetsplit.Cells(lineindex, 11).Value
            gincreaset = tickersym
            End If
            
            If sheetsplit.Cells(lineindex, 11).Value < GDVal Then
            GDVal = sheetsplit.Cells(lineindex, 11).Value
            gdecreaset = tickersym
            End If
            
            If sheetsplit.Cells(lineindex, 12).Value > GTVVal Then
            GTVVal = sheetsplit.Cells(lineindex, 12).Value
            GTV = tickersym
            End If
                      
        'moves onto the next stock
        lineindex = lineindex + 1
        sdlessone = startd + 1
        TSV = 0
        
        End If

Next startd

'print greatest stats
sheetsplit.Range("P2").Value = gincreaset
sheetsplit.Range("Q2").Value = GIVal
sheetsplit.Range("Q2").NumberFormat = "0.00%"
sheetsplit.Range("P3").Value = gdecreaset
sheetsplit.Range("Q3").Value = GDVal
sheetsplit.Range("Q3").NumberFormat = "0.00%"
sheetsplit.Range("P4").Value = GTV
sheetsplit.Range("Q4").Value = GTVVal
sheetsplit.Range("Q4").NumberFormat = "0"

sheetsplit.Columns("A:Q").AutoFit

Next sheetsplit

End Sub
