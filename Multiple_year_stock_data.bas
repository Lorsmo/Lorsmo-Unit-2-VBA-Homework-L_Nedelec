Attribute VB_Name = "Module1"
Option Explicit

Sub Stock_Data()

    ' Set an initial variable for the worksheets
    Dim ws As Worksheet

    'Loop through all sheets
    For Each ws In Worksheets

        ' Create a variable to hold the row counter
        Dim i As Long

        ' Set an initial variable for the ticker's name
        Dim TickerName As String

        ' Set an initial variable for the opening price
        ' Define the first opening price of each worksheet
        Dim TickerOpen As Double
        TickerOpen = ws.Range("C2").Value

        ' Set an initial variable for the closing price
        Dim TickerClose As Double

        ' Set an initial variable for holding the percent change from the opening price to the closing price
        Dim YearlyChange As Double
        YearlyChange = 0

        ' Set an initial variable for holding the yearly percentage change
        Dim PercentChange As Double
        PercentChange = 0

        ' Keep track of the location for ticker name in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Set an initial variable for holding the volume each stock had over year
        Dim Volume As Double
        Volume = 0

        'Find the last row of stock data
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all stock data
        For i = 2 To LastRow

            ' Check if we are still within the same ticker
            ' If the ticker name change,
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the ticker name and print it
                TickerName = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = TickerName
                
                'Set the closing price
                TickerClose = ws.Cells(i, 6).Value

                ' Add to the volume
                Volume = Volume + ws.Cells(i, 7).Value
                
                ' Calculate the yearly changing price and print it
                YearlyChange = TickerClose - TickerOpen
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange

                ' Check if the yearly change value is <0
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                    
                    ' Print the cell in red
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

                ' If the the yearly change value is >0
                Else

                    ' Print the cell in green
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                End If

                ' Check if the opening price is not 0 to avoid division error by 0
                If TickerOpen <> 0 Then
                    
                    'Calculate the percent change
                    PercentChange = YearlyChange / TickerOpen
                
                '  Check if the opening price = 0 and the yearly change = 0
                ElseIf TickerOpen = 0 And YearlyChange = 0 Then
                    
                    ' The percent change = 0 because no evolution
                     PercentChange = 0
                    
                End If

                ' Print the percent change
                ws.Range("K" & Summary_Table_Row).Value = PercentChange

                ' Print the volume
                ws.Range("L" & Summary_Table_Row).Value = Volume
                
                ' Set the new ticker name
                TickerOpen = ws.Cells(i + 1, 3).Value
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the volume
                Volume = 0
            
            ' If the cell immediately following a row is the same ticker
            Else
                
                ' Add to the volume
                Volume = Volume + ws.Cells(i, 7).Value

            End If

        Next i

        ' Print the summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Remove one to the summary table row
        Summary_Table_Row = Summary_Table_Row - 1

        ' Format the summary table
        ws.Range("J2:J" & Summary_Table_Row).NumberFormat = "0.00000000"
        ws.Range("K2:K" & Summary_Table_Row).Style = "  Percent"
        ws.Range("K2:K" & Summary_Table_Row).NumberFormat = "0.00%"
        ws.Range("I1:L" & Summary_Table_Row).BorderAround Weight:=xlMedium
        ws.Range("I1:L1").BorderAround Weight:=xlMedium
        ws.Range("I1:L1").Font.Bold = True
        
        'Set an initial variable for the greatest percentage increase
        Dim Great_Increase As Double
        Great_Increase = 0
        
        'Set an initial variable for the greatest percentage decrease
        Dim Great_Decrease As Double
        Great_Decrease = 0

        'Set an initial variable for the ticker's name with the greatest percentage increase
        Dim TickerNameInc As String
        TickerNameInc = 0

        'Set an initial variable for the ticker's name with the greatest percentage decrease
        Dim TickerNameDec As String
        TickerNameDec = 0

        ' Create a variable to hold the greatest total volume
        Dim Great_Volume As LongLong
        Great_Volume = 0

        ' Create a variable to hold the ticker's name with the greatest total volume
        Dim TickerVol As String

        ' Loop through the summary table
        For i = 2 To Summary_Table_Row

            ' Check if the percent change is greater than the greatest percent change previoulsy set
            If ws.Cells(i, 11).Value > Great_Increase Then
                
                ' Set the new greatest change
                Great_Increase = ws.Cells(i, 11).Value
                
                ' Set the ticker's name
                TickerNameInc = ws.Cells(i, 9).Value

            ' Check if the percent change is lower than the lowest percent change previoulsy set
            ElseIf ws.Cells(i, 11).Value < Great_Decrease Then
                
                ' Set the greatest decrease
                Great_Decrease = ws.Cells(i, 11).Value
                
                ' Set the ticker's name
                TickerNameDec = ws.Cells(i, 9).Value

            End If

            ' Check if the total volume is greater than the greatest total volume previoulsy set
            If ws.Cells(i, 12).Value > Great_Volume Then
                
                ' Set the total volume
                Great_Volume = ws.Cells(i, 12).Value
                
                ' Set the ticker's name
                TickerVol = ws.Cells(i, 9).Value

            End If

        Next i

        ' Print the second table headers and row names
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Print the value in the second table
        ws.Range("P2").Value = TickerNameInc
        ws.Range("Q2").Value = Great_Increase
        ws.Range("P3").Value = TickerNameDec
        ws.Range("Q3").Value = Great_Decrease
        ws.Range("P4").Value = TickerVol
        ws.Range("Q4").Value = Great_Volume
        
        ' Format the second table
        ws.Range("Q2:Q3").Style = " Percent"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("O1:Q4").BorderAround Weight:=xlMedium
        ws.Range("O1:Q1").BorderAround Weight:=xlMedium
        ws.Range("O1:Q1").Font.Bold = True
        
        ' Adjust the columns' size
        ws.Columns("A:Q").AutoFit

    Next ws

End Sub





