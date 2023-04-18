Attribute VB_Name = "Module1"
Sub VBA_Challenge()

    'Declare and set worksheet
    
    Dim ws As Worksheet

    'Loop through all worksheets
    
    For Each ws In Worksheets


        'Set new variables for prices and percent changes
        
        Dim ticker_open As Double
        ticker_open = 0

        Dim ticker_close As Double
        ticker_close = 0

        Dim yearly_change As Double
        yearly_change = 0

        Dim percent_change As Double
        percent_change = 0

        Dim total_stock_volume As Double
        total_stock_volume = 0


        'Create the column headings
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"


        'Define Ticker variable
        
        Dim Ticker As String

        Dim Ticker_Row As Long
        
        
        'Start from row 2 for ticker output
        
        Ticker_Row = 2

        Dim Ticker_volume As Double
        Ticker_volume = 0


        'Define Lastrow of worksheet
        
        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
              
        
        'Start from row 2 for total stock volume output
        
        Dim total_stock_volume_row As Long
        total_stock_volume_row = 2
        
        
        'Loop of current worksheet to Lastrow
        
        For i = 2 To Lastrow

            
            'Ticker symbol Print
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Row).Value = Ticker
                    Ticker_Row = Ticker_Row + 1

                
                'Calculate yearly change in Price
                
                    ticker_close = ws.Cells(i, 6).Value
                    ticker_open = ws.Cells(i, 3).Value
                    yearly_change = ticker_close - ticker_open
                    ws.Range("J" & Ticker_Row - 1).Value = yearly_change
                
               'Calculate percent change
                
                If ticker_open <> 0 Then
                    percent_change = (yearly_change / ticker_open) * 100
                    ws.Range("K" & Ticker_Row - 1).Value = percent_change
                End If
    
    
                'Calculate total stock volume
                
                Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Range("L" & total_stock_volume_row).Value = Ticker_volume
                    Ticker_volume = 0
                    total_stock_volume_row = total_stock_volume_row + 1
                End If
                              
                'Reset variables for next ticker
                
                yearly_change = 0
                percent_change = 0

                ElseIf ticker_open <> 0 Then
                    percent_change = (yearly_change / ticker_open) * 100
                End If

            Next i
        
           'Conditional Colors for Price change for ws
            For c = 2 To 3002
                If ws.Cells(c, 10).Value > 0 Then
                    ws.Cells(c, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(c, 10).Value < 0 Then
                    ws.Cells(c, 10).Interior.ColorIndex = 3
                End If
        
            Next c
            
            Percent_Max = 0
            Percent_Min = 0
            Max_Volume = 0
            Max_Tag = ""
            Min_Tag = ""
            Max_Volume_Tag = ""
            
            For j = 2 To 3002
            
            'Greatest % Increase
                If ws.Cells(j, 11) > Percent_Max Then
                    Percent_Max = ws.Cells(j, 11)
                    Max_Tag = ws.Cells(j, 9)
                End If
            
            'Greatest % Decrease
                If ws.Cells(j, 11) < Percent_Min Then
                    Percent_Min = ws.Cells(j, 11)
                    Min_Tag = ws.Cells(j, 9)
                End If
            
            'Greatest % Total Volume
                If ws.Cells(j, 12) > Max_Volume Then
                    Max_Volume = ws.Cells(j, 12)
                    Max_Volume_Tag = ws.Cells(j, 9)
                End If
            Next j
           
           'Print % increase, decrease, & volume
            ws.Range("P2").Value = Max_Tag
            ws.Range("Q2").Value = Percent_Max
            ws.Range("Q3").Value = Percent_Min
            ws.Range("P3").Value = Min_Tag
            ws.Range("P4").Value = Max_Volume_Tag
            ws.Range("Q4").Value = Max_Volume
            
    Next ws

End Sub

