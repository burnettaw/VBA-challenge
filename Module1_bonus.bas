Attribute VB_Name = "Module1"
Sub ticker_checker()
     
   
    'Variables for calculations
    Dim Ticker As String
    Dim Year_Opening_Value As Double
    Dim Year_Closing_Value As Double
    Dim Ticker_Summary_Row As Integer
    Dim T_total As Double
    Dim greatest_percent_decrease, greatest_percent_increase, greatest_total_volume As Double
    Dim greatest_ticker As String
    
    
    
   
    'display header in range
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Ticker_Summary_Row = 2
    T_total = 0
    Year_Opening_Value = Range("C2").Value
    
    greatest_percent_decrease = 0
    greatest_percent_increase = 0
    'greatest_volume = 0
    greatest_total_volume = 0
    
    For i = 2 To lastRow
        
        ' Check if we are still within the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1).Value
            Year_Closing_Value = Cells(i, 6).Value
            T_total = T_total + Cells(i, 7).Value
           
                      
            'display in cell'
            Range("I" & Ticker_Summary_Row).Value = Ticker
            Range("J" & Ticker_Summary_Row).Value = Year_Closing_Value - Year_Opening_Value
            
            If Year_Opening_Value > 0 Then
                Range("K" & Ticker_Summary_Row).Value = ((Year_Closing_Value - Year_Opening_Value) / Year_Opening_Value) * 100
            End If
            
            Range("L" & Ticker_Summary_Row).Value = T_total
            
            If Range("K" & Ticker_Summary_Row).Value < 0 Then
            ' Set the Cell Colors to Red
                Range("K" & Ticker_Summary_Row).Interior.ColorIndex = 3
            Else
            ' Set the Font Color to Green
                Range("K" & Ticker_Summary_Row).Interior.ColorIndex = 4
            End If
            
            If Range("K" & Ticker_Summary_Row).Value > greatest_percent_increase Then
                greatest_percent_increase = Range("K" & Ticker_Summary_Row).Value
                greatest_increase_ticker = Ticker
                
            ElseIf Range("K" & Ticker_Summary_Row).Value < greatest_percent_increase Then
                greatest_percent_decrease = Range("K" & Ticker_Summary_Row).Value
                greatest_decrease_ticker = Ticker

            End If
            
            If Range("L" & Ticker_Summary_Row).Value > greatest_total_volume Then
                greatest_total_volume = Range("L" & Ticker_Summary_Row).Value
                greatest_ticker = Ticker
            End If
            
            
            Range("Q1").Value = "Value"
            Range("Q2").Value = greatest_percent_increase
            Range("Q3").Value = greatest_percent_decrease
            Range("Q4").Value = greatest_total_volume
            Range("P2").Value = greatest_increase_ticker
            Range("P3").Value = greatest_decrease_ticker
            Range("P4").Value = greatest_ticker
            

            Ticker_Summary_Row = Ticker_Summary_Row + 1
            T_total = 0
            Year_Opening_Value = Cells(i + 1, 3).Value
            
          Else
            
             T_total = T_total + Cells(i, 7).Value
            
        End If
    Next i
    
     last_cell = Cells(Rows.Count, 11).End(xlUp).Row

    
    'display bonus summary in range
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total"
    Range("P1").Value = "Ticker"
    
     

End Sub




