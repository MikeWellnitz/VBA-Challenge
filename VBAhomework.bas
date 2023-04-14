Attribute VB_Name = "Module1"
Option Explicit

Sub stocks()

    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
    
    
    Const INPUT_TICKER_COL As Integer = 1
    Const INPUT_VOLUME_COL = 7
    Const OPEN_PRICE_COL = 3
    Const CLOSE_PRICE_COL = 6
    Const OUTPUT_TICKER_COL = 9
    Const OUTPUT_CHANGE_COL = 10
    Const FIRST_DATA_ROW = 2
    Const OUTPUT_VOLUME = 12
    Const PERCENT_CHANGE_COL = 11
    
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim totalvolume As LongLong
    Dim inputrow As Long
    Dim outputrow As Long
    Dim totalrows As Long
    Dim yearlychangeamount As Double
    Dim yearlychangefraction As Double
    Dim increase As Double
    Dim decrease As Double
    Dim maxvolume As LongLong
    Dim i As Integer
    
    
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    Dim rng1 As Range
    Dim rng2 As Range
    
        
    outputrow = FIRST_DATA_ROW
    totalrows = Application.CountA(Columns(INPUT_TICKER_COL))
    
    For inputrow = FIRST_DATA_ROW To totalrows
        ticker = Cells(inputrow, INPUT_TICKER_COL).Value
        If Cells(inputrow - 1, INPUT_TICKER_COL).Value <> ticker Then
            'We are on the first row of the new stock
            totalvolume = 0
            openprice = Cells(inputrow, OPEN_PRICE_COL)
        End If
        
        totalvolume = totalvolume + Cells(inputrow, INPUT_VOLUME_COL).Value
        
        If Cells(inputrow + 1, INPUT_TICKER_COL).Value <> ticker Then
            'We are on the last row of the current stock
            'inputs
            closeprice = Cells(inputrow, CLOSE_PRICE_COL).Value
            
            'calculations
            yearlychangeamount = closeprice - openprice
            yearlychangefraction = yearlychangeamount / openprice
            
            
            'outputs
            Cells(outputrow, OUTPUT_TICKER_COL).Value = ticker
            Cells(outputrow, OUTPUT_CHANGE_COL).Value = yearlychangeamount
            Cells(outputrow, PERCENT_CHANGE_COL).Value = FormatPercent(yearlychangefraction)
            Cells(outputrow, OUTPUT_VOLUME).Value = totalvolume
                        
            'setup for next stock
            outputrow = outputrow + 1
            
            'conditional format
            Set rng1 = Range("j2", "j3001")
            rng1.FormatConditions.Delete
            
            Set condition1 = rng1.FormatConditions.Add(xlCellValue, xlGreater, "=0")
            Set condition2 = rng1.FormatConditions.Add(xlCellValue, xlLess, "<0")
           
            With condition1
                .Interior.ColorIndex = 4
                End With

            With condition2
                .Interior.ColorIndex = 3
            End With
            
            Set rng2 = Range("k2", "k3001")
            rng2.FormatConditions.Delete
            
            Set condition1 = rng2.FormatConditions.Add(xlCellValue, xlGreater, "=0")
            Set condition2 = rng2.FormatConditions.Add(xlCellValue, xlLess, "<0")
            
            With condition1
                .Interior.ColorIndex = 4
                End With

            With condition2
                .Interior.ColorIndex = 3
            End With
            
            
        End If
           
           
    Next inputrow
    

    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("o1").Value = "Ticker"
    Range("p1").Value = "Value"
    Range("n2").Value = "Greatest % Increase"
    Range("n3").Value = "Greatest % Decrease"
    Range("n4").Value = "Greatest Total Volume"
    
    
    
    
    
    
    
    
    
    Next ws
    
     
    'NEW STUFF BELOW TRIAL
    Dim greatest_increase As Long
    greatest_increase = 0
    
    For i = 2 To 123
        increase = Cells(i, 11).Value 'increase just a name for what's in k column
        
        If increase > greatest_increase Then
            greatest_increase = increase
        
        
        
        End If
    
    Next i
    
    Cells(2, 16).Value = greatest_increase
   
    'totalrows = Application.CountA(Columns(OUTPUT_CHANGE_COL))
    
    'For increase = OUTPUT_CHANGE_COL To totalrows
       ' ticker = Cells(inputrow, INPUT_TICKER_COL).Value
        'If Cells(inputrow - 1, INPUT_TICKER_COL).Value <> ticker Then
            'We are on the first row of the new stock
           ' totalvolume = 0
           ' openprice = Cells(inputrow, OPEN_PRICE_COL)
     '   End If
    

End Sub

