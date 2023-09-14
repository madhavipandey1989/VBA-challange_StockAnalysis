Sub tickerPriceAnalysis()

Dim WS_Count As Integer
Dim I As Integer
' Set WS_Count equal to the number of worksheets in the active workbook.
WS_Count = ActiveWorkbook.Worksheets.Count
' Begin the loop.
For I = 1 To WS_Count
    
    'Find the number of active rows in the worksheet
    Dim LR As Long
    LR = ActiveWorkbook.Worksheets(I).Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    
    
    'Find the number of records for the first ticker, This will help to establish the overal loop
    Dim firstTickerInSheet As String
    firstTickerInSheet = ActiveWorkbook.Worksheets(I).Range("A2").Value
    Dim firstTickersTotalRecords As Integer
    firstTickersTotalRecords = 0
    
    Dim M As Long
    For M = 2 To LR
        If StrComp(ActiveWorkbook.Worksheets(I).Range("A" & M).Value, firstTickerInSheet, vbBinaryCompare) = 0 Then
            firstTickersTotalRecords = firstTickersTotalRecords + 1
        Else
            Exit For
        End If
    Next M
    
    'Setting up each Tickers total records to first Tickers record
    Dim eachTickersTotalRecords As Integer
    eachTickersTotalRecords = firstTickersTotalRecords
    
    'Valiable to hold unique tickets in sheet
    Dim tickerCounter As Integer
    tickerCounter = 1
    
    'Defining the variables which are to be populated in sheet later.
    Dim tickerOpenPrice As Double
    Dim tickerClosePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalTickerVolume As Double
    
    'Defining analysis variables
    Dim greatestPercentIncrease As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolume As Double
    Dim greatestTotalVolumeTicker As String
    
    'Looping through the sheet rows to extract the analysis
    Dim J As Long
    For J = 2 To LR
        'If the value of ticker not written in unique ticker column's last row, then only enter this logic
        If Not StrComp(ActiveWorkbook.Worksheets(I).Range("A" & J).Value, ActiveWorkbook.Worksheets(I).Range("I" & tickerCounter).Value, vbBinaryCompare) = 0 Then
            'Increase the unique ticker count by 1
            tickerCounter = tickerCounter + 1
            'Copy the new ticker in unique ticker column
            ActiveWorkbook.Worksheets(I).Range("I" & tickerCounter).Value = ActiveWorkbook.Worksheets(I).Range("A" & J).Value
            'fetch ticker Open Price from column C of first rwo of ticker records
            tickerOpenPrice = ActiveWorkbook.Worksheets(I).Range("C" & J).Value
            'fetch ticker close price from column F of the last row of ticker records
            tickerClosePrice = ActiveWorkbook.Worksheets(I).Range("F" & (J + eachTickersTotalRecords - 1)).Value
            'eachTickersTotalRecords = 1
            'calculate the yearly change by reducing first records open price from last records close price
            yearlyChange = tickerClosePrice - tickerOpenPrice
            'Writing yearly change to column J of unique ticker record
            ActiveWorkbook.Worksheets(I).Range("J" & tickerCounter).Value = yearlyChange
            'Calculate percent change for yearly change to ticker open price of ticker records
            percentChange = (yearlyChange * 100) / tickerOpenPrice
            'Writing percent change to column K of unique ticker records
            ActiveWorkbook.Worksheets(I).Range("K" & tickerCounter).Value = percentChange
            'Now we will calculate total ticker volume by adding all records of volume column on G
            'Skipping the first row for  writing total volume in sheet column L
            If J > 2 Then
                ActiveWorkbook.Worksheets(I).Range("L" & tickerCounter - 1).Value = totalTickerVolume
            End If
            'assigning the first value of the column J means volume to total ticker volume
            totalTickerVolume = ActiveWorkbook.Worksheets(I).Range("G" & J).Value
            
            'Assigning percent increase and decrease and finding tickers for that
            If percentChange > greatestPercentIncrease Then
                greatestPercentIncrease = percentChange
                greatestPercentIncreaseTicker = ActiveWorkbook.Worksheets(I).Range("A" & J).Value
            ElseIf percentChange < greatestPercentDecrease Then
                greatestPercentDecrease = percentChange
                greatestPercentDecreaseTicker = ActiveWorkbook.Worksheets(I).Range("A" & J).Value
            End If
            
        Else
            totalTickerVolume = totalTickerVolume + ActiveWorkbook.Worksheets(I).Range("G" & J).Value
            'assigning greatest ticket volume and finding ticker for that
            If totalTickerVolume > greatestTotalVolume Then
                greatestTotalVolume = totalTickerVolume
                greatestTotalVolumeTicker = ActiveWorkbook.Worksheets(I).Range("A" & J).Value
            End If
            'eachTickersTotalRecords = eachTickersTotalRecords + 1
        End If
        
    Next J
    
    'Writing analysis values to sheet
    ActiveWorkbook.Worksheets(I).Range("P2").Value = greatestPercentIncreaseTicker
    ActiveWorkbook.Worksheets(I).Range("Q2").Value = greatestPercentIncrease
    ActiveWorkbook.Worksheets(I).Range("P3").Value = greatestPercentDecreaseTicker
    ActiveWorkbook.Worksheets(I).Range("Q3").Value = greatestPercentDecrease
    ActiveWorkbook.Worksheets(I).Range("P4").Value = greatestTotalVolumeTicker
    ActiveWorkbook.Worksheets(I).Range("Q4").Value = greatestTotalVolume
    
    
    'Conditional Formatting for Yearly change column
    'Definining the variables:
    Dim rng As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition

    'Fixing/Setting the range on which conditional formatting is to be desired
    Set rng = ActiveWorkbook.Worksheets(I).Range("J2", "J" & tickerCounter)

    'To delete/clear any existing conditional formatting from the range
    rng.FormatConditions.Delete

    'Defining and setting the criteria for each conditional format
    Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")

    'Defining and setting the format to be applied for each condition
    With condition1
    .Interior.Color = vbGreen
    End With

    With condition2
     .Interior.Color = vbRed
    End With


    'Conditional Formatting for Percent Change column

    Dim range2 As Range
    'Create range object
    Set range2 = ActiveWorkbook.Worksheets(I).Range("K2", "K" & tickerCounter)
    'Delete previous conditional formats
    range2.FormatConditions.Delete


    range2.FormatConditions.AddColorScale ColorScaleType:=3
    'Select color for the lowest value in the range
    range2.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With range2.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = vbRed
    End With
    'Select color for the middle values in the range
    range2.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    range2.FormatConditions(1).ColorScaleCriteria(2).Value = 50
   'Select the color for the midpoint of the range
    With range2.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = vbWhite
    End With
     'Select color for the highest value in the range
    range2.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With range2.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = vbGreen
    End With

    
Next I
End Sub





