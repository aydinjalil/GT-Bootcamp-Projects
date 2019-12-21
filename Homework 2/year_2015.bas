Attribute VB_Name = "Module2"

Sub stock_market()

' This function is performed on the given list of stocks. The list consist of 7 columns, "ticker", "date", "open", "high", "low", "close", "vol". Each column has 705,713 rows.
' <Date> column specify each day the stock has been on market. First task of this function is to find the the value change and percentage change of given stock from starting day
' till the last day the stock been on the market. Secondly total volume of each stock should be populated to the specified column. Additionally, greatest incerase(best performed stock)
' and greatest decrease(worst performing stock) for the given year with its associated ticker are populated. Lastly, the greatest total volume and associated ticker are populated to a specified cell.
'
'
'

   
    ' This following part of the code is to name the columns
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("2015")
    Dim k As Long

    
    Dim open_ As Double ' Value in opening date of the stock on market
    Dim close_ As Double ' Value in final day of the stock on market


    k = sh.Range("B2", Range("B2").End(xlDown)).Rows.count
    Dim ticker_lst As New Collection ' This declares the ticker_lst as Collection type which is equivalent to list. I found Collection to be easier to operate with.
    Dim start_date_lst As New Collection
    Dim end_date_lst As New Collection
    Dim year_change As Double
    Dim count As Integer
    Dim percent_change As Double
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_vol As Double
    Dim greatest_inc_tick As String
    Dim greatest_dec_tick As String
    Dim greatest_vol_tick As String
    
    Dim cond_format As Range
    
    
    'Creation of unique ticker list
    ticker_lst.Add ("A")
    For i = 2 To k
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            ticker_lst.Add (Cells(i + 1, 1).Value)
        End If
    Next i
    
    'printing ticker list elements under given column
    For i = 1 To ticker_lst.count
        Cells(i + 1, 9).Value = ticker_lst(i)
    Next i
    
    ' Stock volume for each ticker name and list of start and end dates of each stock ticker
    ' Stock volume is found using the SumIf method where you perform sum on the cells that meets the condition (i this case the condition is the ticker name)
    ' start and end dates are found using minIfs and maxIfs function. Because the start days and end dates of each stock on the market may vary it is important
    ' to find the minimum and maximum dates associated with each ticker
    
    'Stock volume for each ticker name and list of start and end dates of each stock ticker
    For i = 1 To ticker_lst.count
    
        Cells(i + 1, 12) = Application.WorksheetFunction.SumIf(Range("A2:A760192"), ticker_lst(i), Range("G2:G760192"))
        
        start_date_lst.Add (Application.WorksheetFunction.MinIfs(Range("B2:B760192"), Range("A2:A760192"), ticker_lst(i)))
        end_date_lst.Add (Application.WorksheetFunction.MaxIfs(Range("B2:B760192"), Range("A2:A760192"), ticker_lst(i)))
    
    Next
    
    count = 1 'Because below I am iterating through the total number of the rows under date column and comparing them to the start dates in the list hwich has only just under 3000 elements,
    'to avoid double for loops which will take exponentially greater time to go through all possible iterations I chouse to manually the indexes of the items in the list. For lists the indexes start from 1.
    
    
    For i = 0 To k - 1
        If Cells(i + 2, 2).Value = start_date_lst(count) Then
            open_ = Cells(i + 2, 3).Value
        End If
        If Cells(i + 2, 2).Value = end_date_lst(count) Then
            close_ = Cells(i + 2, 6).Value
            year_change = close_ - open_
            If open_ <> 0 And year_change <> 0 Then
                percent_change = (year_change) / open_ * 100
            Else
                percent_change = 0
            End If
            
            Cells(count + 1, 10).Value = year_change
            ' I commented out the next piece of the code because I though at the beginning that it is not good method to set column.Value to 0 when the open value of the stock is zerol.
            ' Because the formula of the percentage change is year_change/ open_ *100 and the division operation is not defined when the denominator is 0. I could have wrote formula in log form but
            ' it would not elp either. Log(close_) - Log(open_) wwhen open_ = 0 is still not defined. setting the cell.value to n/A would be a problem when calculating the greatest percent changes.
            ' because n/A is of double type.
'
'            If open_ = 0 Then
'                Cells(count + 1, 11).Value = CVErr(xlErrNA) 'This will assign Not Applicable sign to the percentage change when the start value of the stock is 0
'            Else
            Cells(count + 1, 11).Value = Round(percent_change, 2) & "%"
            ' End If
            count = count + 1
        End If
    Next i
    
    ' Challenge task
    
    ' It is fairly simple logic. The idea is to iterate through the rows uner perentage change column and find the minimum and maximum values.
    ' The same is done with greatest volume
    
    greatest_decrease = 0
    greatest_increase = 0
    greatest_vol = 0
    k = sh.Range("J2", Range("J2").End(xlDown)).Rows.count
    For i = 2 To k - 1
        If Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = Cells(i, 11).Value
            greatest_dec_tick = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value > greatest_increase Then
            greatest_increase = Cells(i, 11).Value
            greatest_inc_tick = Cells(i, 9).Value
        End If
        
        If Cells(i, 12).Value > greatest_vol Then
            greatest_vol = Cells(i, 12).Value
            greatest_vol_tick = Cells(i, 9).Value
        End If
        
        
    Next i
    
     ' Population of the results of above calculations
    
    Cells(2, 16).Value = greatest_inc_tick
    Cells(2, 17).Value = greatest_increase * 100 & "%"
    Cells(3, 16).Value = greatest_dec_tick
    Cells(3, 17).Value = greatest_decrease * 100 & "%"
    Cells(4, 16).Value = greatest_vol_tick
    Cells(4, 17).Value = greatest_vol
    
    
    Set cond_format = Worksheets("2015").Range("J2:J3005")

   ' The next for loop does the conditional formatting on range of cells under J column based on the negative and positive values

    For Each cell_ In cond_format
        If cell_ >= 0 Then
            cell_.Interior.ColorIndex = 4
        Else
            cell_.Interior.ColorIndex = 3
        End If
    Next
    
    
    
        
     
End Sub



