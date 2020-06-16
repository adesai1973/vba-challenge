Attribute VB_Name = "Module1"
'=================================================================================
'Challenge 1 - User presses the "Execute button". The following code first gets all the worksheet names in the current workbook.
'                      Then for each worksheet, we look at ticker name and get each ticker's opening price for the year, closing price for the
'                      year and its total stock volume. We store this in new columns on same worksheet.
'=================================================================================
Sub Button1_Click()

    Dim i As Long
    Dim j As Long
    Dim k As Integer                            'This variable is a counter to increment through number of worksheets within a workbook
    Dim myworkbook As Workbook
    Dim Worksheet() As String             'This variable stores the worksheet names within a workbook
    Dim total_sheets As Integer           'This variable stores the number of worksheets within a workbook
    Dim lastrow() As Long                    'This variable stores the last row count in each worksheets
    Dim lastrowString As String
    Dim newlastrow() As Integer          'This variable stores the last row count in each results columns in each worksheet
    Dim tickers() As Variant                 'This variable stores the ticker symbol string for results column
    Dim ticker_first_day() As Long
    Dim ticker_count As Long
    Dim year_open_ticker As Double
    Dim year_close_ticker As Double
    Dim total_stock_volume() As Variant
    Dim tick As String                               'This is for Challenge # 2
    Dim greatest_increase As Double        'This is for Challenge # 2
    Dim greatest_decrease As Double       'This is for Challenge # 2
    Dim greatest_total_volume As Double 'This is for Challenge # 2
    
    'Debug.Print (vbCrLf & "=====================================")
    
    Set myworkbook = ActiveWorkbook
    total_sheets = myworkbook.Sheets.Count
    ReDim Worksheet(total_sheets)
    
    'Get Worksheet names and store it in array Worksheet()
    For i = 0 To (total_sheets - 1)
        Worksheet(i) = myworkbook.Sheets(i + 1).Name
    Next i
    
    ' Write in column I the worksheet names from Worksheet() array - Column I (I6, I7,...)
    For i = 0 To (total_sheets - 1)
        Sheets(Worksheet(i)).Range("I5") = "Worksheet Names"
        Sheets(Worksheet(i)).Range("I5").Interior.ColorIndex = 6
        Sheets(Worksheet(i)).Columns("I").AutoFit
        For j = 0 To (total_sheets - 1)
            Sheets(Worksheet(i)).Range("I" & (j + 6)) = Worksheet(j)
        Next j
    Next i
    
    'MsgBox Join(Worksheet, vbCrLf)
    
    ReDim lastrow(total_sheets)
    lastrow(0) = WorksheetFunction.CountA(Sheets(1).Columns("A:A"))
    'MsgBox lastrow(0)
    
    'Blank out column contents for columns J through O and white out colors in column M - This clears any trace effects from last execution of this program
    For k = 0 To (total_sheets - 1)
        lastrow(k) = WorksheetFunction.CountA(Sheets(k + 1).Columns("A:A"))
        'Assume there will not be more than original number of rows (lastrow) in results worksheet columns
        For i = 1 To lastrow(k)
            myworkbook.Sheets(k + 1).Range("J" & i & ":" & "O" & i) = ""
            Sheets(Worksheet(k)).Range("M" & i).Interior.ColorIndex = 2
        Next i
    Next k


    ReDim tickers(lastrow(0))
    ReDim ticker_first_day(lastrow(0))
    ReDim total_stock_volume(total_sheets - 1, lastrow(0) - 1)
    ReDim total_stock_volume(k, 1000000)

    tickers(0) = ""
    total_stock_volume(0, 0) = 0

    'Start creating the column names and redim the array sizes based on number of rows in that worksheet
    For k = 0 To (total_sheets - 1)
        ticker_count = 0
        Sheets(Worksheet(k)).Range("J1") = "Tickers"
        Sheets(Worksheet(k)).Range("K1") = "Opening Ticker for Year"
        Sheets(Worksheet(k)).Range("L1") = "Closing Ticker for Year"
        ReDim tickers(lastrow(k))
        ReDim ticker_first_day(lastrow(k))
        
        'Populate new columns J, K and L for each ticker begining value and end value for that year
        For i = 2 To lastrow(k)

            total_stock_volume(k, ticker_count) = total_stock_volume(k, ticker_count) + Sheets(Worksheet(k)).Range("G" & i)
            
            'Debug.Print ("k=" & k & ", " & "i=" & i & " Ticker_Count=" & ticker_count)
            'Debug.Print ("Gi= " & Sheets(Worksheet(k)).Range("G" & i))
            'If k = 0 And ticker_count = 171 Then
            '    Debug.Print ("Total Stock Volume = " & total_stock_volume(ticker_count) & " for ticker_count = " & ticker_count & " " & tickers(ticker_count))
            'End If
            'Debug.Print "Tickers = " & (tickers(ticker_count))
            'Debug.Print "Ai = " & Sheets(Worksheet(k)).Range("A" & i)
            
            If (Sheets(Worksheet(k)).Range("A" & i) <> tickers(ticker_count)) Then
                ticker_count = ticker_count + 1
                total_stock_volume(k, ticker_count) = 0
                tickers(ticker_count) = Sheets(Worksheet(k)).Range("A" & i)
                Sheets(Worksheet(k)).Range("J" & ticker_count + 1) = tickers(ticker_count)
                'Debug.Print "Tickers (TC) = " & tickers(ticker_count)
                'Debug.Print "Ji (TC+1)=" & Sheets(Worksheet(k)).Range("J" & ticker_count + 1)
                
                ticker_first_day(ticker_count) = Sheets(Worksheet(k)).Range("B" & i)
                Sheets(Worksheet(k)).Range("K" & ticker_count + 1) = Sheets(Worksheet(k)).Range("C" & i)
                
                If (i <> 2) Then
                    Sheets(Worksheet(k)).Range("L" & ticker_count) = Sheets(Worksheet(k)).Range("F" & i - 1)
                End If
                
                
            End If
            
            If (i = lastrow(k)) Then
                'MsgBox (Sheets(Worksheet(k)).Range("F" & i))
                Sheets(Worksheet(k)).Range("L" & ticker_count + 1) = Sheets(Worksheet(k)).Range("F" & i)
            End If
            
        Next i
    Next k
    
    'Results columns have less number of rowns, so we use newlastrow() to keep number fo rows for results for each worksheet
    ReDim newlastrow(WorksheetFunction.CountA(Sheets(1).Columns("J:J")))
    newlastrow(0) = WorksheetFunction.CountA(Sheets(1).Columns("J:J"))
    
'Calculate columns M, N and O in results columns. Set N column cells to green and red for positive and negative change
   For k = 0 To (total_sheets - 1)
       newlastrow(k) = WorksheetFunction.CountA(Sheets(k + 1).Columns("J:J"))
'       MsgBox newlastrow(k)
'       For i = 0 To (UBound(newlastrow) - 1)
'          lastrowString = lastrowString & newlastrow(i) & vbCr
'       Next i
'       MsgBox lastrowString
'       Debug.Print lastrowString
        Sheets(Worksheet(k)).Range("M1") = "Yearly Change"
        Sheets(Worksheet(k)).Range("N1") = "Percent Change"
        Sheets(Worksheet(k)).Range("O1") = "Total Stock Volume"
        
        For i = 2 To newlastrow(k)
            Sheets(Worksheet(k)).Range("M" & i) = Sheets(Worksheet(k)).Range("L" & i) - Sheets(Worksheet(k)).Range("K" & i)
            
            If (Sheets(Worksheet(k)).Range("K" & i)) <> 0 Then
                Sheets(Worksheet(k)).Range("N" & i) = (Sheets(Worksheet(k)).Range("M" & i)) / (Sheets(Worksheet(k)).Range("K" & i))
                Sheets(Worksheet(k)).Range("N" & i) = Format(Sheets(Worksheet(k)).Range("N" & i), "Percent")
            Else
                Sheets(Worksheet(k)).Range("N" & i) = 0
            End If

           'Debug.Print "Tot Stk Vol [" & i & "]=" & total_stock_volume(k, i)

            Sheets(Worksheet(k)).Range("O" & i) = total_stock_volume(k, i - 1)

            If Sheets(Worksheet(k)).Range("M" & i) >= 0 Then
                Sheets(Worksheet(k)).Range("M" & i).Interior.ColorIndex = 4
                'Sheets(Worksheet(k)).Range("M" & i).Font.Color = 2    ----Use this if black font is hard to see
            Else
                Sheets(Worksheet(k)).Range("M" & i).Interior.ColorIndex = 3
                'Sheets(Worksheet(k)).Range("M" & i).Font.Color = 2    ----Use this if black font is hard to see
            End If
        Next i

    Next k
    
    '================================================================================
    'Challenge 2 - Advanced  --- Find the tickers with greatest % increaseand greatest % decrease and ticker with greatest total volume
    '================================================================================

    For k = 0 To (total_sheets - 1)

        newlastrow(k) = WorksheetFunction.CountA(Sheets(k + 1).Columns("J:J"))

        'Setup new column titles and row title for the summary table
        Sheets(Worksheet(k)).Range("Q2") = "Greatest % Increase"
        Sheets(Worksheet(k)).Range("Q3") = "Greatest % Decrease"
        Sheets(Worksheet(k)).Range("Q4") = "Greatest Total Volume"
        Sheets(Worksheet(k)).Range("R1") = "TICKER"
        Sheets(Worksheet(k)).Range("S1") = "VALUE"

        tick = Sheets(Worksheet(k)).Range("J" & 2)
        greatest_increase = Sheets(Worksheet(k)).Range("N" & 2)
        greatest_decrease = Sheets(Worksheet(k)).Range("N" & 2)
        greatest_total_volume = Sheets(Worksheet(k)).Range("O" & 2)

       'Find greatest % increase and its ticker and save in new summary table (Column R & S)
        For i = 3 To newlastrow(k)

                If Sheets(Worksheet(k)).Range("N" & i) > greatest_increase Then
                    greatest_increase = Sheets(Worksheet(k)).Range("N" & i).Value
                    tick = Sheets(Worksheet(k)).Range("J" & i).Value
                End If

        Next i
        
        Sheets(Worksheet(k)).Range("R2") = tick
        Sheets(Worksheet(k)).Range("S2") = greatest_increase
        Sheets(Worksheet(k)).Range("S2") = Format(Sheets(Worksheet(k)).Range("S2"), "Percent")
        
        'Find greatest % decrease and its ticker and save in new summary table (Column R & S)
        For i = 3 To newlastrow(k)

                If Sheets(Worksheet(k)).Range("N" & i) < greatest_decrease Then
                    greatest_decrease = Sheets(Worksheet(k)).Range("N" & i).Value
                    tick = Sheets(Worksheet(k)).Range("J" & i).Value
                End If

        Next i
        
        Sheets(Worksheet(k)).Range("R3") = tick
        Sheets(Worksheet(k)).Range("S3") = greatest_decrease
        Sheets(Worksheet(k)).Range("S3") = Format(Sheets(Worksheet(k)).Range("S3"), "Percent")
        
        'Find greatest total volume and its ticker and save in new summary table (Column R & S)
        For i = 3 To newlastrow(k)

                If Sheets(Worksheet(k)).Range("O" & i) > greatest_total_volume Then
                    greatest_total_volume = Sheets(Worksheet(k)).Range("O" & i).Value
                    tick = Sheets(Worksheet(k)).Range("J" & i).Value
                End If

        Next i
        Sheets(Worksheet(k)).Range("R4") = tick
        Sheets(Worksheet(k)).Range("S4") = greatest_total_volume
        
        Sheets(Worksheet(k)).Columns("I:S").AutoFit
        
    Next k
    
    'Message Box to user that program finished execution - important for a program that takes a while to run
    MsgBox "TaDa!!!...Program Execution Finished."
    
End Sub


