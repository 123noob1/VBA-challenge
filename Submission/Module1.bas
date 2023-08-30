Attribute VB_Name = "Module1"
Sub PartOne()
    ' Dim variables
    ' ---------------------------------------
    Dim i As Integer
    Dim j As Long               ' Long due to rows exceed Integer limit
    Dim counter As Integer
    Dim totalVolume As Variant  ' Variant due to total can exceed what Long can hold
    Dim openPrice, closePrice As Double
    Dim yearChange, percentChange As Double
    Dim rowCount, colCount As Long
    Dim Ticker As String
    
    ' Start For loop to go through each worksheet
    ' -------------------------------------------
    For i = 1 To Worksheets.Count
    
        With Worksheets(i)
        
            ' Get current total row and column counts and assign to correct variables
            ' -----------------------------------------------------------------------
            rowCount = .Cells(.Rows.Count, 1).End(xlUp).Row
            colCount = .Cells(1, .Columns.Count).End(xlToLeft).Column
            
            ' Add the column headers required for the assignment
            ' --------------------------------------------------
            .Cells(1, colCount + 2).Value = "Ticker"
            .Cells(1, colCount + 2).Font.Bold = True
            .Cells(1, colCount + 3).Value = "Yearly Change"
            .Cells(1, colCount + 3).Font.Bold = True
            .Cells(1, colCount + 4).Value = "Percent Change"
            .Cells(1, colCount + 4).Font.Bold = True
            .Cells(1, colCount + 5).Value = "Total Stock Volume"
            .Cells(1, colCount + 5).Font.Bold = True
            .Cells(2, colCount + 8).Value = "Greatest % Increase"
            .Cells(3, colCount + 8).Value = "Greatest % Decrease"
            .Cells(4, colCount + 8).Value = "Greatest Total Volume"
            .Cells(1, colCount + 9).Value = "Ticker"
            .Cells(1, colCount + 9).Font.Bold = True
            .Cells(1, colCount + 10).Value = "Value"
            .Cells(1, colCount + 10).Font.Bold = True
            .Columns(colCount + 8).EntireColumn.ColumnWidth = 20
            .Columns(colCount + 8).EntireColumn.Font.Bold = True
            
            ' Format cell style for [Percent Change] to show percentage
            ' and [Total Stock Volume] to have comma between thousands
            ' as well as other numeric fields that we populate
            ' ---------------------------------------------------------
            .Columns(colCount + 3).EntireColumn.NumberFormat = "0.00"
            .Columns(colCount + 4).EntireColumn.NumberFormat = "0.00%"
            .Columns(colCount + 5).EntireColumn.NumberFormat = "#,##0"
            
            ' Set counter for [Ticker], [Yearly Change],
            ' [Percent Change], and [Total Stock Volume] insert position
            ' -----------------------------------------------------------
            counter = 2
            
            ' Aggregate the the [Tickers] from <ticker> column to the new Ticker column
            ' -------------------------------------------------------------------------
            For j = 2 To rowCount
                
                ' Grab the first opening price @ beginning of year for Ticker
                ' and assign the Ticker value
                ' -----------------------------------------------------------
                If openPrice = 0 Then
                    openPrice = .Cells(j, 3).Value
                    Ticker = .Cells(j, 1).Value
                End If
                
                ' Add the <vol> to get [Total Stock Volume] while on the same Ticker
                ' ------------------------------------------------------------------
                totalVolume = totalVolume + .Cells(j, colCount).Value
                
                ' If the next cell row is another Ticker then start calculations
                ' for current Ticker and assign it to [Ticker], [Yearly Change],
                ' [Percent Change], and [Total Stock Volume]
                ' --------------------------------------------------------------
                If Ticker <> .Cells(j + 1, 1).Value Then
                                        
                    ' Assign Ticker and closing price
                    ' -------------------------------
                    .Cells(counter, colCount + 2).Value = Ticker
                    closePrice = .Cells(j, 6).Value
                    
                    ' Perform calculations then insert into appropriate cells
                    ' -------------------------------------------------------
                    yearChange = closePrice - openPrice
                    percentChange = Round(yearChange / openPrice, 4)
                    
                    .Cells(counter, colCount + 3).Value = yearChange
                    .Cells(counter, colCount + 4).Value = percentChange
                    .Cells(counter, colCount + 5).Value = totalVolume
                    
                    ' Modify cell fill color for [Yearly Change] based on outcome
                    ' -----------------------------------------------------------
                    If yearChange > 0 Then
                        .Cells(counter, colCount + 3).Interior.Color = vbGreen
                    ElseIf yearChange = 0 Then
                        .Cells(counter, colCount + 3).Interior.Color = vbYellow
                    Else
                        .Cells(counter, colCount + 3).Interior.Color = vbRed
                    End If
                    
                    ' Reset variables to be reuse for next Ticker
                    ' -------------------------------------------
                    Ticker = Empty
                    openPrice = 0
                    closePrice = 0
                    yearChange = 0
                    percentChange = 0
                    totalVolume = 0
                    
                    ' Add to move to next Ticker
                    ' --------------------------
                    counter = counter + 1
                    
                End If
            Next j
            
            ' Reset row and column counts before moving to next worksheet
            ' -----------------------------------------------------------
            rowCount = 0
            colCount = 0
            
        End With
        
    Next i
End Sub

Sub PartTwo()
    ' Dim variables
    ' -------------
    Dim i, j As Integer
    Dim colCount, rowCount As Long
    Dim Ticker(2, 1) As Variant
    
    ' Start worksheet loop
    ' --------------------
    For i = 1 To Worksheets.Count

        With Worksheets(i)
                
            ' Get total row and col from the current worksheet so we
            ' don't have to hard code the number. The colCount here
            ' will include the new columns added to the worksheet and
            ' we'll be using the last 2 columns to insert the Ticker
            ' name and the value (in that order).
            ' ---------------------------------------------------------
            For j = 1 To 50
                If .Cells(1, j).Value = "Ticker" Then
                    rowCount = .Cells(.Rows.Count, j).End(xlUp).Row
                    Exit For
                End If
            Next j
            
            colCount = .Cells(1, .Columns.Count).End(xlToLeft).Column
            
            ' Loop through the worksheet's content
            ' ------------------------------------
            For j = 2 To rowCount

                ' Insert the values into an array for comparison in finding
                ' top ticker for the 3 categories (GT % Increase/Decrease)
                ' and GT Total Volume. First row will get inserted as baseline
                ' then will be compare and replaced by the next Ticker if the
                ' value for that one tops the current one.
                '
                ' Ticker(0,0) is for <GT % Incr>, Ticker(1,0) is for <GT % Decr>,
                ' and Ticker(2,0) is for <GT Tot Vol>. the Ticker(#,1) is the
                ' value for that ticker.
                ' ---------------------------------------------------------------
                If j = 2 Then
                    Ticker(0, 0) = .Cells(j, colCount - 8).Value ' Ticker name
                    Ticker(0, 1) = .Cells(j, colCount - 6).Value ' percent Change column for highest increase
                    Ticker(1, 0) = .Cells(j, colCount - 8).Value ' Ticker name
                    Ticker(1, 1) = .Cells(j, colCount - 6).Value ' Percent Change column for lowest decrease
                    Ticker(2, 0) = .Cells(j, colCount - 8).Value ' Ticker name
                    Ticker(2, 1) = .Cells(j, colCount - 5).Value ' Highest Total Stock Volume
                End If
                
                If Ticker(0, 1) < .Cells(j + 1, colCount - 6).Value Then
                    Ticker(0, 0) = .Cells(j + 1, colCount - 8).Value ' Ticker name
                    Ticker(0, 1) = .Cells(j + 1, colCount - 6).Value ' percent Change column for highest increase
                End If
                
                If Ticker(1, 1) > .Cells(j + 1, colCount - 6).Value Then
                    Ticker(1, 0) = .Cells(j + 1, colCount - 8).Value ' Ticker name
                    Ticker(1, 1) = .Cells(j + 1, colCount - 6).Value ' Percent Change column for lowest decrease
                End If

                If Ticker(2, 1) < .Cells(j + 1, colCount - 5).Value Then
                    Ticker(2, 0) = .Cells(j + 1, colCount - 8).Value ' Ticker name
                    Ticker(2, 1) = .Cells(j + 1, colCount - 5).Value ' Highest Total Stock Volume
                End If
                
            Next j
            
            ' Populate the top tickers into their respective rows
            ' Also modify the format
            ' ---------------------------------------------------
            .Cells(2, colCount - 1).Value = Ticker(0, 0)
            .Cells(2, colCount).Value = Ticker(0, 1)
            .Cells(2, colCount).NumberFormat = "0.00%"
            .Cells(3, colCount - 1).Value = Ticker(1, 0)
            .Cells(3, colCount).Value = Ticker(1, 1)
            .Cells(3, colCount).NumberFormat = "0.00%"
            .Cells(4, colCount - 1).Value = Ticker(2, 0)
            .Cells(4, colCount).Value = Ticker(2, 1)
            .Cells(4, colCount).NumberFormat = "#,##0"

        End With

    Next i
    
End Sub


