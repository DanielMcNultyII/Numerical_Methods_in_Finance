'-----------------------------------------------------------------------------------------
' Daniel McNulty II
'
' Code to calculate the European Call option price and greeks for given option parameters
' within Excel, as well as to both calculate and plot the values of two greeks chosen by
' users for a range of spot prices which is also chosen by users.
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' value() function to calculate the price of a European Call option corresponding to user
' input option parameters using the Black-Scholes formula.
'-----------------------------------------------------------------------------------------
Function value(S0, K, rd, rf, T, d1, d2)

    ' Use the cumulative distribution worksheet function built-in to Excel with d1 and d2
    ' as input.
    Nd1 = WorksheetFunction.Norm_S_Dist(d1, True)
    Nd2 = WorksheetFunction.Norm_S_Dist(d2, True)
    
    ' Use input parameters, the normal cumulative distribution values found above, and the
    ' exponential function built-in to Excel in the Black-Scholes formula to calculate the
    ' value of the European Call option corresponding to the input option parameters.
    value = (S0 * Math.Exp(-rf * T) * Nd1) - (K * Math.Exp(-rd * T) * Nd2)
    
End Function

'-----------------------------------------------------------------------------------------
' delta() function to calculate the delta of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function delta(rf, T, d1)

    ' Use the cumulative distribution worksheet function built-in to Excel with d1 as
    ' input.
    Nd1 = WorksheetFunction.Norm_S_Dist(d1, True)
    
    ' Use input parameters, the exponential function built-in to Excel, and the cumulative
    ' distribution value found above to calculate the delta of the European Call option
    ' corresponding to the input option parameters.
    delta = Math.Exp(-rf * T) * Nd1

End Function

'-----------------------------------------------------------------------------------------
' gamma() function to calculate the gamma of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function gamma(S0, Vol, rf, T, d1)

    ' Use the probability mass worksheet function built-in to Excel with d1 as input.
    Pd1 = WorksheetFunction.Norm_S_Dist(d1, False)
    
    ' Use input parameters, the exponential and square root functions built-in to Excel,
    ' and the probability mass value found above to calculate the gamma of the European
    ' Call option corresponding to the input option parameters.
    gamma = (Math.Exp(-rf * T) * Pd1) / (S0 * Vol * Math.Sqr(T))

End Function

'-----------------------------------------------------------------------------------------
' vega() function to calculate the vega of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function vega(S0, Vol, rf, T, d1)

    ' Use the probability mass worksheet function built-in to Excel with d1 as input.
    Pd1 = WorksheetFunction.Norm_S_Dist(d1, False)
    
    ' Use input parameters, the exponential and square root functions built-in to Excel,
    ' and the probability mass value found above to calculate the vega of the European
    ' Call option corresponding to the input option parameters.
    vega = (S0 * Math.Exp(-rf * T) * Pd1 * Math.Sqr(T)) / 100

End Function

'-----------------------------------------------------------------------------------------
' theta() function to calculate the theta of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function theta(S0, K, Vol, rd, rf, T, d1, d2)

    ' Use the cumulative distribution worksheet function built-in to Excel with d1 and d2
    ' as input.
    Nd1 = WorksheetFunction.Norm_S_Dist(d1, True)
    Nd2 = WorksheetFunction.Norm_S_Dist(d2, True)
    
    ' Use the probability mass worksheet function built-in to Excel with d1 as input.
    Pd1 = WorksheetFunction.Norm_S_Dist(d1, False)
    
    ' Use input parameters, the exponential and square root functions built-in to Excel,
    ' and the cumulative distribution values found above, and the probability mass value
    ' found above to calculate the theta of the European Call option corresponding to the
    ' input option parameters.
    theta = -(((S0 * Vol * Math.Exp(-rf * T) * Pd1) / (2 * Math.Sqr(T))) _
            - (rd * K * Math.Exp(-rd * T) * Nd2) _
            + ((S0 * rf * Math.Exp(-rf * T) * Nd1))) / 365
    
End Function

'-----------------------------------------------------------------------------------------
' phi() function to calculate the phi of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function phi(S0, rf, T, d1)

    ' Use the cumulative distribution worksheet function built-in to Excel with d1 as
    ' input.
    Nd1 = WorksheetFunction.Norm_S_Dist(d1, True)
    
    ' Use input parameters, the exponential function built-in to Excel, and the cumulative
    ' distribution value found above to calculate the phi of the European Call option
    ' corresponding to the input option parameters.
    phi = -(S0 * T * Math.Exp(-rf * T) * Nd1) / 100

End Function

'-----------------------------------------------------------------------------------------
' rho() function to calculate the rho of a European Call option corresponding to user
' input option parameters.
'-----------------------------------------------------------------------------------------
Function rho(K, rd, T, d2)

    ' Use the cumulative distribution worksheet function built-in to Excel with d2 as
    ' input.
    Nd2 = WorksheetFunction.Norm_S_Dist(d2, True)
    
    ' Use input parameters, the exponential function built-in to Excel, and the cumulative
    ' distribution value found above to calculate the rho of the European Call option
    ' corresponding to the input option parameters.
    rho = (K * T * Math.Exp(-rd * T) * Nd2) / 100

End Function

'-----------------------------------------------------------------------------------------
' calculate() sub procedure that:
'   - Take in the European Call parameters input by the user in the Excel worksheet.
'   - Calculate the price and greeks of the European Call option corresponding to the user
'     input using the functions defined above.
'   - Output the calculated European Call price and greeks to the Excel worksheet.
'-----------------------------------------------------------------------------------------
Sub calculate()
    
    ' Take in the European Call parameters input by the user in the Excel worksheet cells
    ' B2:B4 and B9.
    S0 = Range("B2").value
    K = Range("B6").value
    Vol = Range("B5").value
    rd = Range("B3").value
    rf = Range("B4").value
    T = Range("B9").value / 365
    
    ' Calculate d1 and d2 using the user input option parameters as well as the log and
    ' square root built-in Excel functions.
    d1 = (Math.Log(S0 / K) + ((rd - rf + ((Vol * Vol) / 2)) * T)) _
         / (Vol * Math.Sqr(T))
    d2 = d1 - (Vol * Math.Sqr(T))
    
    ' Call the previously defined functions for European Call pricing and calculating
    ' European Call greeks using the user input option parameters and output the results
    ' to the Excel worksheet in cells B12:B18.
    Range("B12").value = value(S0, K, rd, rf, T, d1, d2)
    Range("B13").value = delta(rf, T, d1)
    Range("B14").value = gamma(S0, Vol, rf, T, d1)
    Range("B15").value = vega(S0, Vol, rf, T, d1)
    Range("B16").value = theta(S0, K, Vol, rd, rf, T, d1, d2)
    Range("B17").value = phi(S0, rf, T, d1)
    Range("B18").value = rho(K, rd, T, d2)
    
End Sub

'-----------------------------------------------------------------------------------------
' plot() function to create plots using input
'   - y_title, which is the title of the data in the y-axis of the plot.
'   - x_data, which is used as the x-axis data to be plotted.
'   - y_data, which is used as the y-axis data to be plotted.
'   - top_disp, which is used to input how far away from the top of the worksheet the plot
'     should appear.
'-----------------------------------------------------------------------------------------
Function plot(y_title, x_data, y_data, top_disp)

    ' Add a chart object to the worksheet, specifying the location of the chart.
    Set Chrt = ActiveSheet.ChartObjects.Add(Left:=ActiveSheet.Columns(ActiveWindow.ScrollColumn).Left + 510, _
                                            Top:=ActiveSheet.Rows(ActiveWindow.ScrollRow).Top + top_disp, _
                                            Width:=400, _
                                            Height:=200).chart
    
    ' Count the number of rows and columns there are in y_data.
    Data_Rows = y_data.Rows.Count
    Data_Cols = y_data.Columns.Count
    
    ' Make adjustments to the chart that was created above.
    With Chrt
        
        ' Make the chart a line plot with markers for each datapoint.
        .ChartType = xlLineMarkers
        
        ' Make the title of the plot the input y_title versus the first cell of the x_data range.
        .HasTitle = True
        .ChartTitle.Characters.Text = y_title & " vs " & x_data.Cells(1, 1)
        
        ' Delete the series currently in the chart
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop
        
        ' Loop through the columns in y_data
        For Curr_Col = 1 To Data_Cols
        
            ' Create a new series for the current column in y_data
            Set Ser = .SeriesCollection.NewSeries
            
            ' Make adjustments to the newly created series
            With Ser
                
                ' Name the new series the 1st value in y_data
                .Name = y_data.Cells(1, Curr_Col)
                
                ' Set the values of the new series to the values from the 2nd value in y_data through the last
                ' value in y_data
                .Values = y_data.Range(Cells(2, Curr_Col), Cells(Data_Rows, Curr_Col))
                
                ' Set the x values of the new series to the values from the 2nd value in x_data through the
                ' last value in x_data
                .XValues = x_data.Range(Cells(2, 1), Cells(Data_Rows, 1))
            End With
        Next
        
        ' Label the x-axis of the chart with the 1st value in x_data
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Characters.Text = x_data.Cells(1, 1)
        End With
    End With
    
End Function

'-----------------------------------------------------------------------------------------
' chart() sub procedure that:
'   - Creates a table showing the spot price range specified by user input initial and end
'     spot prices, as well as the user input stepsize, along with the values of the two
'     greeks specified by user input which correspond to the spot price of a given row.
'   - Create three plots from this generated table:
'       1) The first greek chosen against the spot price.
'       2) The second greek chosen against the spot price.
'       3) Both the first and second greeks chosen against the spot price.
'-----------------------------------------------------------------------------------------
Sub chart()
    ' Take in the European Call start and end spot prices input by the user in the Excel
    ' worksheet cells F4:F5, as well as parameters input by the user in the Excel worksheet
    ' cells B3:B6 and B9.
    S_Start = Range("F4").value
    S_End = Range("F5").value
    K = Range("B6").value
    Vol = Range("B5").value
    rd = Range("B3").value
    rf = Range("B4").value
    T = Range("B9").value / 365
    
    ' Calculate the step size needed to get from the starting spot price to the ending spot
    ' price using the user input number of steps from the Excel worksheet cell F6.
    dS = (S_End - S_Start) / (Range("F6").value + 1)
    
    ' Initialize variable Cell_Row to equal 1.
    Cell_Row = 1
    
    ' Output the header titles for the output table in cells E11:G11 in the order of spot
    ' price, first greek chosen, and second greek chosen, taking the chosen greeks from the
    ' Excel worksheet cells F2:F3.
    Range("E11").value = "Spot"
    Range("F11").value = Range("F2").value
    Range("G11").value = Range("F3").value
    
    ' Loop from the first spot price to the last spot price with increments of the step size
    ' calculated above.
    For S = S_Start To S_End Step dS
        
        ' Output the current spot price in the cell specified by row Cell_Row and column E.
        Range("E12").Cells(Cell_Row, 1).value = S
    
        ' Calculate d1 and d2 using the user input option parameters as well as the log and
        ' square root built-in Excel functions.
        d1 = (Math.Log(S / K) + ((rd - rf + ((Vol * Vol) / 2)) * T)) _
             / (Vol * Math.Sqr(T))
        d2 = d1 - (Vol * Math.Sqr(T))
        
        ' If/Else statement which identifies which greek the user chose in cell F2, then
        ' calls the corresponding function defined above which calculates that greek and
        ' stores the result in the cell specified by row Cell_Row and column F.
        If Range("F2").value = "Delta" Then
            Range("F12").Cells(Cell_Row, 1).value = delta(rf, T, d1)
        ElseIf Range("F2").value = "Gamma" Then
            Range("F12").Cells(Cell_Row, 1).value = gamma(S, Vol, rf, T, d1)
        ElseIf Range("F2").value = "Vega" Then
            Range("F12").Cells(Cell_Row, 1).value = vega(S, Vol, rf, T, d1)
        ElseIf Range("F2").value = "Theta" Then
            Range("F12").Cells(Cell_Row, 1).value = theta(S, K, Vol, rd, rf, T, d1, d2)
        ElseIf Range("F2").value = "Phi" Then
            Range("F12").Cells(Cell_Row, 1).value = phi(S, rf, T, d1)
        ElseIf Range("F2").value = "Rho" Then
            Range("F12").Cells(Cell_Row, 1).value = rho(K, rd, T, d2)
        Else
            Range("F12").Cells(Cell_Row, 1).value = "N/A"
        End If
        
        ' If/Else statement which identifies which greek the user chose in cell F3, then
        ' calls the corresponding function defined above which calculates that greek and
        ' stores the result in the cell specified by row Cell_Row and column G.
        If Range("F3").value = "Delta" Then
            Range("G12").Cells(Cell_Row, 1).value = delta(rf, T, d1)
        ElseIf Range("F3").value = "Gamma" Then
            Range("G12").Cells(Cell_Row, 1).value = gamma(S, Vol, rf, T, d1)
        ElseIf Range("F3").value = "Vega" Then
            Range("G12").Cells(Cell_Row, 1).value = vega(S, Vol, rf, T, d1)
        ElseIf Range("F3").value = "Theta" Then
            Range("G12").Cells(Cell_Row, 1).value = theta(S, K, Vol, rd, rf, T, d1, d2)
        ElseIf Range("F3").value = "Phi" Then
            Range("G12").Cells(Cell_Row, 1).value = phi(S, rf, T, d1)
        ElseIf Range("F3").value = "Rho" Then
            Range("G12").Cells(Cell_Row, 1).value = rho(K, rd, T, d2)
        Else
            Range("G12").Cells(Cell_Row, 1).value = "N/A"
        End If
        
        ' Increment Cell_Row by 1.
        Cell_Row = Cell_Row + 1
        
    Next S
    
    ' Create ranges that hold all of the spot prices (Spot_Vals), first chosen greek values
    ' (Greek_1_Vals), and second chosen greek values (Greek_2_Vals).
    Set Spot_Vals = Range("E11", Range("E11").End(xlDown))
    Set Greek_1_Vals = Range("F11", Range("F11").End(xlDown))
    Set Greek_2_Vals = Range("G11", Range("G11").End(xlDown))
    
    ' Create plots for:
    '   - The first greek chosen against the spot price.
    '   - The second greek chosen against the spot price.
    '   - The first and second greeks chosen against the spot price.
    greek_1_chart = plot(Range("F2").value, Spot_Vals, Greek_1_Vals, 0)
    greek_2_chart = plot(Range("F3").value, Spot_Vals, Greek_2_Vals, 210)
    greeks_chart = plot(Range("F2").value & " and " & Range("F3").value, Spot_Vals, Union(Greek_1_Vals, Greek_2_Vals), 420)
    
End Sub

'-----------------------------------------------------------------------------------------
' clear_calculate() sub process that clears the data held in cells B12:B18.
'-----------------------------------------------------------------------------------------
Sub clear_calculate()
    
    ' Set the values in cells B12:B18 to "" (Blank)
    Range("B12:B18").value = ""

End Sub

'-----------------------------------------------------------------------------------------
' clear_chart() sub process that deletes all plots in the active sheet and the values held
' in columns E:G from row 11 to the last filled row.
'-----------------------------------------------------------------------------------------
Sub clear_chart()
    
    ' Loop through all active ChartObjects in the ActiveSheet and delete each.
    For Each Chrt In ActiveSheet.ChartObjects
        Chrt.Delete
    Next
    
    ' Delete the values held in columns E:G from row 11 to the last filled row.
    Range("E11", Range("G11").End(xlDown)).value = ""
    
End Sub