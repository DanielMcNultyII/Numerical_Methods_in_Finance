'-----------------------------------------------------------------------------------------
' Daniel McNulty II
'
' Code to calculate European Call option implied volatility for given option parameters
' within Excel, as well as to calculate Bull Spread implied volatility for given Bull
' spread parameters within Excel.
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' BS_Euro_Call_Price() function to calculate the price of a European Call option
' corresponding to input option parameters using the Black-Scholes formula.
'-----------------------------------------------------------------------------------------
Function BS_Euro_Call_Price(S0, K, rd, rf, Vol, T)

    ' Calculate d1 and d2 using the user input option parameters as well as the log and
    ' square root built-in Excel functions.
    d1 = (Math.Log(S0 / K) + ((rd - rf + ((Vol ^ 2) / 2)) * T)) _
         / (Vol * Math.Sqr(T))
    d2 = d1 - (Vol * Math.Sqr(T))
    
    ' Use the cumulative distribution worksheet function built-in to Excel with d1 and d2
    ' as input.
    Nd1 = WorksheetFunction.Norm_S_Dist(d1, True)
    Nd2 = WorksheetFunction.Norm_S_Dist(d2, True)
    
    ' Use input parameters, the normal cumulative distribution values found above, and the
    ' exponential function built-in to Excel in the Black-Scholes formula to calculate the
    ' value of the European Call option corresponding to the input option parameters.
    BS_Euro_Call_Price = (S0 * Math.Exp(-rf * T) * Nd1) - (K * Math.Exp(-rd * T) * Nd2)
    
End Function

'-----------------------------------------------------------------------------------------
' BS_Euro_Call_Vega() function to calculate the vega of a European Call option
' corresponding to user input option parameters.
'-----------------------------------------------------------------------------------------
Function BS_Euro_Call_Vega(S0, K, rd, rf, Vol, T)

    ' Calculate d1 using the user input option parameters as well as the log and square
    ' root built-in Excel functions.
    d1 = (Math.Log(S0 / K) + ((rd - rf + ((Vol ^ 2) / 2)) * T)) _
         / (Vol * Math.Sqr(T))
    
    ' Use the probability mass worksheet function built-in to Excel with d1 as input.
    Pd1 = WorksheetFunction.Norm_S_Dist(d1, False)
    
    ' Use input parameters, the exponential and square root functions built-in to Excel,
    ' and the probability mass value found above to calculate the vega of the European
    ' Call option corresponding to the input option parameters.
    BS_Euro_Call_Vega = (S0 * Math.Exp(-rf * T) * Pd1 * Math.Sqr(T))
    
End Function

'-----------------------------------------------------------------------------------------
' Euro_Call_Newt_Raph() function to calculate the implied volatility of a European Call
' option with given input parameters and market price using the Newton Raphson method.
' Also takes a guess at the implied vol as input variable Current_Vol, as well as the
' maximum number of iterations to allow the for loop within the function to perform and
' the tolerance to which the user wants the output to be accurate to.
'-----------------------------------------------------------------------------------------
Function Euro_Call_Newt_Raph(S0, K, rd, rf, T, C_mrkt, Current_Vol, MaxIter, Tol)
    
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
        
        ' Calculate the call price with the input option parameters and the Current_Vol.
        C_imp = BS_Euro_Call_Price(S0, K, rd, rf, Current_Vol, T)
        
                
        ' Check if the difference between the European Call price calculated with the
        ' input option parameters and the Current_Vol and the market call price is less
        ' than the input tolerance. If it is, output the Current_Vol to the Excel
        ' spreadsheet and exit the for loop.
        If (Abs(C_imp - C_mrkt) < Tol) Then
            Range("E10").Value = Current_Vol
            Exit For
        End If
        
        ' If the difference between the European Call price calculated with the input
        ' option parameters and the Current_Vol and the market call price is greater than
        ' the input tolerance, calculate the vega of the European Call with the input
        ' option parameters and the Current_Vol.
        vega = BS_Euro_Call_Vega(S0, K, rd, rf, Current_Vol, T)
        
        ' Calculate the next Current_Vol and continue to the next iteration of the loop.
        Current_Vol = Current_Vol - ((C_imp - C_mrkt) / vega)
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print
    ' an error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("E10").Value = "Did not converge to " & Tol & " tolerance. (More than " & MaxIter & " iterations needed)"
    End If
    
End Function

'-----------------------------------------------------------------------------------------
' Euro_Call_Regula_Falsi() function to calculate the implied volatility of a European Call
' option with given input parameters and market price using the Regula Falsi method. Also
' takes 2 guesses for the implied vol as inputs Low_Vol and High_Vol, as well as the
' maximum number of iterations to allow the for loop within the function to perform and
' the tolerance to which the user wants the output to be accurate to.
'-----------------------------------------------------------------------------------------
Function Euro_Call_Regula_Falsi(S0, K, rd, rf, T, C_mrkt, Low_Vol, High_Vol, MaxIter, Tol)
        
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the Low_Vol and High_Vol variables.
        Low_C = BS_Euro_Call_Price(S0, K, rd, rf, Low_Vol, T)
        High_C = BS_Euro_Call_Price(S0, K, rd, rf, High_Vol, T)
        
        ' Determine the implied volatility from the previously calculated European Call
        ' prices and the current Low_Vol and High_Vol variables. Then calculate the price
        ' of the European Call option with the input parameters and the Current_Vol.
        Current_Vol = Low_Vol - ((Low_C - C_mrkt) * ((High_Vol - Low_Vol) / (High_C - Low_C)))
        Current_C = BS_Euro_Call_Price(S0, K, rd, rf, Current_Vol, T)

        ' Check if the difference between the European Call price calculated with the
        ' input option parameters and the Current_Vol and the market call price is less
        ' than the input tolerance. If it is, output the Current_Vol to the Excel
        ' spreadsheet and exit the for loop.
        If (Abs(Current_C - C_mrkt) < Tol) Then
            Range("E11").Value = Current_Vol
            Exit For
        End If
        
        ' If the the difference between the European Call price calculated with the input
        ' option parameters and the Current_Vol and the market call price is less than the
        ' input tolerance, check if the sign of the European Call price calculated with
        ' the Current_Vol matches the sign of the European Call price calculated with the
        ' current Low_Vol. If the signs match, assign the value of Current_Vol to Low_Vol
        ' and continue to the next iteration of this for loop.
        If (Sgn(Current_C) = Sgn(Low_C)) Then
            Low_Vol = Current_Vol
        ' If the above check fails, check if the sign of the European Call price calculated
        ' with the Current_Vol matches the sign of the European Call price calculated with
        ' the current High_Vol. If the signs match, assign the value of Current_Vol to
        ' High_Vol and continue to the next iteration of this for loop.
        ElseIf (Sgn(Current_C) = Sgn(High_C)) Then
            High_Vol = Current_Vol
        ' If both of the above checks fail, output an error message to the Excel worksheet.
        Else
            Range("E11").Value = "Did not converge (Current call price estimate sign matches neither initial call price estimate's sign)"
        End If
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print
    ' an error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("E11").Value = "Did not converge to " & Tol & " tolerance. (More than " & MaxIter & " iterations needed)"
    End If
    
End Function

'-----------------------------------------------------------------------------------------
' Euro_Call_Secant() function to calculate the implied volatility of a European Call
' option with given input parameters and market price using the Secant method. Also takes
' 2 guesses for the implied vol as inputs x1_Vol and x2_Vol, as well as the maximum
' number of iterations to allow the for loop within the function to perform and the
' tolerance to which the user wants the output to be accurate to.
'-----------------------------------------------------------------------------------------
Function Euro_Call_Secant(S0, K, rd, rf, T, C_mrkt, x1_Vol, x2_Vol, MaxIter, Tol)
    
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the x1_Vol and x2_Vol variables.
        x1_C = BS_Euro_Call_Price(S0, K, rd, rf, x1_Vol, T)
        x2_C = BS_Euro_Call_Price(S0, K, rd, rf, x2_Vol, T)
        
        ' Determine the implied volatility from the previously calculated European Call
        ' prices and the current x1_Vol and x2_Vol variables. Then calculate the price of
        ' the European Call option with the input parameters and the Est_Vol.
        Est_Vol = ((x1_Vol * (x2_C - C_mrkt)) - (x2_Vol * (x1_C - C_mrkt))) / (x2_C - x1_C)
        Est_C = BS_Euro_Call_Price(S0, K, rd, rf, Est_Vol, T)
        
        ' Check if the difference between the European Call price calculated with the
        ' input option parameters and the Est_Vol and the market call price is less than
        ' the input tolerance. If it is, output the Current_Vol to the Excel spreadsheet
        ' and exit the for loop.
        If (Abs(Est_C - C_mrkt) < Tol) Then
            Range("E12").Value = Est_Vol
            Exit For
        End If
        
        ' If the the difference between the European Call price calculated with the input
        ' option parameters and the Est_Vol and the market call price is greater than the
        ' input tolerance, check if the sign of the European Call price calculated with
        ' the Est_Vol matches the sign of the European Call price calculated with the
        ' current x1_Vol. If the signs match, assign the value of Est_Vol to x1_Vol and
        ' continue to the next iteration of this for loop.
        If (Sgn(Est_C) = Sgn(x1_C)) Then
            x1_Vol = Est_Vol
        ' If the above check fails, check if the sign of the European Call price calculated
        ' with the Est_Vol matches the sign of the European Call price calculated with the
        ' current x2_Vol. If the signs match, assign the value of Est_Vol to x2_Vol and
        ' continue to the next iteration of this for loop.
        ElseIf (Sgn(Est_C) = Sgn(x2_C)) Then
            x2_Vol = Est_Vol
        ' If both of the above checks fail, output an error message to the Excel worksheet.
        Else
            Range("E12").Value = "Did not converge (Current call price estimate sign matches neither initial call price estimate's sign)"
        End If
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print
    ' an error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("E12").Value = "Did not converge to " & Tol & " tolerance. (" & MaxIter & " iterations performed)"
    End If
    
End Function

'-----------------------------------------------------------------------------------------
' Euro_Call_Calcs() sub procedure that:
'   - Takes in the European Call parameters input by the user in the Excel worksheet, as
'     well as the user input market price of the European Call option, the initial guesses
'     to be used for all methods of determining the implied volatility of the European
'     Call option (Newton Raphson, Regula Falsi, and Secant methods), the maximum number of
'     iterations to use for all 3 methods, and the tolerance to be used for all 3 methods.
'
'   - Calculates the implied volatility of the European call with the user input option
'     parameters and market price using all 3 method functions defined above.
'-----------------------------------------------------------------------------------------
Sub Euro_Call_Calcs()

    ' Take in the European Call parameters input by the user in the Excel worksheet cells
    ' B7:B12.
    S0 = Range("B7").Value
    K = Range("B8").Value
    rd = Range("B9").Value
    rf = Range("B10").Value
    T = Range("B12").Value
    
    ' Take in the European Call price and volatility guesses input by the user in the Excel
    ' worksheet cells E3:E6 and F5:F6.
    C_mrkt = Range("E3").Value
    Current_Vol = Range("E4").Value
    Low_Vol = Range("E5").Value
    High_Vol = Range("F5").Value
    x1_Vol = Range("E6").Value
    x2_Vol = Range("F6").Value
    
    ' Take in the user input maximum number of iterations and tolerance to use for the
    ' following methods of determining the implied volatility from cells B3:B4.
    MaxIter = Range("B3").Value
    Tol = 10 ^ Range("B4").Value
    
    ' Call the previously defined functions for finding the implied volatility of a European
    ' Call using the inputs taken from the Excel spreadsheet above.
    Euro_Call_Newton = Euro_Call_Newt_Raph(S0, K, rd, rf, T, C_mrkt, Current_Vol, MaxIter, Tol)
    Euro_Call_Reg_Fal = Euro_Call_Regula_Falsi(S0, K, rd, rf, T, C_mrkt, Low_Vol, High_Vol, MaxIter, Tol)
    Euro_Call_Sec_Meth = Euro_Call_Secant(S0, K, rd, rf, T, C_mrkt, x1_Vol, x2_Vol, MaxIter, Tol)
    
End Sub

'---------------------------------------------------------------------------------------------
' Bull_Spread_Newt_Raph() function to calculate the implied volatility of a Bull Spread
' with given input parameters and market price using the Newton Raphson method. Also takes
' takes a guess at the implied vol as input variable Current_Vol, as well as the maximum
' number of iterations to allow the for loop within the function to perform and the tolerance
' tolerance to which the user wants the output to be accurate to.
'---------------------------------------------------------------------------------------------
Function Bull_Spread_Newt_Raph(S0, Long_K, Short_K, rd, rf, T, C_mrkt, Current_Vol, MaxIter, Tol)
    
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the Long_K and Short_K variables with the Current_Vol. Then determine the
        ' price of the Bull Spread using the Current_Vol by taking the difference between the
        ' European Call price using the Long_K and the European Call price using the Short_K.
        Long_C_imp = BS_Euro_Call_Price(S0, Long_K, rd, rf, Current_Vol, T)
        Short_C_imp = BS_Euro_Call_Price(S0, Short_K, rd, rf, Current_Vol, T)
        C_imp = Long_C_imp - Short_C_imp
        
        ' Check if the difference between the Bull Spread price determined using the
        ' Current_Vol and the user input market price of the Bull Spread is within the user
        ' input tolerance. If it is, then output the Current_Vol to the Excel spreadsheet and
        ' exit the for loop.
        If (Abs(C_imp - C_mrkt) < Tol) Then
            Range("I10").Value = Current_Vol
            Exit For
        End If
        
        ' If the above check fails, calculate the vegas for the European Calls with the option
        ' parameters input and both the Long_K and Short_K variables with the Current_Vol. Then
        ' determine the vega of the Bull Spread using the Current_Vol by taking the difference
        ' between the European Call price using the Long_L and the European Call price using the
        ' Short_K.
        Long_vega = BS_Euro_Call_Vega(S0, Long_K, rd, rf, Current_Vol, T)
        Short_vega = BS_Euro_Call_Vega(S0, Short_K, rd, rf, Current_Vol, T)
        Imp_vega = Long_vega - Short_vega

        ' Check if the vega found above equals 0. If it does, output an error message to the
        ' Excel worksheet and exit the for loop.
        If (Imp_vega = 0) Then
            Range("I10").Value = "Did not converge, implied vega equal to 0"
            Exit For
        End If
        
        ' Calculate the next Current_Vol and continue to the next iteration of the loop.
        Current_Vol = Current_Vol - ((C_imp - C_mrkt) / Imp_vega)
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print an
    ' error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("I10").Value = "Did not converge to " & Tol & " tolerance. (More than " & MaxIter & " iterations needed)"
    End If
    
End Function

'-----------------------------------------------------------------------------------------
' Bull_Spread_Regula_Falsi() function to calculate the implied volatility of a Bull Spread
' with given input parameters and market price using the Regula Falsi method. Also takes 2
' guesses for the implied vol as inputs x1_Vol and x2_Vol, as well as the maximum number
' of iterations to allow the for loop within the function to perform and the tolerance to
' which the user wants the output to be accurate to.
'-----------------------------------------------------------------------------------------
Function Bull_Spread_Regula_Falsi(S0, Long_K, Short_K, rd, rf, T, C_mrkt, x1_Vol, x2_Vol, MaxIter, Tol)
    
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the x1_Vol and x2_Vol variables with the Long_K. Then determine the price of
        ' the Bull Spread using the Long_K by taking the difference between the European
        ' Call price using the x1_Vol and the European Call price using the x2_Vol.
        Long_x1_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, x1_Vol, T)
        Short_x1_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, x1_Vol, T)
        x1_C = Long_x1_C - Short_x1_C
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the x1_Vol and x2_Vol variables with the Short_K. Then determine the price of
        ' the Bull Spread using the Short_K by taking the difference between the European
        ' Call price using the x1_Vol and the European Call price using the x2_Vol.
        Long_x2_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, x2_Vol, T)
        Short_x2_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, x2_Vol, T)
        x2_C = Long_x2_C - Short_x2_C
        
        ' Determine the implied volatility from the previously calculated Bull Spread
        ' prices and the current x1_Vol and x2_Vol variables. Then calculate the prices of
        ' the European Call option with the input parameters, the Est_Vol, and both the
        ' Short_K and Long_K. Lastly, determine the Bull Spread price from these
        ' European Call prices.
        Est_Vol = x1_Vol - ((x1_C - C_mrkt) * ((x2_Vol - x1_Vol) / (x2_C - x1_C)))
        Long_Est_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, Est_Vol, T)
        Short_Est_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, Est_Vol, T)
        Est_C = Long_Est_C - Short_Est_C
        
        ' Check if the difference between the Bull Spread price determined using the Est_Vol
        ' and the user input market price of the Bull Spread is within the user input
        ' input tolerance. If it is, then output the Est_Vol to the Excel spreadsheet and
        ' exit the for loop.
        If (Abs(Est_C - C_mrkt) < Tol) Then
            Range("I11").Value = Est_Vol
            Exit For
        End If
        
        ' If the the difference between the Bull Spread price calculated with the input
        ' option parameters and the Est_Vol and the market Bull Spread price is greater
        ' than the input tolerance, check if the sign of the Bull Spread price
        ' calculated with the Est_Vol matches the sign of the Bull Spread price
        ' calculated with the current x1_Vol. If the signs match, assign the value of
        ' Est_Vol to x1_Vol and continue to the next iteration of this for loop.
        If (Sgn(Est_C) = Sgn(x1_C)) Then
            x1_Vol = Est_Vol
        ' If the above check fails, check if the sign of the Bull Spread price calculated
        ' with the Est_Vol matches the sign of the Bull Spread price calculated with the
        ' current x2_Vol. If the signs match, assign the value of Est_Vol to x2_Vol and
        ' continue to the next iteration of this for loop.
        ElseIf (Sgn(Est_C) = Sgn(x2_C)) Then
            x2_Vol = Est_Vol
        ' If both of the above checks fail, output an error message to the Excel worksheet.
        Else
            Range("I11").Value = "Did not converge (Est call price estimate sign matches neither initial call price estimate's sign)"
        End If
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print an
    ' error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("I11").Value = "Did not converge to " & Tol & " tolerance. (More than " & MaxIter & " iterations needed)"
    End If
    
End Function

'-------------------------------------------------------------------------------------------
' Euro_Call_Secant() function to calculate the implied volatility of a European Call
' option with given input parameters and market price using the Secant method. Also takes
' 2 guesses for the implied vol as inputs x1_Vol and x2_Vol, as well as the maximum
' number of iterations to allow the for loop within the function to perform and the
' tolerance to which the user wants the output to be accurate to.
'-------------------------------------------------------------------------------------------
Function Bull_Spread_Secant(S0, Long_K, Short_K, rd, rf, T, C_mrkt, x1_Vol, x2_Vol, MaxIter, Tol)
    
    ' Initialize the for loop from 1 to the maximum number of iterations input by the user.
    For Iter = 1 To MaxIter
    
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the x1_Vol and x2_Vol variables with the Long_K. Then determine the price of
        ' the Bull Spread using the Long_K by taking the difference between the European
        ' Call price using the x1_Vol and the European Call price using the x2_Vol.
        Long_x1_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, x1_Vol, T)
        Short_x1_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, x1_Vol, T)
        x1_C = Long_x1_C - Short_x1_C
        
        ' Calculate the prices for the European Calls with the option parameters input and
        ' both the x1_Vol and x2_Vol variables with the Short_K. Then determine the price of
        ' the Bull Spread using the Short_K by taking the difference between the European
        ' Call price using the x1_Vol and the European Call price using the x2_Vol.
        Long_x2_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, x2_Vol, T)
        Short_x2_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, x2_Vol, T)
        x2_C = Long_x2_C - Short_x2_C
        
        ' Determine the implied volatility from the previously calculated Bull Spread
        ' prices and the current x1_Vol and x2_Vol variables. Then calculate the prices of
        ' the European Call option with the input parameters, the Est_Vol, and both the
        ' Short_K and Long_K. Lastly, determine the Bull Spread price from these
        ' European Call prices.
        Est_Vol = ((x1_Vol * (x2_C - C_mrkt)) - (x2_Vol * (x1_C - C_mrkt))) / (x2_C - x1_C)
        Long_Est_C = BS_Euro_Call_Price(S0, Long_K, rd, rf, Est_Vol, T)
        Short_Est_C = BS_Euro_Call_Price(S0, Short_K, rd, rf, Est_Vol, T)
        Est_C = Long_Est_C - Short_Est_C
        
        ' Check if the difference between the Bull Spread price determined using the Est_Vol
        ' and the user input market price of the Bull Spread is within the user input
        ' input tolerance. If it is, then output the Est_Vol to the Excel spreadsheet and
        ' exit the for loop.
        If (Abs(Est_C - C_mrkt) < Tol) Then
            Range("I12").Value = Est_Vol
            Exit For
        End If
        
        ' If the the difference between the Bull Spread price calculated with the input
        ' option parameters and the Est_Vol and the market Bull Spread price is greater
        ' than the input tolerance, check if the sign of the Bull Spread price
        ' calculated with the Est_Vol matches the sign of the Bull Spread price
        ' calculated with the current x1_Vol. If the signs match, assign the value of
        ' Est_Vol to x1_Vol and continue to the next iteration of this for loop.
        If (Sgn(Est_C) = Sgn(x1_C)) Then
            x1_Vol = Est_Vol
        ' If the above check fails, check if the sign of the Bull Spread price calculated
        ' with the Est_Vol matches the sign of the Bull Spread price calculated with the
        ' current x2_Vol. If the signs match, assign the value of Est_Vol to x2_Vol and
        ' continue to the next iteration of this for loop.
        ElseIf (Sgn(Est_C) = Sgn(x2_C)) Then
            x2_Vol = Est_Vol
        ' If both of the above checks fail, output an error message to the Excel worksheet.
        Else
            Range("I12").Value = "Did not converge (Current call price estimate sign matches neither initial call price estimate's sign)"
        End If
        
    Next Iter
    
    ' If the maximum number of iterations is reached before the for loop is broken, print an
    ' error message to the Excel spreadsheet.
    If (Iter = MaxIter + 1) Then
        Range("I12").Value = "WARNING: Did not converge to " & Tol & " tolerance. (" & MaxIter & " iterations performed)"
    End If
    
End Function

'-----------------------------------------------------------------------------------------
' Bull_Spread_Calcs() sub procedure that:
'   - Takes in the Bull Spread parameters input by the user in the Excel worksheet, as
'     well as the user input market price of the Bull Spread, the initial guesses to be
'     used for all methods of determining the implied volatility of the Bull Spread
'     (Newton Raphson, Regula Falsi, and Secant methods), the maximum number of
'     iterations to use for all 3 methods, and the tolerance to be used for all 3 methods.
'
'   - Calculates the implied volatility of the Bull Spread with the user input parameters
'     and market price using all 3 method functions defined above.
'-----------------------------------------------------------------------------------------
Sub Bull_Spread_Calcs()

    ' Take in the Bull Spread parameters input by the user in the Excel worksheet cells
    ' B15:B21.
    S0 = Range("B15").Value
    Long_K = Range("B16").Value
    Short_K = Range("B17").Value
    rd = Range("B18").Value
    rf = Range("B19").Value
    T = Range("B21").Value
    
    ' Take in the European Call price and volatility guesses input by the user in the Excel
    ' worksheet cells I3:I6 and J5:J6.
    C_mrkt = Range("I3").Value
    Current_Vol = Range("I4").Value
    Low_Vol = Range("I5").Value
    High_Vol = Range("J5").Value
    x1_Vol = Range("I6").Value
    x2_Vol = Range("J6").Value
    
    ' Take in the user input maximum number of iterations and tolerance to use for the
    ' following methods of determining the implied volatility from cells B3:B4.
    MaxIter = Range("B3").Value
    Tol = 10 ^ Range("B4").Value

    ' Call the previously defined functions for finding the implied volatility of a Bull
    ' Spread using the inputs taken from the Excel spreadsheet above.
    Bull_Spread_Newton = Bull_Spread_Newt_Raph(S0, Long_K, Short_K, rd, rf, T, C_mrkt, Current_Vol, MaxIter, Tol)
    Bull_Spread_Reg_Fal = Bull_Spread_Regula_Falsi(S0, Long_K, Short_K, rd, rf, T, C_mrkt, Low_Vol, High_Vol, MaxIter, Tol)
    Bull_Spread_Sec_Meth = Bull_Spread_Secant(S0, Long_K, Short_K, rd, rf, T, C_mrkt, x1_Vol, x2_Vol, MaxIter, Tol)
    
End Sub