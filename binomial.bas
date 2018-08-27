Attribute VB_Name = "Binomial"
' GENERAL ADDITIVE BINOMIAL TREE MODEL (GABTM)

 ' output is an array of {call.value,put.value} or for 'touch' type: {onetouch.value,notouch.value}
   ' approximation for the process: dS=(rd - rf).S dt + sigma.S dz
   ' where dz describes a Weiner process random walk.
   '
   ' Assuming S is distributed log-normally as x = ln(S), then:
   ' dx = nu dt + sigma dz
   ' where: nu = r - 1/2.sigma^2
   '
   ' Approximating this with by discrete time steps of dt = t / N
   ' and defining up and down changes in x by:
   '  x + dxu , x - dxd, where dxu = -dxd = dx
   '
   ' and associated probabilties for up and down movements of pu and pd, where pu + pd = 1
   ' Equating mean and variance gives:
   '
   ' E[dx] = pu.dx - pd.dx = nu.dt
   ' E[dx^2] = pu.dx^2 + pd.dx^2 = sigma^2.dt + nu^2.dt^2
   '
   ' solving simultaneously gives:
   '
   ' dx = sqrt( sigma^2.dt + nu^2.dt )
   ' pu = 1/2 + 1/2 . nu.dt / dx
   ' and we know that pd = 1 -pu
   '
   ' with the spot price S on node(i,j) computed as:
   ' St(i,j) = exp(x(i,j)) = exp(x + j.dx - (i-j).dx)
   '
   ' terminal option values V are given on the final j nodes as:
   ' vanilla_call = V(j,1) = Max( 0, St(j) - K )
   ' vanilla_put = V(j,2) = Max( 0, K - St(j) )
   ' call = V(j,1) = 1 if St(j) >= K ; 0 otherwise
   ' put = V(j,2) = 1 if St(j) < K ; 0 otherwise


Option Explicit

Function GABTM(OptionType As String, S As Double, K As Double, t As Double, rd As Double, rf As Double, sigma As Double, N As Double)
    ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        GABTM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
        
        Dim dt As Double, nu As Double, dxu As Double, dxd As Double, pu As Double, pd As Double, disc As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To N), V(0 To N, 1 To 2)
        Dim i As Integer, j As Integer
         Dim temp(1 To 2) As Variant
        
        'compute coefficients and constants
        dt = t / N
        nu = rd - rf - 0.5 * sigma ^ 2
        dxu = Sqr(sigma ^ 2 * dt + (nu * dt) ^ 2)
        dxd = -dxu
        pu = 1 / 2 + 1 / 2 * (nu * dt / dxu)
        pd = 1 - pu
        disc = Exp(-rd * dt)
    
        'initialise asset prices and option values at maturity N
        St(0) = S * Exp(N * dxd)
        For j = 0 To N Step 1
            If j > 0 Then St(j) = St(j - 1) * Exp(dxu - dxd)
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(j, 1) = .Max(0, St(j) - K)
                    V(j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(j, 1) = 1 Else V(j, 1) = 0
                If St(j) < K Then V(j, 2) = 1 Else V(j, 2) = 0
            End If
        Next j
        
        'step back through tree to compute option values through to time zero
        For i = N - 1 To 0 Step -1
            For j = 0 To i Step 1
                If j > 0 Then St(j) = St(j - 1) * Exp(dxu - dxd) Else St(0) = S * Exp(i * dxd)
                
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(j, 1) = 1 Else V(j, 1) = disc * (pu * V(j + 1, 1) + pd * V(j, 1))     ' touch up
                    If St(j) <= K Then V(j, 2) = 1 Else V(j, 2) = disc * (pu * V(j + 1, 2) + pd * V(j, 2))     ' touch down
                Else 'for european style vanilla and binary
                    V(j, 1) = disc * (pu * V(j + 1, 1) + pd * V(j, 1))
                    V(j, 2) = disc * (pu * V(j + 1, 2) + pd * V(j, 2))
                End If
                
            Next j
        Next i
    
        If OptionType = "touch" Then
            'choose touch up or touch down depening on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, 1) Else temp(1) = V(0, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, 1)
            temp(2) = V(0, 2)
        End If
    
        GABTM = temp
    
    
    End If
    
End Function



Sub GABT_test_convergence()

    Sheets("Binomial").Activate
    Range("GABT_convergence_test").CurrentRegion.Select
    Selection.ClearContents
    
    Dim i As Integer, j As Integer
    For i = Range("GABT_parameters").Cells(2, 1).Value To Range("GABT_parameters").Cells(3, 1).Value Step Range("GABT_parameters").Cells(4, 1).Value
        j = j + 1
        Range("GABT_N") = i
        Calculate
        Range("GABT_convergence_test").Cells(j, 1) = i
        Range("GABT_convergence_test").Cells(j, 2) = Range("GABT_" & Range("GABT_parameters").Cells(1, 1) & "_result").Cells(1, 1).Value
        Range("GABT_convergence_test").Cells(j, 3) = Range("GABT_" & Range("GABT_parameters").Cells(1, 1) & "_result").Cells(2, 1).Value
    Next i
    
End Sub

Function TVBTM(OptionType As String, S As Double, K As Double, t_rng As Range, rd_rng As Range, rf_rng As Range, sigma_rng As Range, N As Double)
    ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        TVBTM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
              
        Dim TimeValue As Double, dt_bar As Double, sigma_bar As Double, nu_bar As Double
        Dim dt As Double, pu As Double, pd As Double, nu As Double, disc As Double
        Dim dxu As Double, dxd As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To N), V(0 To N, 1 To 2)
        Dim i As Integer, j As Integer, x As Integer
        Dim temp(1 To 2) As Variant
        Dim temp_sigma As Double
        
        'convert input ranges to arrays
        Dim t As Variant, rd As Variant, rf As Variant, sigma As Variant
        Dim t_temp As Variant, rd_temp As Variant, rf_temp As Variant, sigma_temp As Variant
        
        t_temp = t_rng
        If UBound(t_temp, 2) > UBound(t_temp, 1) Then t = TransposeArray(t_temp) Else t = t_temp
        rd_temp = rd_rng
        If UBound(rd_temp, 2) > UBound(rd_temp, 1) Then rd = TransposeArray(rd_temp) Else rd = rd_temp
        rf_temp = rf_rng
        If UBound(rf_temp, 2) > UBound(rf_temp, 1) Then rf = TransposeArray(rf_temp) Else rf = rd_temp
        sigma_temp = sigma_rng
        If UBound(sigma_temp, 2) > UBound(sigma_temp, 1) Then sigma = TransposeArray(sigma_temp) Else sigma = sigma_temp
        

        
        
        'compute initial coefficients and constants
        TimeValue = t(UBound(t, 1), 1)
        dt_bar = TimeValue / N
        sigma_bar = 0
        nu_bar = 0
        
        


        For x = 1 To N - 1
            temp_sigma = LookupValue(x * dt_bar, t, sigma) / (N - 1)
            sigma_bar = sigma_bar + temp_sigma
            nu_bar = LookupValue(x * dt_bar, t, rd) - LookupValue(x * dt_bar, t, rf) - 0.5 * temp_sigma ^ 2
        Next x
        
        dxu = Sqr(sigma_bar ^ 2 * dt_bar + (nu_bar * dt_bar) ^ 2)
        dxd = -dxu
        

    
        'initialise asset prices and option values at maturity N
        St(0) = S * Exp(N * dxd)
        For j = 0 To N Step 1
            If j > 0 Then St(j) = St(j - 1) * Exp(dxu - dxd)
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(j, 1) = .Max(0, St(j) - K)
                    V(j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(j, 1) = 1 Else V(j, 1) = 0
                If St(j) < K Then V(j, 2) = 1 Else V(j, 2) = 0
            End If
        Next j
        

        
        
        'step back through tree to compute option values through to time zero
        For i = (N - 1) To 0 Step -1
            For j = 0 To i Step 1
            
                'compute tree branching variables
                TimeValue = (i + 1) * dt_bar  'needs an approximate starting point for looking up the appropriate variables
                temp_sigma = LookupValue(TimeValue, t, sigma)
                nu = LookupValue(TimeValue, t, rd) - LookupValue(TimeValue, t, rf) - 0.5 * temp_sigma ^ 2
                dt = (-(temp_sigma ^ 2) + Sqr(temp_sigma ^ 4 + 4 * nu ^ 2 * dxu ^ 2)) / (2 * nu ^ 2)
                
                pu = 1 / 2 + 1 / 2 * nu * dt / dxu
                pd = 1 - pu
                disc = Exp(-LookupValue(TimeValue, t, rd) * dt)
                If j > 0 Then St(j) = St(j - 1) * Exp(dxu - dxd) Else St(0) = S * Exp(i * dxd)
                
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(j, 1) = 1 Else V(j, 1) = disc * (pu * V(j + 1, 1) + pd * V(j, 1))     ' touch up
                    If St(j) <= K Then V(j, 2) = 1 Else V(j, 2) = disc * (pu * V(j + 1, 2) + pd * V(j, 2))     ' touch down
                Else 'for european style vannila and binary
                    V(j, 1) = disc * (pu * V(j + 1, 1) + pd * V(j, 1))
                    V(j, 2) = disc * (pu * V(j + 1, 2) + pd * V(j, 2))
                End If
                
            Next j
        Next i
    
    
    
        If OptionType = "touch" Then
            'choose touch up or touch down depending on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, 1) Else temp(1) = V(0, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, 1)
            temp(2) = V(0, 2)
        End If
    
        TVBTM = temp
    
    End If
    
End Function






Function LookupValue(TimeValue As Double, TimeArray As Variant, ValueArray As Variant)
        'looks up nearest lower value in TimeArray and returns corresponding ValueArray value
        
        Dim i As Integer
        Dim x As Variant, y As Variant
        Dim x_temp As Variant, y_temp As Variant
        Dim temp As Double
        
        x_temp = TimeArray
        y_temp = ValueArray
        
        If UBound(x_temp, 2) > UBound(x_temp, 1) Then x = TransposeArray(x_temp) Else x = x_temp
        If UBound(y_temp, 2) > UBound(y_temp, 1) Then y = TransposeArray(y_temp) Else y = y_temp
               
        If TimeValue > x(UBound(x, 1), 1) Then TimeValue = x(UBound(x, 1), 1)
               
        i = 0
        Do
            i = i + 1
            temp = y(i, 1)
        Loop While i <= UBound(x, 1) And TimeValue > x(i, 1)
        
        LookupValue = temp
    
End Function


Public Function TransposeArray(myarray As Variant) As Variant
Dim x As Long
Dim y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For x = 1 To Xupper
        For y = 1 To Yupper
            tempArray(x, y) = myarray(y, x)
        Next y
    Next x
    TransposeArray = tempArray
End Function


