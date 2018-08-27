Attribute VB_Name = "Trinomial"
' TRINOMIAL TREE MODEL (TTM)

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
   ' and associated probabilties for up, unchanged and down movements of pu, pm and pd, where pu + pm + pd = 1
   ' Equating mean and variance gives:
   '
   ' E[dx] = pu.dx + pm.(0) + pd.-dx = nu.dt
   ' E[dx^2] = pu.dx^2 + pm.(0) + pd.-dx^2 = sigma^2.dt + v^2.dt^2
   '
   ' solving simultaneously gives:
   '
   ' dx = sqrt( sigma^2.dt + nu^2.dt )
   ' pu = 1/2.[ (sigma^2.dt +nu^2.dt^2)/dx^2 + 1/2 . nu.dt / dx]
   ' pm = 1 - (sigma^2 + nu^2.dt^2)/dt^2
   ' pu = 1/2.[ (sigma^2.dt +nu^2.dt^2)/dx^2 - 1/2 . nu.dt / dx]
   '
   ' tree node branching:
   '
   '          | (i+1,j+2)   : S up with probability pu
   '  ( i,j ) | (i+1,j+1)   : S equal with probablility pm
   '          | (i+1),j)      : S down with pd
   '
   ' mapped to arrays as:
   '                                                | (N,N)    -> all up
   '                                                | .....
   '                                           .... | (N,8)
   '                                           .... | (N,7)
   '                               | (3,6) | .... | (N,6)
   '                               | (3,5) | .... | (N,5)
   '                     | (2,4) | (3,4) | .... | (N,4)
   '                     | (2,3) | (3,3) | .... | (N,3)
   '           | (1,2) | (2,2) | (3,2) | .... | (N,2)                       |dx
   '           | (1,1) | (2,1) | (3,1) | .... | (N,1)
   '   (0,0) | (1,0) | (2,0) | (3,0) | .... | (N,0)   -> all down
   '
   '           |--->|dt
   '
   '
   ' with the spot price S on node(i,j) computed as:
   ' St(i,j) = S * Exp(-i * dx - j * dx)
   '
   '
   ' terminal option values V are given on the final j nodes as:
   ' vanilla_call = V(i,j,1) = Max( 0, St(j) - K )
   ' vanilla_put = V(i,j,2) = Max( 0, K - St(j) )
   ' call = V(i,j,1) = 1 if St(i,j) >= K ; 0 otherwise
   ' put = V(i,j,2) = 1 if St(i,j) < K ; 0 otherwise
'
'
'


Option Explicit


Function TTM(OptionType As String, S As Double, K As Double, t As Double, rd As Double, rf As Double, sigma As Double, N As Double, dx As Double)

 ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        TTM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
        
        Dim dt As Double, nu As Double, edx As Double, pu As Double, pm As Double, pd As Double, disc As Double
        Dim Z As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To 2 * N), V(0 To N, 0 To 2 * N, 1 To 2)
        Dim i As Integer, j As Integer
        Dim temp(1 To 2) As Variant
        
        'compute coefficients and constants
        dt = t / N
        nu = rd - rf - 0.5 * sigma ^ 2
        edx = Exp(dx)
        pu = 0.5 * ((sigma ^ 2 * dt + nu ^ 2 * dt ^ 2) / dx ^ 2 + nu * dt / dx)
        pm = 1 - (sigma ^ 2 * dt + nu ^ 2 * dt ^ 2) / dx ^ 2
        pd = 0.5 * ((sigma ^ 2 * dt + nu ^ 2 * dt ^ 2) / dx ^ 2 - nu * dt / dx)
        disc = Exp(-rd * dt)

        'initialise asset prices and option values at maturity N

        St(0) = S * Exp(-N * dx)
        For j = 0 To 2 * N Step 1
            If j > 0 Then St(j) = St(j - 1) * edx
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(N, j, 1) = .Max(0, St(j) - K)
                    V(N, j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(N, j, 1) = 1 Else V(N, j, 1) = 0
                If St(j) < K Then V(N, j, 2) = 1 Else V(N, j, 2) = 0
            End If
        Next j
        
        'step back through tree to compute option values through to time zero
        For i = N - 1 To 0 Step -1
            For j = 0 To 2 * i Step 1
                If j > 0 Then St(j) = St(j - 1) * edx Else St(0) = S * Exp(-i * dx)
                
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(i, j, 1) = 1 Else V(i, j, 1) = disc * (pu * V(i + 1, j + 2, 1) + pm * V(i + 1, j + 1, 1) + pd * V(i + 1, j, 1)) ' touch up
                    If St(j) <= K Then V(i, j, 2) = 1 Else V(i, j, 2) = disc * (pu * V(i + 1, j + 2, 2) + pm * V(i + 1, j + 1, 2) + pd * V(i + 1, j, 2)) ' touch down
                Else 'for european style vanilla and binary
                    V(i, j, 1) = disc * (pu * V(i + 1, j + 2, 1) + pm * V(i + 1, j + 1, 1) + pd * V(i + 1, j, 1))
                    V(i, j, 2) = disc * (pu * V(i + 1, j + 2, 2) + pm * V(i + 1, j + 1, 2) + pd * V(i + 1, j, 2))
                End If
                
            Next j
        Next i
    
        If OptionType = "touch" Then
            'choose touch up or touch down depending on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, 0, 1) Else temp(1) = V(0, 0, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, 0, 1)
            temp(2) = V(0, 0, 2)
        End If
    
        TTM = temp
    
    End If
    
End Function




Function EFDM(OptionType As String, S As Double, K As Double, t As Double, rd As Double, rf As Double, sigma As Double, N As Double, Nj As Double, dx As Double)
'EXPLICIT FINITE DIFFERENCE MODEL
 ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        EFDM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
        
        Dim dt As Double, nu As Double, edx As Double, pu As Double, pm As Double, pd As Double, disc As Double
        Dim Z As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To 2 * N), V(0 To N, 0 To 2 * N, 1 To 2)
        Dim i As Integer, j As Integer
        Dim temp(1 To 2) As Variant
        
        'compute coefficients and constants
        dt = t / N
        nu = rd - rf - 0.5 * sigma ^ 2
        edx = Exp(dx)
        pu = 0.5 * dt * ((sigma / dx) ^ 2 + nu / dx)
        pm = 1 - dt * (sigma / dx) ^ 2 - rd * dt
        pd = 0.5 * dt * ((sigma / dx) ^ 2 - nu / dx)
        disc = Exp(-rd * dt)

        'initialise asset prices and option values at maturity N

        St(0) = S * Exp(-Nj * dx)
        For j = 1 To 2 * Nj Step 1
            If j > 0 Then St(j) = St(j - 1) * edx
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(N, j, 1) = .Max(0, St(j) - K)
                    V(N, j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(N, j, 1) = 1 Else V(N, j, 1) = 0
                If St(j) < K Then V(N, j, 2) = 1 Else V(N, j, 2) = 0
            End If
        Next j
        
        'step back through tree to compute option values through to time zero
        For i = N - 1 To 0 Step -1

            For j = 1 To 2 * Nj - 1 Step 1
                If j > 0 Then St(j) = St(j - 1) * edx Else St(0) = S * Exp(-i * dx)
                
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(i, j, 1) = 1 Else V(i, j, 1) = disc * (pu * V(i + 1, j + 1, 1) + pm * V(i + 1, j, 1) + pd * V(i + 1, j - 1, 1)) ' touch up
                    If St(j) <= K Then V(i, j, 2) = 1 Else V(i, j, 2) = disc * (pu * V(i + 1, j + 1, 2) + pm * V(i + 1, j, 2) + pd * V(i + 1, j - 1, 2)) ' touch down
                Else 'for european style vanilla and binary
                    V(i, j, 1) = (pu * V(i + 1, j + 1, 1) + pm * V(i + 1, j, 1) + pd * V(i + 1, j - 1, 1))
                    V(i, j, 2) = (pu * V(i + 1, j + 1, 2) + pm * V(i + 1, j, 2) + pd * V(i + 1, j - 1, 2))
                End If
            Next j
        'boundary conditions
        V(i, 0, 1) = V(i, 1, 1)
        V(i, 0, 2) = V(i, 1, 2)
        V(i, 2 * Nj, 1) = V(i, 2 * Nj - 1, 1) + (St(2 * Nj) - St(2 * Nj - 1))
        V(i, 2 * Nj, 2) = V(i, 2 * Nj - 1, 2) + (St(2 * Nj) - St(2 * Nj - 1))

        Next i
    
        If OptionType = "touch" Then
            'choose touch up or touch down depending on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, Nj, 1) Else temp(1) = V(0, Nj, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, Nj, 1)
            temp(2) = V(0, Nj, 2)
        End If
    
        EFDM = temp
    
    End If
    
End Function


Function IFDM(OptionType As String, S As Double, K As Double, t As Double, rd As Double, rf As Double, sigma As Double, N As Double, Nj As Double, dx As Double)
'IMPLICIT FINITE DIFFERENCE MODEL
 ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        IFDM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
        
        Dim dt As Double, nu As Double, edx As Double, pu As Double, pm As Double, pd As Double, disc As Double
        Dim Z As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To 2 * N), V(0 To N, 0 To 2 * N, 1 To 2)
        Dim i As Integer, j As Integer
        Dim temp(1 To 2) As Variant
        Dim lamda_L As Double, lamda_U As Double
        Dim pmp As Variant, pp As Variant
        ReDim pmp(0 To 2 * Nj)
        ReDim pp(0 To 2 * Nj, 1 To 2)
        
        
        'compute coefficients and constants
        dt = t / N
        nu = rd - rf - 0.5 * sigma ^ 2
        edx = Exp(dx)
        pu = -0.5 * dt * ((sigma / dx) ^ 2 + nu / dx)
        pm = 1 + dt * (sigma / dx) ^ 2 + rd * dt
        pd = -0.5 * dt * ((sigma / dx) ^ 2 - nu / dx)
        disc = Exp(-rd * dt)

        'initialise asset prices and option values at maturity N

        St(0) = S * Exp(-Nj * dx)
        For j = 1 To 2 * Nj Step 1
            If j > 0 Then St(j) = St(j - 1) * edx
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(N, j, 1) = .Max(0, St(j) - K)
                    V(N, j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(N, j, 1) = 1 Else V(N, j, 1) = 0
                If St(j) < K Then V(N, j, 2) = 1 Else V(N, j, 2) = 0
            End If
        Next j
        
        'compute derivative boundary condition
        lamda_L = -1 * (St(1) - St(0))
        lamda_U = 0
        
        'step back through lattice to compute option values through to time zero
        For i = N - 1 To 0 Step -1
        
            'solve tridiagonal system
            pmp(1) = pm + pd
            pp(1, 1) = V(i + 1, 1, 1) + pd * lamda_L
            pp(1, 2) = V(i + 1, 1, 2) + pd * lamda_L
                'eliminate upper diagonal
                For j = 2 To 2 * Nj - 1
                    pmp(j) = pm - pu * pd / pmp(j - 1)
                    pp(j, 1) = V(i + 1, j, 1) - pp(j - 1, 1) * pd / pmp(j - 1)
                    pp(j, 2) = V(i + 1, j, 2) - pp(j - 1, 2) * pd / pmp(j - 1)
                Next j
                    'use boundary condition at j = 2 * Nj and equation at j = 2 * Nj -1
                    V(i, 2 * Nj, 1) = (pp(2 * Nj - 1, 1) + pmp(2 * Nj - 1) * lamda_U) / (pu + pmp(2 * Nj - 1))
                    V(i, 2 * Nj, 2) = (pp(2 * Nj - 1, 2) + pmp(2 * Nj - 1) * lamda_U) / (pu + pmp(2 * Nj - 1))
                    V(i, 2 * Nj - 1, 1) = V(i + 1, 2 * Nj, 1) - lamda_U
                    V(i, 2 * Nj - 1, 2) = V(i + 1, 2 * Nj, 2) - lamda_U
                        'back substitution
                        For j = 2 * Nj - 1 To 1 Step -1
                            V(i, j, 1) = (pp(j, 1) - pu * V(i, j + 1, 1)) / pmp(j)
                            V(i, j, 2) = (pp(j, 2) - pu * V(i, j + 1, 2)) / pmp(j)
                        Next j
                        V(i, 0, 1) = V(i, 1, 1) - lamda_L
                        V(i, 0, 2) = V(i, 1, 2) - lamda_L
                        
            'calculate expected values on each node of the lattice, applying early exercise conditions for touch options
            For j = 1 To 2 * Nj - 1 Step 1
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(i, j, 1) = 1 ' touch up
                    If St(j) <= K Then V(i, j, 2) = 1 ' touch down
                End If
            Next j

        Next i

 
        'Return final option value from the lattice
        If OptionType = "touch" Then
            'choose touch up or touch down depending on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, Nj, 1) Else temp(1) = V(0, Nj, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, Nj, 1)
            temp(2) = V(0, Nj, 2)
        End If
    
        IFDM = temp
    
    End If
    
End Function


Function CNFDM(OptionType As String, S As Double, K As Double, t As Double, rd As Double, rf As Double, sigma As Double, N As Double, Nj As Double, dx As Double)
'CRANK-NICOLSON FINITE DIFFERENCE MODEL
 ' OptionType inputs: "vanilla" , "binary" , "touch"
    ' check validity
    OptionType = LCase(OptionType)
    If Not (OptionType = "vanilla" Or OptionType = "binary" Or OptionType = "touch") Then
        CNFDM = "Invalid Option Type. Select one of: vanilla, binary, touch"
        Else
        
        Dim dt As Double, nu As Double, edx As Double, pu As Double, pm As Double, pd As Double, disc As Double
        Dim Z As Double
        Dim St As Variant, V As Variant
        ReDim St(0 To 2 * N), V(0 To N, 0 To 2 * N, 1 To 2)
        Dim i As Integer, j As Integer
        Dim temp(1 To 2) As Variant
        Dim lamda_L As Double, lamda_U As Double
        Dim pmp As Variant, pp As Variant
        ReDim pmp(0 To 2 * Nj)
        ReDim pp(0 To 2 * Nj, 1 To 2)
        
        
        'compute coefficients and constants
        dt = t / N
        nu = rd - rf - 0.5 * sigma ^ 2
        edx = Exp(dx)
        pu = -0.25 * dt * ((sigma / dx) ^ 2 + nu / dx)
        pm = 1 + 0.5 * dt * (sigma / dx) ^ 2 + 0.5 * rd * dt
        pd = -0.25 * dt * ((sigma / dx) ^ 2 - nu / dx)
        disc = Exp(-rd * dt)

        'initialise asset prices and option values at maturity N

        St(0) = S * Exp(-Nj * dx)
        For j = 0 To 2 * Nj Step 1
            If j > 0 Then St(j) = St(j - 1) * edx
            If OptionType = "vanilla" Then
                With Application.WorksheetFunction
                    V(N, j, 1) = .Max(0, St(j) - K)
                    V(N, j, 2) = .Max(0, K - St(j))
                End With
            Else 'for binary & touch
                If St(j) >= K Then V(N, j, 1) = 1 Else V(N, j, 1) = 0
                If St(j) < K Then V(N, j, 2) = 1 Else V(N, j, 2) = 0
            End If
        Next j
        
        'compute derivative boundary condition
        lamda_L = -1 * (St(1) - St(0))
        lamda_U = 0
        
        'step back through lattice to compute option values through to time zero
        For i = N - 1 To 0 Step -1
        
            'solve Crank-Nicolson tridiagonal system
            'substitute boundary condition at j=0 into j=1
            pmp(0) = pm + pd
            pmp(1) = pm + pd
            pp(1, 1) = -pu * V(i + 1, 2, 1) - (pm - 2) * V(i + 1, 1, 1) - pd * V(i + 1, 0, 1) + pd * lamda_L
            pp(1, 2) = -pu * V(i + 1, 2, 2) - (pm - 2) * V(i + 1, 1, 2) - pd * V(i + 1, 0, 2) + pd * lamda_L
                'eliminate upper diagonal
                For j = 2 To 2 * Nj - 1
                    pmp(j) = pm - pu * pd / pmp(j - 1)
                    pp(j, 1) = -pu * V(i + 1, j + 1, 1) - (pm - 2) * V(i + 1, j, 1) - pd * V(i + 1, j - 1, 1) - pp(j - 1, 1) * pd / pmp(j - 1)
                    pp(j, 2) = -pu * V(i + 1, j + 1, 2) - (pm - 2) * V(i + 1, j, 2) - pd * V(i + 1, j - 1, 2) - pp(j - 1, 2) * pd / pmp(j - 1)
                Next j
                    'use boundary condition at j = 2 * Nj and equation at j = 2 * Nj -1
                    V(i, 2 * Nj, 1) = (pp(2 * Nj - 1, 1) + pmp(2 * Nj - 1) * lamda_U) / (pu + pmp(2 * Nj - 1))
                    V(i, 2 * Nj, 2) = (pp(2 * Nj - 1, 2) + pmp(2 * Nj - 1) * lamda_U) / (pu + pmp(2 * Nj - 1))
                    V(i, 2 * Nj - 1, 1) = V(i + 1, 2 * Nj, 1) - lamda_U
                    V(i, 2 * Nj - 1, 2) = V(i + 1, 2 * Nj, 2) - lamda_U
                        'back substitution
                        For j = 2 * Nj - 1 To 1 Step -1
                            V(i, j, 1) = (pp(j, 1) - pu * V(i, j + 1, 1)) / pmp(j)
                            V(i, j, 2) = (pp(j, 2) - pu * V(i, j + 1, 2)) / pmp(j)
                        Next j
                        V(i, 0, 1) = V(i, 1, 1) - lamda_L
                        V(i, 0, 2) = V(i, 1, 2) - lamda_L
                        
            'calculate expected values on each node of the lattice, applying early exercise conditions for touch options
            For j = 0 To 2 * Nj Step 1
                If OptionType = "touch" Then 'for american style
                    If St(j) >= K Then V(i, j, 1) = 1 ' touch up
                    If St(j) <= K Then V(i, j, 2) = 1 ' touch down
                End If
            Next j

        Next i

 
        'Return final option value from the lattice
        If OptionType = "touch" Then
            'choose touch up or touch down depending on direction of barrier from spot price
            If K >= S Then temp(1) = V(0, Nj, 1) Else temp(1) = V(0, Nj, 2)
            temp(2) = disc - temp(1)
        Else
            temp(1) = V(0, Nj, 1)
            temp(2) = V(0, Nj, 2)
        End If
    
        CNFDM = temp
    
    End If
    
End Function


