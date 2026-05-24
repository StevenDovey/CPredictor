Attribute VB_Name = "Module2"
'This module contains all functions for the DFNat growthmodel
'All scripting by Lars Wichmann Hansen, Forest Research.
'Functionality from various models and publications. See separate docs for this.

Option Explicit

Public Tinit As Double, Ninit As Double, ncal As Double, Tcal As Double

'This is a function for calculating the mean height from mean top height and stocking
'Input is mean top height and stocking, output is mean height.
Public Function calcMeanht_dfir(mth, N)
    Dim A As Double, B As Double
    A = 0.106
    B = -0.228
    calcMeanht_dfir = mth * (1 - A * (1 - Exp(B * (N - 100) / 100)))
End Function

'This is a function for calculating the mean top height from mean  eight and stocking
'Input is mean height and stocking, output is mean top height.
Public Function MH2MTH_dfir(MH, N)
    Dim A As Double, B As Double
    A = 0.106
    B = -0.228
    MH2MTH_dfir = (1 / MH * (1 - A * (1 - Exp(B * (N - 100) / 100)))) ^ -1
End Function

'This is a function for calculating the crown length of the unpruned element
'Input is mean top height and stocking, output is crown length
Public Function CLucalc(mth, sph)
If sph > 0 Then
    Dim k As Double, A As Double, B As Double
    k = 0.8429
    A = 6.9833
    B = 2028
    CLucalc = Excel.WorksheetFunction.Min(k * (mth - 0.1), A + B / sph)
Else
    CLucalc = 0
End If
End Function

'This function predicts basal area one month ahead
'Input is SBAP, stand basal area, stand crown length, start time, and end time
'Output is basal area at end time
Private Function BAcalc(SBAP, BAs, CL, Ts, Tf, t_last_thin)
    Dim B As Double, C As Double, D As Double, F As Double, g As Double, Crwn As Double, Aget As Double, _
        Comp As Double, dt_thin, ba_thin_ratio, shock_reduction, SBAP_t
    
    'reduce SBAP if we are close to a thinning (thinning shock)
    dt_thin = Ts - t_last_thin(1, 1)
    If (t_last_thin(1, 1) > 0 And dt_thin > 0 And dt_thin < 10) Then
        ba_thin_ratio = t_last_thin(1, 2)
        shock_reduction = 1 - (0.15 * Exp(-0.27 * dt_thin)) 'Thinning shock coefficient reduced in July 2010
'        shock_reduction = 1 - (0.31 * Exp(-0.27 * dt_thin))    'Thinning shock effect as implemented by Lars
        shock_reduction = WorksheetFunction.Max(shock_reduction, 0.1)
        SBAP_t = SBAP * shock_reduction
    Else
        SBAP_t = SBAP
    End If
    
    B = -0.0002059
    C = 3.0955
    D = 50
    F = -5.46
    g = 0.1217
    Crwn = (1 - Exp(B * CL))
    Aget = Excel.WorksheetFunction.Max(1, C + (1 - C) / D * Ts)
    Comp = 1 - Exp(F + g * BAs / SBAP_t) / (1 + Exp(F + g * BAs / SBAP_t))
    BAcalc = BAs + SBAP_t * Crwn * Aget * Comp * (Tf - Ts)
End Function

'This is a stocking/mortality model
'Input is stocking, mean top height, mean dbh, start time, finish time, Mort Adj
'Output is stocking at finish time
Private Function SPHcalc(Ns, Hs, Ds, Ts, Tf, MA)
    Dim A As Double, B As Double, C As Double, D As Double, X As Double, Y As Double
    A = 0.00007
    B = -20.43
    C = 1.517
    D = 0.3714
    X = Exp(B + C * (Log(Ns) + D * Log(Hs * Ds ^ 2)))
    Y = A + (1 - A) * X / (1 + X)
    
  'Overwrite mortality function when calibrating SI and SBAP or when running in offset mode
    If (Worksheets("500 Index").Range("W2") = 1 Or Worksheets("500 Index").Cells(17, 10) = 2) _
        And Tf <= Tcal + 0.01 And Tf > Tinit Then
           
        SPHcalc = ncal + (Ninit - ncal) * ((Tcal - Tf) / (Tcal - Tinit))  'with actual mortality
       
    'Standard mortality function without calibration
    
    ElseIf MA >= -2 And MA <= 2 Then
        SPHcalc = Ns * (1 - (1 + MA) * Y * (Tf - Ts))
        SPHcalc = WorksheetFunction.Min(SPHcalc, Ns)
               
    Else
        SPHcalc = Ns * (1 - Y * (Tf - Ts))
        
    End If

End Function

'This is Minas new function to predict mean top height from site index and time and latitude
Public Function MTHcalc(SI, T, Lat)
    Dim A As Double, B As Double, P As Double, q As Double
    A = -3.7082
    B = 0.3844
    P = 0.0338
    q = -0.00057
    MTHcalc = 0.25 + (SI - 0.25) * _
    ((1 - Exp(-Exp(A) * T)) / (1 - Exp(-Exp(A) * 40))) ^ _
    (1 / (B + (P + q * Lat) * SI))
End Function

'This function predicts a starting Basal Area, if nothing else is known
'Input SBAP, mean top height, stand age, and stocking
'Output is DBH
Public Function StartDBH(SBAP, mth, T, N)
    Dim A As Double, B As Double, C As Double, D As Double, F As Double, g As Double
    If T >= 10 Then
        A = 3.4
        B = 0.485
        C = 0.17
        D = 0
        F = 1.01
        g = -0.479
        StartDBH = A * (mth - 1.4) ^ B * (1 + C * (SBAP - 1.9)) * (1 + F * T * N ^ g)
    Else
        StartDBH = 0.001
    End If
End Function

'This is the thinning function
'Input is basal area, stocking before, stocking after and thinning coefficient
'Output is basal area after
Private Function Thincalc(BAs, Ns, Nf, A)
    If A > 0 Then
        Thincalc = BAs * (Nf / Ns) ^ A
    Else
        Thincalc = 10      'This is Lars original code, but Im not sure what it does
    End If
End Function

'This function calculates the dosheight for a given age
'Input: Stand age and an array with pruning information
'The pruning information is a 3x5 matrix with rows: Age at pruning, Pruned height, and stems pruned.
'Each column represent one pruning operation.
'Output is an array (1x5) of dos heights
Private Function Doshtcalc(age, Prune)
    Doshtcalc = 0.2
    For i = 1 To 5
        If Prune(1, i) <> 0 And Prune(1, i) <= age Then
            Doshtcalc = Prune(2, i) + 0.2
        End If
    Next i
End Function

'This function calculates the green crown length of the pruned elements at the time of pruning
'Its a seperate function because of array size limitations in the main function
Public Function GCL(T, Ages, CLu, MH, PrH)
    If T > 0 Then
        i = 1
        While Ages(i, 1) < T And Ages(i, 1) > 0
            i = i + 1
        Wend
        GCL = Excel.WorksheetFunction.Min(CLu(i, 1), MH(i, 1) - PrH)
    Else
        GCL = 0
    End If
End Function

'Converts BA to DBH
Public Function BA2DBH(BA, sph)
    If sph > 0 Then
        BA2DBH = (BA / sph / Excel.WorksheetFunction.Pi * 4) ^ 0.5 * 100
    Else
        BA2DBH = 0
    End If
End Function

'This function calculates the crown height by elements
Private Function CHcalc(Ts, Prune)
    Dim aCH(1, 6) As Double 'Array to keep crown height for each element
    Dim N As Long 'Counter variable
    For N = 0 To 5
        If Prune(1, N + 1) <= Ts Then  'If we're past the time of the pruning
            aCH(0, N + 1) = Prune(2, N + 1) 'Set the number of pruned stems in that element
        Else
            aCH(0, N + 1) = 0 'Otherwise set to zero
        End If
    Next N
    '''''
    CHcalc = aCH
End Function

'This function distributes stems to pruned elements
Private Function PrunedStems(Tf, Ts, sph, Prune, TT, pSPH, PT)
    Dim aSPH(1, 6) As Double 'Array to keep stocking for each element
    Dim N, Nominal, Alive As Double, i As Long, Nused As Double
    Dim ratio As Double, lastsph As Double

    'See if there is a pruning going on
    PT = prundicator(Tf, Ts, Prune)

    'If we're pruning or thinning in the period, then split into elements
    If PT >= 1 Or TT >= 1 Then
        For i = 1 To 5
            'First find out how many there are nominally
            N = 5 - i 'Start with the most extreme pruning
            If Prune(1, N + 1) <= Tf Then 'If we're past the time of the pruning
                Nused = Excel.WorksheetFunction.Sum(aSPH) 'Determine number of higher pruned stems
                If sph - Prune(3, N + 1) >= 0 Then 'If there're more trees in total than nominally pruned
                    Nominal = Prune(3, N + 1) - Nused 'Then all nominally pruned trees are pruned minus those pruned higher
                Else 'If there are more nominally pruned than there are trees in total
                    If sph - Nused > 0 Then 'if there are some trees but not all the nominally pruned
                        Nominal = sph - Nused 'Then count those left
                    Else 'Else if there are no more trees left
                        Nominal = 0 'then there are no trees in that element
                    End If
                End If
            Else 'If we're not past the time of pruning
                Nominal = 0 'Then there are no trees in that element
            End If
        
            'Determine number of live stems
            If pSPH(0, N + 1) > 0 Then 'If there were stems in previous step use that value
                Alive = pSPH(0, N + 1)
            Else 'If there were no stems in previous step,
                If Prune(1, N + 1) <= Ts - 1 / 12 Then 'then check if it is
                    Alive = 0 'because they are all dead or pruned higher,
                Else 'or
                    Alive = sph 'If there never was any then set value to maximum
                End If
            End If
            ''''
        
            aSPH(0, N + 1) = Excel.WorksheetFunction.Min(Nominal, Alive)
        Next i
    
        If sph - Excel.WorksheetFunction.Sum(aSPH) <= 0 Then
            aSPH(0, 0) = 0 'There are no unpruned stems
        Else
            aSPH(0, 0) = sph - Excel.WorksheetFunction.Sum(aSPH) 'The rest is unpruned
        End If
    
    Else 'if we're not pruning or thinning then distribute death
        lastsph = Excel.WorksheetFunction.Sum(pSPH)
        If lastsph <> 0 Then
            ratio = sph / lastsph 'Ratio of mortality
            For i = 0 To 5
                aSPH(0, i) = pSPH(0, i) * ratio
            Next i
        End If
    End If
    '''''

    PrunedStems = aSPH 'Return the array of SPH's
End Function

'This function converts crown heights to crown lengths and multiplies by SPH
'And it returns the total stand crown length
Private Function CLcalc(MH, mth, sph, pSPH, pCH)
    Dim CL As Double, CLp As Double, CLu As Double, i As Long
    CLu = CLucalc(mth, sph) 'Calculate crown length for unpruned trees
    CL = 0
    For i = 0 To 5
        CLp = MH - pCH(0, i) 'Calculate crown length for a pruned element
        CL = CL + pSPH(0, i) * Excel.WorksheetFunction.Min(CLu, CLp) 'Pick the one that is the shortest
    Next i
    CLcalc = CL
End Function

'This function calculates how the ba increment is distributed to the pruned elements
Private Function PrunedGrowth(MH, mth, sph, pSPH, pCH, CLtotal, Prune, dBA)
    Dim Out(1, 6) As Double 'Array for individual crown lengths
    Dim CL, CLp As Double, CLu, i As Long
    CLu = CLucalc(mth, sph) 'Calculate crown length for unpruned trees
    CL = 0
    For i = 0 To 5
        If (pCH(0, i) > 0) Or i = 0 Then
            CLp = MH - pCH(0, i) 'Calculate crown length for a pruned element
            CL = Excel.WorksheetFunction.Min(CLu, CLp) 'Pick the one that is the shortest
            Out(0, i) = CL * pSPH(0, i) / CLtotal * dBA
        Else
            Out(0, i) = 0
        End If
    Next i
    PrunedGrowth = Out
End Function

'This function decides if there is a pruning
Private Function prundicator(Tf, Ts, Prune) As Double
    Dim i As Long
    prundicator = 0
    For i = 1 To 5
        If Prune(1, i) > Ts And Prune(1, i) <= Tf Then
            prundicator = Prune(1, i)
        End If
    Next i
End Function

'This function decides if there is user input
Private Function Uindicator(Ts, Uinput)
    Dim Out(3, 1), i As Long, Tf
    Out(0, 0) = 0
    For i = 1 To 5
        If Uinput(1, i) > Ts And Uinput(1, i) <= Tf Then

            Out(0, 0) = 1
            Out(1, 0) = Uinput(2, i)
            Out(2, 0) = Uinput(3, i)
            
        End If
    Next i
    Uindicator = Out
End Function

'This function decides if there is a thinning
Private Function Thindicator(Tf, Ts, Thin)
    Dim Out(3, 1), i As Long
    For i = 1 To 5
        If Thin(1, i) > Ts And Thin(1, i) <= Tf Then
            If Thin(2, i) > 0 And Thin(3, i) > 0 Then
                Out(0, 0) = 1
                Out(1, 0) = Thin(2, i)
                Out(2, 0) = Thin(3, i)
                Out(3, 0) = Thin(1, i)
            End If
        End If
    Next i
    Thindicator = Out
End Function


'This function keeps track of the cumulative basal area by pruned elements
Function pBAcalc(pBA, pGrowth, pSPH, pSPHb4, BA, sph, PT, TT, Ut)
    Dim Out(0, 5) As Double
    Dim GrowF(0, 5) As Double

    Dim alpha As Double, i As Long, BAratio As Double

    alpha = 0.705 'Selection ratio

    'If we're simply distributing the increment to elements
    For i = 0 To 5
        Out(0, i) = pBA(0, i) + pGrowth(0, i)
        GrowF(0, i) = pBA(0, i) + pGrowth(0, i) 'Grow the stand foreward before distributing BA
    Next i

    'If we're pruning or thinning then we need to redistribute already cumulated BA
    If (PT >= 1 Or TT >= 1) Then
    
        For i = 1 To 5
        
            'if the number of stems in the element increases markedly
            If pSPH(0, i) - pSPHb4(0, i) > 4 Then
                        
                'if there is enough to select from in the element below, then take'em
                If pSPH(0, i) <= pSPHb4(0, i - 1) Then
                    Out(0, i) = GrowF(0, i - 1) * (pSPH(0, i) / pSPHb4(0, i - 1)) ^ alpha
                    Out(0, i - 1) = Out(0, i - 1) - Out(0, i)
                Else 'Otherwise take it all
                    Out(0, i) = pBA(0, i - 1) + pGrowth(0, i - 1)
                    Out(0, i - 1) = 0
                End If
            End If
            '''''
        
            'if the number of stems in the element decreases markedly
            If pSPH(0, i) = 0 Then 'if the stocking goes to zero,
                Out(0, i) = 0 'Then there is no basal area
                'If Tt >= 1 Then 'If its a thinning, and basal area is disappearing
                '    sumofpruned = Excel.WorksheetFunction.Sum(Out)
                '    If BA > sumofpruned Then
                '        ratio = BA / sumofpruned
                '        For j = 1 To 5
                '            Out(0, j) = Out(0, j) * ratio
                '        Next j
                '    End If
                'End If
            End If
            ''''
        
        Next i
    End If

    'Take care of the unpruned element if its a thinning
    If TT >= 1 And Abs(pSPH(0, 0) - pSPHb4(0, 0)) >= 1 And pSPHb4(0, 0) > 0 Then 'if things are changing
        Out(0, 0) = pBA(0, 0) * (pSPH(0, 0) / pSPHb4(0, 0)) ^ alpha 'BA of the selected
    
        'We can also thin pruned elements
        For i = 1 To 5
            If Abs(pSPH(0, i) - pSPHb4(0, i)) >= 1 And pSPHb4(0, i) > 0 Then
                Out(0, i) = pBA(0, 0) * (pSPH(0, i) / pSPHb4(0, i)) ^ alpha 'BA of the selected
            End If
        Next i
        '''''
    
    End If
    ''''''

    'Take care of redistributing the basal area if there is userinput
    If Ut >= 1 And pSPHb4(0, 0) > 0 Then 'if there are stems in the element
        'Redistribute the basal area relative to how much there was before the user-input
        BAratio = BA / Excel.WorksheetFunction.Sum(pBA)
        For i = 0 To 5
            Out(0, i) = pBA(0, i) * BAratio 'BA of the selected
        Next i
    End If
    ''''

    pBAcalc = Out
End Function


'This function combines the individual components to predict a stands life
Public Function Grow(SBAP As Double, SI As Double, SPHin As Double, Rotation As Double, Thin, Prune, Lat, MA)
    Dim Out(350, 20) As Double 'Declare output matrix for standlevel data
    Dim Dummy(1, 6) 'Empty dummy matrix
    Dim SPHs As Variant 'Declare variable to keep track of stocking
    Dim PT, TT, Ut, Ts, dBA, BAb4 As Double 'Variabel declaration
    Dim ReportIndicator As Long 'Thinning time and report indicator
    Dim t_last_thin(2, 2) As Double 'A variable to store info about last thinning

    Dim TsU As Double, TsD As Double, TsM As Double, MTH_TsM_Old As Double, MTH_TsM As Double, N As Double, _
        MH_TsM As Double, MTH_start As Double, i As Long, TsY As Long, MTHs As Double, MHs As Double, _
        DBHs As Double, BAs As Double, Tf As Double, CLs As Double, _
        iC As Long, MTHprev As Double, _
        SPHf As Double, A As Double, thin_exact_age As Double, TAge As Double, TSPH As Double, TBA As Double, _
        Kurt As Double, Msg, Style, TITLE, Response, pBA() As Double, pSPH() As Double, pSPHb4() As Double, Temp, _
        pCH, pGrowth

    '''''

    'Check if the stand parameters are within the range required
    'Note from 2009: Changed SBAP lower limit from 0.5 to 0.2 for PSP batch run
    If SBAP >= 0.2 And SBAP < 5 And SI > 15 And SI < 50 And Rotation > 20 And Rotation <= 90 Then

    'Find start time
    TsU = 14
    TsD = 0
    TsM = (TsU + TsD) / 2
    MTH_TsM_Old = 0
    MTH_TsM = MTHcalc(SI, TsM, Lat)
    MH_TsM = calcMeanht_dfir(MTH_TsM, N)
    MTH_start = 4
    While (Abs(MTH_TsM_Old - MTH_TsM) > 0.05 And (MH_TsM >= MTH_start + 0.05 Or MH_TsM <= MTH_start))
        If (MH_TsM < MTH_start) Then
            TsD = TsM
        Else
            TsU = TsM
        End If
        TsM = (TsU + TsD) / 2
        MTH_TsM_Old = MTH_TsM
        MTH_TsM = MTHcalc(SI, TsM, Lat)
        MH_TsM = calcMeanht_dfir(MTH_TsM, SPHin)
    Wend
    If TsM >= 10 Then
        TsM = 10
        i = 1
    Else
        TsY = WorksheetFunction.Round(TsM, 0)
        TsM = WorksheetFunction.Round(TsM * 12, 0) / 12 'Make sure we start the nearest month
        i = Abs((TsM - TsY) * 12) + 1
    End If
    '''''

    ''Here we go
    Ts = TsM 'Set start time
    TT = 1 'Assume that the start is like a thinning
    PT = 1 'Assume that the start is also a pruning
    MTHs = MTHcalc(SI, Ts, Lat) 'Calculate starting mean top height
    SPHs = SPHin 'Set stocking to the value inputtet
    MHs = calcMeanht_dfir(MTHs, SPHs) 'Calculate mean height given the mth, stocking
    DBHs = StartDBH(SBAP, MTHs, Ts, SPHs) 'Calculate start DBH
    BAs = DBHs ^ 2 * Excel.WorksheetFunction.Pi / 40000 * SPHin 'Caculate start BA
    pSPH = PrunedStems(Tf, Ts, SPHs, Prune, TT, Dummy, PT) 'Calculate number of stems in each element
    pCH = CHcalc(Ts, Prune) 'Calculate crown height for all pruned elements
    CLs = CLcalc(MHs, MTHs, SPHs, pSPH, pCH) 'Calculate the total crown length of the stand
    pGrowth = PrunedGrowth(MHs, MTHs, SPHs, pSPH, pCH, CLs, Prune, dBA) 'Calculate growth proportion to elements
    pBA = pBAcalc(Dummy, Dummy, pSPH, Dummy, BAs, SPHs, 1, 1, 1) 'Calculate cummulated basal area, empty run
    pBA(0, 0) = BAs + pBA(0, 0) 'Input the starting basal area
    ''That was the first step ... now keep on going

    iC = 0 'Counter to keep track of which row in the output matrix we've reached
    'i = 1

    While Ts <= Rotation 'Iterate
        Tf = Ts + 1 / 12 'Next time
    
        'Remember values from previous step
        BAb4 = BAs 'Remeber last BA
        MTHprev = MTHs 'Remember last MTH
        pSPHb4 = pSPH 'Remember last SPH by element
        '''''
    
        'Determine if there is a thinning in this period
        Temp = Thindicator(Tf, Ts, Thin) 'Find out if there is a thinning
        TT = Temp(0, 0) 'Remember if it is now
        SPHf = Temp(1, 0) 'Remember stocking after thinning
        A = Temp(2, 0) 'Remember thinning coefficient
        thin_exact_age = Temp(3, 0) 'Remember thin age
        If TT = 1 Then Tf = thin_exact_age 'If thinning, set end of step to thin age
        '''''''
    
        'Calculate new mean top height
        MTHs = MTHcalc(SI, Tf, Lat) 'Mean top height
        ''''''
    
        'Determine if there is a pruning in this period
        PT = prundicator(Tf, Ts, Prune) 'Determine if there is a pruning NOW
        '''''''
    
        'Calculate what would happen if no thinning or user-input is taking place
        SPHs = SPHcalc(SPHs, MTHs, DBHs, Ts, Tf, MA) 'Calculate stocking
        BAs = BAcalc(SBAP, BAs, CLs, Ts, Tf, t_last_thin) 'Basal area
        dBA = BAs - BAb4 'Calculate BA increment
        DBHs = BA2DBH(BAs, SPHs) 'Calculate quadratic mean DBH
        MHs = calcMeanht_dfir(MTHs, SPHs) 'Calculate mean height
        pCH = CHcalc(Tf, Prune) 'Calculate crown height for all pruned elements
        pSPH = PrunedStems(Tf, Ts, SPHs, Prune, TT, pSPH, PT) 'Calculate number of stems in each element
        CLs = CLcalc(MHs, MTHs, SPHs, pSPH, pCH) 'Calculate the total crown length of the stand
        '''''''
           
        'if there is a thinning NOW, then apply it
        If TT > 0 Then
    
            Out(iC, 0) = Tf
            Out(iC, 1) = DBHs
            Out(iC, 2) = MTHs
            Out(iC, 3) = MHs
            Out(iC, 4) = SPHs
            Out(iC, 5) = BAs
            Out(iC, 6) = CLs
            Out(iC, 7) = pBA(0, 1)
            Out(iC, 8) = pBA(0, 2)
            Out(iC, 9) = pBA(0, 3)
            Out(iC, 10) = pBA(0, 4)
            Out(iC, 11) = pBA(0, 5)
            Out(iC, 12) = TAge
            Out(iC, 13) = TSPH
            Out(iC, 14) = TBA
            Out(iC, 15) = pSPH(0, 1)
            Out(iC, 16) = pSPH(0, 2)
            Out(iC, 17) = pSPH(0, 3)
            Out(iC, 18) = pSPH(0, 4)
            Out(iC, 19) = pSPH(0, 5)
            iC = iC + 1 'Increase report counter
            Tf = Tf + 1 / 12 'Increase time by one increment before reporting post-thin values
            i = i + 1 'Make sure the time reported is whole years
        
            t_last_thin(1, 2) = BAs 'Remember the BA before thinning
            BAs = Thincalc(BAs, SPHs, SPHf, A) 'Calculate after thinning basal area
            t_last_thin(1, 2) = t_last_thin(1, 2) / BAs 'Remember the BA before thinning
            TAge = Ts 'Remember thinning age
            TBA = BAb4 - BAs 'Calculate thinned basal area
            TSPH = SPHs - SPHf 'Calculate thinned stems
            SPHs = SPHf 'Set new stocking
            DBHs = BA2DBH(BAs, SPHs) 'Calculate quadratic mean DBH
            pCH = CHcalc(Tf, Prune) 'Calculate crown height for all pruned elements
            pSPH = PrunedStems(Tf, Ts, SPHs, Prune, TT, pSPH, PT) 'Calculate number of stems in each element
            CLs = CLcalc(MHs, MTHs, SPHs, pSPH, pCH) 'Calculate the total crown length of the stand
            t_last_thin(1, 1) = Ts 'Remeber the last thinning age
        End If
        '''''''
        
'        'If there is user input then apply it
'        Temp = Uindicator(Tf, Uinput) 'Find out if it is NOW
'        Ut = Temp(0, 0) 'Save as indicator
'        If Ut = 1 Then 'if there is user input then use those values
'            If (Temp(1, 0) > 0) Then 'Only do so if there is a value
'                SPHs = Temp(1, 0) 'Determine SPH as measured by the user
'            End If
'            If (Temp(2, 0) > 0) Then 'Only if there is a value
'                BAs = Temp(2, 0) 'Determine basal area as measured by user
'            End If
'            DBHs = BA2DBH(BAs, SPHs) 'Calculate quadratic mean DBH
'            pCH = CHcalc(Tf, Prune) 'Calculate crown height for all pruned elements
'            pSPH = PrunedStems(Tf, Ts, SPHs, Prune, TT, pSPH, PT) 'Calculate number of stems in each element
'            CLs = CLcalc(MHs, MTHs, SPHs, pSPH, pCH) 'Calculate the total crown length of the stand
'        End If
        ''''''''''
    
        'Keep track of basal area by element
        pGrowth = PrunedGrowth(MHs, MTHs, SPHs, pSPH, pCH, CLs, Prune, dBA) 'Calculate growth proportion to elements
        pBA = pBAcalc(pBA, pGrowth, pSPH, pSPHb4, BAs, SPHs, PT, TT, Ut) 'Calculate cumulative basal area
        Kurt = Kurt + 1
        ''''''
            
        'Determine if we're approaching rotation or are close to a calibration point
        ReportIndicator = 0
        If Tf < 20 Or TBA > 0 Then
            If i Mod 1 = 0 Then
                ReportIndicator = 1
            End If
        Else
            If Abs(Tf - Round(Tf, 0)) < 1 / 13 Then
                ReportIndicator = 1
            End If
        End If
        i = i + 1
        ''''''''
        Ts = Tf 'Next step in time
    
        'Collect data for output, only when silviculture takes place, or when we calibrate
        If ReportIndicator = 1 Then
            Out(iC, 0) = Tf
            Out(iC, 1) = DBHs
            Out(iC, 2) = MTHs
            Out(iC, 3) = MHs
            Out(iC, 4) = SPHs
            Out(iC, 5) = BAs
            Out(iC, 6) = CLs
            Out(iC, 7) = pBA(0, 1)
            Out(iC, 8) = pBA(0, 2)
            Out(iC, 9) = pBA(0, 3)
            Out(iC, 10) = pBA(0, 4)
            Out(iC, 11) = pBA(0, 5)
            Out(iC, 12) = TAge
            Out(iC, 13) = TSPH
            Out(iC, 14) = TBA
            Out(iC, 15) = pSPH(0, 1)
            Out(iC, 16) = pSPH(0, 2)
            Out(iC, 17) = pSPH(0, 3)
            Out(iC, 18) = pSPH(0, 4)
            Out(iC, 19) = pSPH(0, 5)
               
            TBA = 0
            TSPH = 0
            TAge = 0
            iC = iC + 1
        End If
        '''''
    
        If iC > 350 Then
            Msg = "NB! Stand outside the possible range of the calculator!"    ' Define message.
            Style = vbCritical    ' Define buttons.
            TITLE = "Douglas-fir calculator error!"    ' Define title.
            Response = MsgBox(Msg, Style, TITLE)
            Ts = Rotation + 1
        End If
    
    Wend
    End If 'if checking on Site index and SBAP

    Grow = Out
End Function


'The following are solver routines for Douglas-fir.

'Combined SI solver and SBAP solver
Sub CombineSolver()
    Dim X As Long, thinage
    
    Worksheets("500 Index").Range("W2") = 1
    ncal = Worksheets("500 Index").Cells(9, 10)
    Tcal = Worksheets("500 Index").Cells(5, 10)
    Tinit = 0
    Ninit = Worksheets("500 Index").Cells(2, 6)
    
    Call SIsolver
    
    
    For X = 3 To 7
    
        thinage = Worksheets("500 Index").Cells(14, X)
        If thinage <= Tcal And thinage <> 0 Then
            Tinit = thinage
            Ninit = Worksheets("500 Index").Cells(15, X)
        End If
        
    Next X
        
    Call SBAPsolver
    
    Worksheets("500 Index").Range("W2") = 0
    
End Sub

'Estimate the 500 Index from one measurement
Sub SBAPsolver()
    Dim Msg, Style, TITLE, Response
    Dim dI500 As Double, mI500 As Double, uI500 As Double, dBA As Double, mba As Double, uba As Double, targetba As Double

    If Worksheets("500 Index").Cells(5, 10) > 0 And Worksheets("500 Index").Cells(6, 10) > 0 Then
'        Application.Calculation = xlCalculationManual
        dI500 = 1
        uI500 = 40
'        Worksheets("500 Index").Range("L2") = Data(1, 1)
        While Abs(dI500 - uI500) > 0.02
            mI500 = (dI500 + uI500) / 2
            Worksheets("500 Index").Cells(2, 2) = mI500
            Worksheets("500 Index").Cells(18, 2).Calculate
            Worksheets("500 Index").Calculate
            If Worksheets("500 Index").Cells(14, 14) < Worksheets("500 Index").Cells(6, 10) Then
                dI500 = mI500
            Else
                uI500 = mI500
            End If
        Wend
        Worksheets("500 Index").Cells(2, 2) = mI500
    
        'Previous version v3.1 diplayed the predicted SPH from worksheet 500 Index
        'From Feb 2010, it's changed to actual stocking from PSP
        'Worksheets("User interface").Range("G38") = Worksheets("500 Index").Range("O2")
    
'        Application.Calculation = xlCalculationAutomatic
    Else
        Msg = "Missing values, please input stand age and basal area!"    ' Define message.
        Style = vbCritical    ' Define buttons.
        TITLE = "SBAP calibration - missing values"    ' Define title.
        Response = MsgBox(Msg, Style, TITLE)
    End If
End Sub

'Estimate the site index from one measurement
Sub SIsolver()
    Dim Msg, Style, TITLE, Response, TT, th, Lat, uSI, dSI, uH, dH

'    If (Range("J8") > 0 And Not WorksheetFunction.IsNumber(Range("J7"))) Then
'        Call SBAPsolver
'        Range("J7") = MH2MTH_dfir(Range("J6"), Worksheets("500 Index").Range("W5"))
'    End If

    If Worksheets("500 Index").Cells(5, 10) > 0 And Worksheets("500 Index").Cells(7, 10) > 0 Then
'        Application.Calculation = xlCalculationManual
        TT = Worksheets("500 Index").Cells(5, 10)
        th = Worksheets("500 Index").Cells(7, 10)
        Lat = Worksheets("500 Index").Cells(2, 4)
        If Lat = 0 Then Lat = 41    'Default latitude is 41
            uSI = 50
            dSI = 10
            uH = MTHcalc(uSI, TT, Lat)
            dH = MTHcalc(dSI, TT, Lat)
            While uH - dH > 0.1
                If (Abs(uH - th) > Abs(dH - th)) Then
                    uSI = uSI - (uSI - dSI) / 4
                Else
                dSI = dSI + (uSI - dSI) / 4
                End If
                uH = MTHcalc(uSI, TT, Lat)
                dH = MTHcalc(dSI, TT, Lat)
            Wend
        Worksheets("500 Index").Cells(3, 2) = (uSI + dSI) / 2
        Worksheets("500 Index").Calculate
'        Application.Calculation = xlCalculationAutomatic
    Else
        Msg = "Missing values, please input stand age and target mean top height!"    ' Define message.
        Style = vbCritical    ' Define buttons.
        TITLE = "Site index calibration - missing values"    ' Define title.
        Response = MsgBox(Msg, Style, TITLE)
    End If
End Sub

'This function calculates SBAP from 500-index
Function Five2SBAP(FiveIndex, SI)
    Five2SBAP = (FiveIndex / (0.0971 * SI ^ 1.344)) ^ 1.03
End Function

' Some more Douglas-fir functions

'This function calculates BA from DBH
Function DBH2BA(dbh)
    DBH2BA = dbh ^ 2 * 3.14159265359873 / 40000
End Function

'This is newly fitted Beekhuiz stand volume function
Public Function Volume(BA, H, sph)
    Dim A, B, D As Double
    If BA > 0 Then
        A = 0.928
        B = 0.3208
        Volume = BA * (A + B * H)
    Else
        Volume = 0
    End If
End Function

Sub earlyield_dfir()
'This subroutine corrects early volume predictions in D-fir yield table
'
'   An earlier version of this subroutine was included in FCP 4.10
'   This version was implemented in FCP4.11 in May 2013
'
'   This routine assumes Beets volume at planting is 0.0000064 m3. Secondly, early volume growth of individual trees is assumed
'   proportional to Age^3 (based on early D-fir PSP data). Although DFIRNAT predicts DBH from about age 5 years, in practice
'   DBH predictions are suspect for several years from this starting year. Therefore, the module uses predicted volume 2 years
'   after the first DBH prediction, and interpolates back to year 0 using a power function
'
    Dim initialvol As Double, age As Double, treevolinc As Double, i As Long, _
        dbh As Double, prevdbh As Double, j As Long, k As Double, T As Double
    Worksheets("500 Index").Activate
    Range(Cells(23, 39), Cells(43, 39)).ClearContents
    initialvol = 0.0000064  'Volume of seedling at planting (m3) - Beets
    dbh = 0
    For i = 23 To 38
        prevdbh = dbh
        dbh = Cells(i, 35)
        If prevdbh = 0 And dbh > 0 Then
            treevolinc = Cells(i + 3, 38) / Cells(i + 3, 31) - initialvol
            age = Cells(i + 3, 30)
            k = treevolinc / (age ^ 3)  'parameter k ensures that volume is correct when T=age
            For j = i + 3 - 1 To 23 Step -1
                T = Cells(j, 30)
                If T < 0 Then T = 0
                Cells(j, 39) = (initialvol + k * T ^ 3) * Cells(j, 31)
            Next j
        End If
    Next i
End Sub

'This function looks up the report values
Function ReportLookup_dfir(T, age, X)
    Dim i As Long, DT As Double
    If T > 0 Then
        i = 1
        While Excel.WorksheetFunction.Round(age(i), 2) <= Excel.WorksheetFunction.Round(T, 2) And age(i) > 0
            i = i + 1
        Wend
        If (age(i) > T) And i <> 1 Then
            DT = T - age(i - 1, 1)
            ReportLookup_dfir = X(i - 1) + (X(i) - X(i - 1)) * DT
        ElseIf i <> 1 Then
            ReportLookup_dfir = X(i - 1)
        Else
            ReportLookup_dfir = 0
        End If
    End If
End Function


