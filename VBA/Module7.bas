Attribute VB_Name = "Module7"
'**********************************************************************************************************************************
'
'   Forest Carbon Predictor
'
'   This module contains code for running the C_Change & 300Index models in interactive mode from the
'       worksheet "C Change". It also contains code for performing 300 Index and C_Change batch runs.
'
'   History:
'
'   1. FCP 1.0 was the first inmplementation of the 300 Index, C_Change and radiata wood density models.
'       The wood density and 300 Index models were coded in VBA and C_Change was called as an executable
'       file complied from the original FORTRAN implementation.
'
'   2. FCP 2.0 was developed in January 2009. This version was functionally almost identical to FCP 2.1
'       but it contained a VBA implementation of the C_Change model.
'
'   3. FCP 2.2 contained several updates including an adjustment to the wood density model based on
'       validation results. Several minor bugs were also corrected.
'
'   4. FCP 2.3.1 contains the following modification:
'           1.  New temperature dependent (Q10) decay rate functions for estimating losses of carbon
'                from the above-ground dead wood and litter pools.
'           2.  New breast height pith-to-bark density functions for estimating growth ring density.
'           3.  Adjustment functions to improve accuracy of carbon stock estimates by pool when applying
'                the C_Change model to Douglas-fir (when based on the 300 Index growth model, species code DFIR).
'           4.  A new stand tending module that predicts silvicultural history information for a plot
'                when stand data are incomplete.
'           5.  Includes a new wood density model described in the report for MFE: Kimberley and Beets
'                 April 2010 "A new model for predicting breast height wood density by ring number in
'                 radiata pine". Also, the wood density function has been modified by setting a minimum
'                 ring width limit of 1.5 mm.
'
'   5. FCP3Beta. In April 2011, new decay functions for dead wood were included in C_Change.
'
'   6. FCP3 In June 2011, modifications to the 300Index model which include corrections for old ages and
'       high stockings were made.
'
'   7. FCP4.01 In March 2012, the 500 Index model was ported from the D-fir Calculator and the MAF Douglas-fir
'       wood density model added to the system.
'
'   8. FCP4.02-4.04 In March-April 2012, minor bugs and updates fixed
'
'   9. FCP4.05-4.08 In April-May 2012,
'       Implementation of Version 1.07 of the 300 Index Model.
'       Implementation of three operating modes:
'           1. Standard mode: Use calibration meaurement if available to estimate productivity indices. Otherwise
'               use specified productivity indices.
'           2. Offset mode: Uses specified productivity indices but apply offsets to ensure that MTH and BA are
'               consistent with calibration measurement.
'           3. Index mode: Ignore calibration measurement and use specified productivity indices.
'       Allow user to specify pruning selection coefficients when running 300 Index in its stand-alone sheet
'
'   10. FCP4.09 In May 2012, Implementation of Version 1.08 of the 300 Index Model.
'
'   11. FCP4.10 In September 2012, Carbon calculations for PMEN were updated.
'            1. The coefficients for predicting density of the outer ring at reference age 30 years is now:
'                   Density=33.47796+27.88*LN(Age)+22.76*MAT+3.31*C/(N-0.014)
'                   (cells V14-V16 in 500 Index worksheet)
'            2. Douglas-fir carbon adjustments for stem, crown, litter and dead wood for use with 500 Index model
'                   and density function (species code PMEN) added to Module 7
'
'   12. FCP4.11 In May 2013, the subroutine earlyield_dfir in Module 2 which predicts early PMEN volume was updated.
'
'   13. FCP4.12 In June 2014,
'           1. Implements Version 1.08 of the 300 Index model. This has a change to the subroutine Thinning in
'               Module 1 which corrects an error in the calculation of the thinning age shift that occured
'               when a late thining was applied with a non-zero drift factor. Also widens the allowed limits for
'               latitude.
'           2. Fixes a bug in Module 7 that meant that the needle retention score in the Plots worksheet was ignored
'               when running the FCP in batch mode. Also, leaves details of the last plot in the C Change
'               worksheet at the end of a batch run.
'
'   14. FCP4.14 In January 2016, initial implementation of NuBalm
'
'   15. FCP4.15 In June 2016
'          1. Implementation of NuBalm model. This is an addition to FCP that predicts Nitrogen and Phosphorus contents
'               in each component of the tree. To do this, code was added to the C_Change programme, and 2 worksheets
'               were added giving N and P content predictions. The code is run automatically when C_Change runs. There
'               are two inputs in the C Change worksheet for adjusting the standard N and P Puruki models.
'          2. Removal of crown and forest floor added to C_Change. This affects predictions of N and P content.
'               Percentage crown removal and forest floor removal added to C Change worksheet.
'
'   16. FCP4.16 In September 2016
'          1. If XTRACC = -1, FCP applies the LTSP1 double slash treatment to Rotation 2
'          2. N & P are retained during the decay of attached branches
'
'   17. FCP5.1 In February 2017
'          1. New Zealand component-specific carbon fractions are applied in place of the generic 0.5 values used previously.
'               The C fraction for stem bark is a function of stand age.
'          2. Dfir carbon adjustment approach has been revised according to a paper prepared on this subject
'                   - stem wood now has no adjustment but a separate adjustment is applied to stem bark
'                   - adjustments for crown, litter and dead wood still apply (but slightly adjusted)
'          3. Initial dead wood and litter pools from harvest residues from the previous stand for C, N & P
'               can be specified as initial conditions. Otherwise the current default values of zero apply.
'
'   18. FCP5.2 In June 2019
'          1. New NuBalm functions for estimating component N and P concentrations from Soil C, N, Organic P were implemented.
'                This required changes to inputs in "C Change" and "Plots" worksheets, and changes to the Nconc & Pconc functions
'                   in C-Change
'          2. Newly published Zealand specific carbon fractions for radiata pine were implemented for softwood species while IPCC
'                   defaults are used for eucalypts
'          3. Newly published C-Change carbon adjustments were implemented for stem, crown, litter and dead wood for Douglas-fir,
'                   lusitanica and eucalypts
'
'**********************************************************************************************************************************


Option Explicit
    
Dim rotlth As Long, NR As Double, Initialsph As Double, sheathdens As Boolean, nodisturb As Long, _
    topheight As Boolean, prodthinpct As Double, clearfellpct As Double, outyieldrow As Long, _
    crownremovalpct As Double, forestfloorremovalpct As Double, _
    outdistrow As Long, inthinrow As Long, inprunerow As Long, inyieldrow As Long, _
    thinage As Long, thinsph As Double, prodthin As Boolean, pruneage As Long, _
    prunesph As Long, pruneht As Double, age As Long, stocking As Double, sphaftthin As Double, _
    Height As Double, Volume  As Double, Volaftthin As Double, voldead As Double, density As Double, _
    GrossVol As Double, inrow As Long, AGL As Double, BGL As Double, DWL As Double, FL As Double, _
    Total As Double, i As Long, rotlth2 As Long, rotlth1 As Long, _
    height2 As Double, prev_vol As Double, prev_dens_wt As Double, dens_wt As Double, dens_sh As Double, _
    CSHRUB As Double, prht(10) As Double, prsph(10) As Double, prelements As Long, totalprht As Double, _
    avprht As Double, totalprsph As Double, lastpruneage As Long, species As String, RunNuBalM As Boolean, _
    RunCChange As Boolean
    
Sub Run_C_Change()
'This subroutine predicts carbon using C Change from the yield table in the C Change worksheet
    Dim stem As Double, crown As Double, Stembark_adj As Double, Stemwood_adj As Double, Deadwood_adj As Double, _
        Crown_adj As Double, Litter_adj, CFraction_wood As Double, CFraction_roots As Double, CFraction_needles As Double, _
        CFraction_branches As Double, CFraction_DWL As Double, CFraction_bark As Double, UnAdj_AGL As Double, _
        Adj_ratio As Double, Initial_DryMat(5) As Double, Initial_N(5) As Double, Initial_P(5) As Double

    'Generate C_Change input file
    
    Application.ScreenUpdating = False
    Worksheets("C Change").Activate
    Range(Cells(6, 16), Cells(200, 28)).ClearContents
    species = Cells(3, 3)
    rotlth1 = Cells(5, 3)
    rotlth2 = Cells(6, 3)
    If rotlth2 > rotlth1 Then
        rotlth = rotlth2
    Else
        rotlth = rotlth1
    End If
    Initialsph = Cells(4, 3)
    If LCase(Cells(58, 3)) = "t" Then
        sheathdens = False
    Else
        sheathdens = True
    End If
    
    If LCase(Cells(59, 3)) = "m" Then
        topheight = False
    Else
        topheight = True
    End If
    
    prodthinpct = Cells(55, 3)
    clearfellpct = Cells(54, 3)
    crownremovalpct = Cells(56, 3)
    forestfloorremovalpct = Cells(57, 3)
    NR = Cells(38, 3) / 3
    'Default needle retention
    If NR = 0 Then NR = 2.1 / 3
    
    outyieldrow = 18
    outdistrow = 6
    inthinrow = 11
    inprunerow = 20
    inyieldrow = 6
    
    For i = 1 To 5
        Initial_DryMat(i) = Cells(63 + i, 3)
        Initial_N(i) = Cells(63 + i, 4)
        Initial_P(i) = Cells(63 + i, 5)
    Next i
    
    Worksheets("INPUT").Activate
    Range(Cells(1, 1), Cells(200000, 10)).ClearContents
    Cells(1, 1) = 2
    Cells(2, 1) = "Rotation 1"
    Cells(3, 1) = 0
    Cells(3, 2) = 0
    Cells(3, 3) = 1
    Cells(3, 4) = 1 'YESN=1 provides nitrogen modelling
    Cells(3, 5) = 0
    Cells(3, 6) = 1
    Cells(3, 7) = 2
    Cells(3, 8) = 0
    Cells(4, 1) = rotlth1
    Cells(4, 2) = 6
    Cells(4, 3) = 3
    Cells(4, 4) = 0
    Cells(4, 5) = nodisturb
        
    For i = 1 To 10
        Cells(4 + i, 1) = 0
        Cells(4 + i, 2) = 0
        Cells(4 + i, 3) = -1
        Cells(4 + i, 4) = -1
        Cells(4 + i, 5) = 0
        Cells(4 + i, 6) = 0
        Cells(4 + i, 7) = 0
    Next i
    
    Cells(15, 1) = -0.001825
    Cells(15, 2) = 1.77250004
    Cells(15, 3) = 30   'Dummy value for SI
    Cells(15, 4) = 0.05
    Cells(15, 5) = NR
    Cells(15, 6) = 7.2
    Cells(15, 7) = 1
    Cells(15, 8) = 200
    Cells(15, 9) = 6
    Cells(15, 10) = 0.3
    
    Cells(16, 1) = 0.0043
    Cells(16, 2) = 0
    Cells(16, 3) = 0
    Cells(16, 4) = 0
    Cells(16, 5) = 0
    Cells(16, 6) = 0.0029
    Cells(16, 7) = 0
    Cells(16, 8) = 0.0011
    Cells(16, 9) = 0
    Cells(16, 10) = 0
    
    For i = 1 To 5
        Cells(16, 10 + i) = Initial_DryMat(i)
        Cells(16, 15 + i) = Initial_N(i)
        Cells(16, 20 + i) = Initial_P(i)
    Next i
    
    nodisturb = 2
    Cells(5, 2) = Initialsph
    prev_vol = 0
    prev_dens_wt = 0
    prelements = 0
    
    Do
        Worksheets("C Change").Activate
        thinage = Cells(inthinrow, 2)
        thinsph = Cells(inthinrow, 3)
        If LCase(Cells(inthinrow, 1)) = "p" Then
            prodthin = True
        Else
            prodthin = False
        End If
        pruneage = Cells(inprunerow, 2)
        prunesph = Cells(inprunerow, 3)
        pruneht = Cells(inprunerow, 4)
        age = Cells(inyieldrow, 7)
        stocking = Cells(inyieldrow, 8)
        sphaftthin = Cells(inyieldrow, 9)
        Height = Cells(inyieldrow, 10)
        height2 = Height
        If topheight Then
            If sphaftthin <> 0 Then height2 = calcMeanht(Height, sphaftthin)
            Height = calcMeanht(Height, stocking)
        End If
        Volume = Cells(inyieldrow, 11)
        Volaftthin = Cells(inyieldrow, 12)
        voldead = Cells(inyieldrow, 13)
        density = Cells(inyieldrow, 14)
        'if input density is whole tree density, calculate sheath density
        If sheathdens Then
            dens_sh = density
            dens_wt = (dens_sh * (Volume - prev_vol) + prev_dens_wt * prev_vol) / Volume
        Else
            dens_wt = density
            dens_sh = (dens_wt * Volume - prev_dens_wt * prev_vol) / (Volume - prev_vol)
        End If
        GrossVol = Volume + voldead
        stocking = Cells(inyieldrow, 8)
        inyieldrow = inyieldrow + 1
        
        Worksheets("INPUT").Activate
        Cells(outyieldrow, 1) = age
        Cells(outyieldrow, 2) = stocking
        Cells(outyieldrow, 3) = Height
        Cells(outyieldrow, 4) = Volume
        Cells(outyieldrow, 5) = GrossVol
        Cells(outyieldrow, 6) = Volume  'Dummy value representing BA
        Cells(outyieldrow, 7) = dens_wt * 0.001
        Cells(outyieldrow, 8) = dens_sh * 0.001
        outyieldrow = outyieldrow + 1
        
        If (age = thinage Or age = pruneage) And age <> 0 Then
            Cells(outdistrow, 1) = age
            If age = thinage And thinage <> 0 Then
                stocking = sphaftthin
                Volume = Volaftthin
                GrossVol = Volume + voldead
                Cells(outdistrow, 2) = thinsph
                Cells(outdistrow, 3) = 30   'Dummy BA after thin value
                If prodthin Then Cells(outdistrow, 5) = prodthinpct * 0.01
                If prodthin Then Cells(outdistrow, 6) = crownremovalpct * 0.01
                inthinrow = inthinrow + 1
            End If
            If age = pruneage And pruneage <> 0 Then
                If prunesph = 0 Or prunesph > stocking Then prunesph = stocking
                prelements = prelements + 1
                prht(prelements) = pruneht
                prsph(prelements) = prunesph
                totalprht = 0
                totalprsph = 0
                For i = prelements To 1 Step -1
                    If prsph(i) > stocking Then prsph(i) = stocking
                    If totalprsph < prsph(i) Then
                        totalprht = totalprht + (prsph(i) - totalprsph) * prht(i)
                        totalprsph = prsph(i)
                    End If
                Next i
                avprht = totalprht / stocking
                
                Cells(outdistrow, 4) = avprht
                inprunerow = inprunerow + 1
            End If
            
            Cells(outyieldrow, 1) = age
            Cells(outyieldrow, 2) = stocking
            Cells(outyieldrow, 3) = height2
            Cells(outyieldrow, 4) = Volume
            Cells(outyieldrow, 5) = GrossVol
            Cells(outyieldrow, 6) = Volume  'Dummy value representing BA
            Cells(outyieldrow, 7) = dens_wt * 0.001
            Cells(outyieldrow, 8) = dens_sh * 0.001
            outyieldrow = outyieldrow + 1
               
            nodisturb = nodisturb + 1
            outdistrow = outdistrow + 1
            
        End If
        prev_vol = Volume
        prev_dens_wt = dens_wt
    Loop Until age = rotlth
    
    Cells(outdistrow, 1) = rotlth1
    Cells(outdistrow, 3) = -2
    Cells(outdistrow, 5) = clearfellpct * 0.01
    Cells(outdistrow, 6) = crownremovalpct * 0.01
    Cells(outdistrow, 7) = forestfloorremovalpct * 0.01
    Cells(4, 5) = nodisturb
    Cells(17, 1) = rotlth1 + nodisturb - 1
    
    Range(Cells(3, 1), Cells(outyieldrow - 1, 10)).Copy
    Cells(18 + rotlth1 + nodisturb, 1).Select
    ActiveSheet.Paste
    Range(Cells(17 + rotlth1 + nodisturb, 1), Cells(17 + rotlth1 + nodisturb, 10)).ClearContents
    Cells(17 + rotlth1 + nodisturb, 1) = "Rotation 2"
    Cells(19 + rotlth1 + nodisturb, 1) = rotlth2
    Cells(19 + rotlth1 + nodisturb * 2, 1) = rotlth2
    Cells(17 + rotlth1 + nodisturb + 13, 7) = 2
    Cells(19 + rotlth1 + nodisturb + 13, 1) = rotlth2 + nodisturb - 1
    
    Range(Cells(34 + rotlth1 + rotlth2 + (nodisturb - 1) * 2 + 1, 1), Cells(10000, 10)).ClearContents
    
    Call DRYMAT
    
    inrow = 3
        
    'C_Change adjustments for non-radiata pine species
    Stembark_adj = 1
    Stemwood_adj = 1
    Deadwood_adj = 1
    Crown_adj = 1
    Litter_adj = 1
        
    If species = "PMEN" Then    'Implementation for Douglas-fir using the 500 Index model
        Stembark_adj = 1.3034
        Stemwood_adj = 1.031
        Crown_adj = 1.5109
        Litter_adj = 0.9491
        Deadwood_adj = 1.1243
    End If
    
    If species = "CLUS" Then    'Implementation for Cupressus lustitanica and other cypresses
        Stembark_adj = 1.0729
        Stemwood_adj = 1.0729
        Crown_adj = 1.7561
        Litter_adj = 1.7561
        Deadwood_adj = 1.0729
    End If
    
    If species = "EUC" Then    'Implementation for Euclayptus species and other hardwoods
        Stembark_adj = 1.1572
        Stemwood_adj = 1.1572
        Crown_adj = 0.3725
        Litter_adj = 0.3725
        Deadwood_adj = 1.1572
    End If
    
    If species = "DFIR" Then    'Implementation for Douglas-fir using the 300 Index model
        Stembark_adj = 1.0877
        Stemwood_adj = 1.0877
        Deadwood_adj = 1.0877
        Crown_adj = 1.8982
        Litter_adj = 1.8763
    End If
    
    'First rotation
    Do
        Worksheets("LP1OUT").Activate
        age = Cells(inrow, 1)
        
        'Carbon fractions for softwoods (radiata, Douglas-fir & lusitanica)
        CFraction_wood = 0.498
        CFraction_roots = 0.501
        CFraction_needles = 0.514
        CFraction_branches = 0.507
        CFraction_DWL = 0.507
        If age >= 5 Then CFraction_bark = 0.551 * (1 - 0.291 * Exp(-0.28 * age))
        If age < 5 Then CFraction_bark = 0.503
        
        'Use IPCC default temperate broadleaf carbon fraction for hardwoods (Eucalypts)
        If species = "EUC" Then
            CFraction_wood = 0.48
            CFraction_roots = 0.48
            CFraction_needles = 0.48
            CFraction_branches = 0.48
            CFraction_DWL = 0.48
            CFraction_bark = 0.48
        End If
       
        'Calculate AGL, BGL, DWL and FL carbon
        UnAdj_AGL = Cells(inrow, 9) + _
            Cells(inrow, 17) + _
            (Cells(inrow, 4) + Cells(inrow, 5) + Cells(inrow, 6)) + _
            (Cells(inrow, 7) + Cells(inrow, 8))
        AGL = Cells(inrow, 9) * Stemwood_adj * CFraction_wood + _
            Cells(inrow, 17) * Stembark_adj * CFraction_bark + _
            (Cells(inrow, 4) + Cells(inrow, 5) + Cells(inrow, 6)) * Crown_adj * CFraction_needles + _
            (Cells(inrow, 7) + Cells(inrow, 8)) * Crown_adj * CFraction_branches
        Adj_ratio = AGL / UnAdj_AGL
        BGL = (Cells(inrow, 10) + Cells(inrow, 15)) * Adj_ratio
        DWL = (Cells(inrow, 13) + Cells(inrow, 14)) * Deadwood_adj * CFraction_DWL
        FL = (Cells(inrow, 11) + Cells(inrow, 12)) * Litter_adj * CFraction_needles
        
        CSHRUB = Cells(inrow, 27)
        Total = AGL + BGL + DWL + FL
        inrow = inrow + 1
        
        Worksheets("C Change").Activate
        Cells(6 + age, 16) = Total
        Cells(6 + age, 17) = AGL
        Cells(6 + age, 18) = BGL
        Cells(6 + age, 19) = DWL
        Cells(6 + age, 20) = FL
        Cells(6 + age, 27) = CSHRUB
    Loop Until age = rotlth1
    
    'Second rotation
    inrow = inrow + 1
    Do
        Worksheets("LP1OUT").Activate
        age = Cells(inrow, 1)
       
        'Carbon fractions for softwoods (radiata, Douglas-fir & lusitanica)
        CFraction_wood = 0.498
        CFraction_roots = 0.501
        CFraction_needles = 0.514
        CFraction_branches = 0.507
        CFraction_DWL = 0.507
        If age >= 5 Then CFraction_bark = 0.551 * (1 - 0.291 * Exp(-0.28 * age))
        If age < 5 Then CFraction_bark = 0.503
        
        'Use IPCC default temperate broadleaf carbon fraction for hardwoods (Eucalypts)
        If species = "EUC" Then
            CFraction_wood = 0.48
            CFraction_roots = 0.48
            CFraction_needles = 0.48
            CFraction_branches = 0.48
            CFraction_DWL = 0.48
            CFraction_bark = 0.48
        End If
       
        'Calculate AGL, BGL, DWL and FL carbon
        UnAdj_AGL = Cells(inrow, 9) + _
            Cells(inrow, 17) + _
            (Cells(inrow, 4) + Cells(inrow, 5) + Cells(inrow, 6)) + _
            (Cells(inrow, 7) + Cells(inrow, 8))
        AGL = Cells(inrow, 9) * Stemwood_adj * CFraction_wood + _
            Cells(inrow, 17) * Stembark_adj * CFraction_bark + _
            (Cells(inrow, 4) + Cells(inrow, 5) + Cells(inrow, 6)) * Crown_adj * CFraction_needles + _
            (Cells(inrow, 7) + Cells(inrow, 8)) * Crown_adj * CFraction_branches
        Adj_ratio = AGL / UnAdj_AGL
        BGL = (Cells(inrow, 10) + Cells(inrow, 15)) * Adj_ratio
        DWL = (Cells(inrow, 13) + Cells(inrow, 14)) * Deadwood_adj * CFraction_DWL
        FL = (Cells(inrow, 11) + Cells(inrow, 12)) * Litter_adj * CFraction_needles
        
        CSHRUB = Cells(inrow, 27)
        Total = AGL + BGL + DWL + FL
        inrow = inrow + 1
        
        Worksheets("C Change").Activate
        Cells(6 + age, 21) = Total
        Cells(6 + age, 22) = AGL
        Cells(6 + age, 23) = BGL
        Cells(6 + age, 24) = DWL
        Cells(6 + age, 25) = FL
        Cells(6 + age, 28) = CSHRUB
    Loop Until age = rotlth2
    Application.ScreenUpdating = True
    
End Sub

Sub Yield_Table()
'This subroutine generates a yield table in the C Change worksheet using 300 Index Model or 500 Index Model
        
    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
    
    If LCase(Worksheets("C Change").Cells(3, 3)) = "pmen" Then
        'Use 500 Index Model
        
        'Adjust options on sheet to account for the fact that the 500 Index model predicts MTH and sheath density
        Worksheets("C Change").Cells(58, 3) = "S"
        Worksheets("C Change").Cells(59, 3) = "T"
        
        'Transfer information from C Change sheet to 500 Index sheet
        Call transfer_CC_500I
        
        If Worksheets("C Change").Cells(60, 3) = 1 Then
            'Mode 1 - standard mode: use calibration measurement if specified, otherwise use specified productivity indices
            If Worksheets("C Change").Cells(29, 3) = 0 Then
                'If there is no calibration measurement, produce yield table for specified SI & 500 Index
                Worksheets("500 Index").Calculate
            Else
                'Calculate 500 Index & SI and generate yield table
                Call CombineSolver
                'Transfer 500 Index & SI estimates to C Change worksheet
                Worksheets("C Change").Cells(41, 3) = Worksheets("500 Index").Cells(2, 2)
                Worksheets("C Change").Cells(42, 3) = Worksheets("500 Index").Cells(3, 2)
            End If
        Else
            'Modes 2 or 3, produce yield table for specified SI & 500 Index with (mode=2) or without (mode=3) offset
            Worksheets("500 Index").Calculate
        End If
        
        'Predict volume yield for first few years of growth
        Call earlyield_dfir
        
        'Transfer yield predictions from 500 Index sheet to C Change sheet
        Call transfer_500I_CC
 
    Else
        'Use 300 Index Model
        
        'Transfer information from C Change sheet to 300 Index sheet
        Call transfer_CC_300I
        
        'Calculate height model coefficients
        Worksheets("300 Index").Activate
        Call calcheightcoeff(SI)
        
        'Estimate Site Index if required
        If Worksheets("C Change").Cells(30, 3) <> 0 Then Call siteIndex
        
        'Fill in missing stand history
        Call Prune_History
        Call Thin_history
        
        'Transfer information from C Change sheet to 300 Index sheet
        Call transfer_CC_300I
    
        Worksheets("300 Index").Activate
        If Worksheets("C Change").Cells(60, 3) = 1 Then
            'Mode 1 - standard mode: use calibration measurement if specified, otherwise use specified productivity indices
            If Worksheets("C Change").Cells(29, 3) = 0 Then
                'If there is no calibration measurement, produce yield table for specified SI & 300 Index
                Call OutputGrowth
            Else
                'Calculate 300 Index & SI and generate yield table
                Call siteIndex
                Call Calc300Index
                'Copy 300 Index & SI estimates into C Change worksheet
                Worksheets("C Change").Cells(41, 3) = Worksheets("300 Index").Cells(3, 3)
                Worksheets("C Change").Cells(42, 3) = Worksheets("300 Index").Cells(4, 3)
            End If
        Else
            'Modes 2 or 3, produce yield table for specified SI & 300 Index with (mode=2) or without (mode=3) offset
            Call OutputGrowth
        End If
        
        'Transfer yield predictions from 300 Index sheet to C Change sheet
        Call transfer_300I_CC
    
    End If
    
    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
        
End Sub

Sub transfer_CC_300I()
'This suboutine transfers information from "C Change" worksheet to "300 Index" worksheet
    Dim latitude As Double, altitude As Double, calage As Double, calsph As Double, calBA As Double, _
        calMTH As Double, I300 As Double, SI As Double, SoilC As Double, SoilN As Double, Temp As Double, _
        GeneticAdj As Double, CoreDens As Double, CoreAge As Long, InnerRing As Long, _
        OuterRing As Long, densitymodel As Long, j As Long, lastthinage As Double, _
        drift As Double, mortadd As Double, mortmult As Double
    
    Worksheets("C Change").Activate
    Initialsph = Cells(4, 3) * Cells(39, 3) / 100
    I300 = Cells(41, 3)
    SI = Cells(42, 3)
    latitude = Cells(33, 3)
    altitude = Cells(34, 3)
    rotlth1 = Cells(5, 3)
    rotlth2 = Cells(6, 3)
    If rotlth2 > rotlth1 Then
        rotlth = rotlth2
    Else
        rotlth = rotlth1
    End If
    calage = Cells(27, 3)
    calsph = Cells(28, 3)
    calBA = Cells(29, 3)
    calMTH = Cells(30, 3)
    SoilC = Cells(35, 3)
    SoilN = Cells(36, 3)
    Temp = Cells(40, 3)
    
    'Default values
    If SoilC = 0 Then SoilC = 5.57
    If SoilN = 0 Then SoilN = 0.296
    If Temp = 0 Then Temp = 12
        
    CoreDens = Cells(45, 3)
    CoreAge = Cells(46, 3)
    InnerRing = Cells(47, 3)
    OuterRing = Cells(48, 3)
    
    drift = Cells(51, 3)
    mortadd = Cells(52, 3)
    mortmult = Cells(53, 3)
    
    densitymodel = 2    'Use 2010 wood density model
    inthinrow = 11
    inprunerow = 20
    inyieldrow = 6
    
    Worksheets("300 Index").Activate
    Cells(8, 6) = Worksheets("C Change").Cells(60, 3)  'Mode
    Range(Cells(75, 4), Cells(82, 4)).ClearContents
    Range(Cells(3, 3), Cells(4, 3)).ClearContents
    Range(Cells(3, 6), Cells(4, 6)).ClearContents
    Range(Cells(7, 3), Cells(11, 3)).ClearContents
    Range(Cells(14, 3), Cells(15, 3)).ClearContents
    Cells(19, 3).ClearContents
    Range(Cells(20, 2), Cells(36, 6)).ClearContents
    Range(Cells(40, 2), Cells(44, 4)).ClearContents
    If latitude <> 0 Then Cells(3, 6) = Abs(latitude)
    If altitude <> 0 Then Cells(4, 6) = altitude
    If I300 <> 0 Then Cells(3, 3) = I300
    If SI <> 0 Then Cells(4, 3) = SI
    If SoilC <> 0 Then Cells(75, 4) = SoilC
    If SoilN <> 0 Then Cells(76, 4) = SoilN
    If Temp <> 0 Then Cells(77, 4) = Temp
    If CoreDens <> 0 Then Cells(78, 4) = CoreDens
    If CoreAge <> 0 Then Cells(79, 4) = CoreAge
    If InnerRing <> 0 Then Cells(80, 4) = InnerRing
    If OuterRing <> 0 Then Cells(81, 4) = OuterRing
    Cells(83, 4) = 2    'Use 2010 wood density model
            
    Cells(64, 6) = drift    'Drift parameter
    If mortadd <> 0 Then    'Mortality model adjustments
        Cells(68, 6) = mortadd
    Else
        Cells(68, 6).ClearContents
    End If
    If mortmult <> 0 Then
        Cells(69, 6) = mortmult
    Else
        Cells(69, 6).ClearContents
    End If
   
    Cells(47, 3) = rotlth
    Cells(19, 3) = Initialsph
    If calage <> 0 Then
        Cells(7, 3) = calage
        Cells(14, 3) = calage
        Cells(8, 3) = calsph
        Cells(10, 3) = calBA
        Cells(15, 3) = calMTH
    End If
    
    j = 0
    lastthinage = 0
    For i = 1 To 5
        Worksheets("C Change").Activate
        thinage = Cells(inthinrow - 1 + i, 2)
        thinsph = Cells(inthinrow - 1 + i, 3)
        Worksheets("300 Index").Activate
        If calage <> 0 And calage > lastthinage And (calage < thinage Or thinage = 0) Then
            Cells(20 + j, 2) = calage
            Cells(20 + j, 3) = calsph
            j = j + 1
        End If
        If thinage = 0 Then Exit For
        Cells(20 + j, 2) = thinage
        Cells(20 + j, 4) = thinsph
        lastthinage = thinage
        j = j + 1
    Next i
    
    lastpruneage = 0
    j = 0
    For i = 1 To 5
        Worksheets("C Change").Activate
        pruneage = Cells(inprunerow - 1 + i, 2)
        prunesph = Cells(inprunerow - 1 + i, 3)
        pruneht = Cells(inprunerow - 1 + i, 4)
        If pruneage <> 0 Then
            Worksheets("300 Index").Activate
            If pruneage > lastpruneage Then j = j + 1
            Cells(39 + j, 2) = pruneage
            Cells(39 + j, 4) = prunesph
            Cells(39 + j, 3) = pruneht
            lastpruneage = pruneage
        End If
    Next i
End Sub

Sub transfer_300I_CC()
    'Copy predicted yield table from "300 Index" worksheet into "C Change" worksheet
    Worksheets("C Change").Activate
    
    Range(Cells(6, 7), Cells(1000, 28)).ClearContents
    
    Worksheets("300 Index").Activate
    Range(Cells(5, 7), Cells(5 + rotlth, 9)).Copy
    Worksheets("C Change").Activate
    Cells(6, 7).Select
    Selection.PasteSpecial Paste:=xlPasteValues
        
    If LCase(Cells(59, 3)) = "m" Then
        Worksheets("300 Index").Activate
        Range(Cells(5, 18), Cells(5 + rotlth, 18)).Copy
    Else
        Worksheets("300 Index").Activate
        Range(Cells(5, 10), Cells(5 + rotlth, 10)).Copy
    End If
    Worksheets("C Change").Activate
    Cells(6, 10).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Worksheets("300 Index").Activate
    Range(Cells(5, 12), Cells(5 + rotlth, 13)).Copy
    Worksheets("C Change").Activate
    Cells(6, 11).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Worksheets("300 Index").Activate
    Range(Cells(5, 39), Cells(5 + rotlth, 39)).Copy
    Worksheets("C Change").Activate
    Cells(6, 13).Select
    Selection.PasteSpecial Paste:=xlPasteValues
        
    If LCase(Cells(58, 3)) = "t" Then
        Worksheets("300 Index").Activate
        Range(Cells(5, 41), Cells(5 + rotlth, 41)).Copy
    Else
        Worksheets("300 Index").Activate
        Range(Cells(5, 40), Cells(5 + rotlth, 40)).Copy
    End If
    Worksheets("C Change").Activate
    Cells(6, 14).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    For i = 0 To rotlth
        Cells(6 + i, 14) = Cells(6 + i, 14) * 1000
    Next i
    
End Sub

Sub Thin_history()
'This subroutine predicts missing thinning & stocking history
    Dim calage As Double, calsph As Double, i As Long, npremeasthin As Long, thinage As Long, _
        predict_initial As Boolean, thinratio As Double, predictrow As Long, predictsph As Double, _
        j As Long
        
    'Find number of pre-measurement thinnings
    Worksheets("C Change").Activate
    calage = Cells(27, 3)
    calsph = Cells(28, 3)
    npremeasthin = 0
    For i = 1 To 5
        If Cells(10 + i, 1) <> 0 And Cells(10 + i, 2) <= calage Then npremeasthin = npremeasthin + 1
    Next i
    
    'If number of pre-measurement thinnings=1 and thinning age unknown then predict missing thinning age
    If npremeasthin = 1 And Cells(11, 2) = 0 Then
        thinage = Int(CalcAgefromMTH(SI, 12) + 0.5)
        If thinage > calage Then thinage = Int(calage)
        Cells(11, 2) = thinage
    End If
    
    'If number of pre-measurement thinnings=1 and initial stocking unknown then predict initial stocking
    predict_initial = False
    If npremeasthin = 1 And Cells(4, 3) = 0 Then
        predict_initial = True
        thinratio = 322 * calsph ^ -0.833
        If thinratio < 1 Then thinratio = 1
        Cells(4, 3) = thinratio * calsph
    End If
    
    'If number of pre-measurement thinnings=0 and initial stocking unknown, then predict it
    predict_initial = False
    If npremeasthin = 0 And Cells(4, 3) = 0 Then
        predict_initial = True
        predictrow = 4
        predictsph = calsph + 0.0001
        Cells(predictrow, 3) = predictsph
        Call transfer_CC_300I
        Worksheets("300 Index").Activate
        Call siteIndex
        Call Calc300Index
        'Copy 300 Index & SI estimates into C Change worksheet
        Worksheets("C Change").Cells(41, 3) = Worksheets("300 Index").Cells(3, 3)
        Worksheets("C Change").Cells(42, 3) = Worksheets("300 Index").Cells(4, 3)
        For i = 1 To 3
            Worksheets("C Change").Activate
            Cells(27, 3).ClearContents
            Cells(28, 3).ClearContents
            Call transfer_CC_300I
            Worksheets("300 Index").Activate
            Call OutputGrowth
            Call transfer_300I_CC
            predictsph = calsph / Cells(Int(calage) + 6, 8) * predictsph
            Worksheets("C Change").Cells(predictrow, 3) = predictsph
        Next i
        Worksheets("C Change").Activate
        Cells(27, 3) = calage
        Cells(28, 3) = calsph
    End If
    
    'If thin stocking in last thinning prior to measurement unknown, then predict it
    predictrow = 10 + npremeasthin
    If npremeasthin = 0 Then predictrow = 4
    predictsph = calsph
    If Cells(predictrow, 3) = 0 Then
        Cells(predictrow, 3) = predictsph
        If predict_initial Then
            thinratio = 322 * calsph ^ -0.833
            If thinratio < 1 Then thinratio = 1
            Cells(4, 3) = thinratio * predictsph
        End If
        
        Call transfer_CC_300I
        Worksheets("300 Index").Activate
        Call siteIndex
        Call Calc300Index
        'Copy 300 Index & SI estimates into C Change worksheet
        Worksheets("C Change").Cells(41, 3) = Worksheets("300 Index").Cells(3, 3)
        Worksheets("C Change").Cells(42, 3) = Worksheets("300 Index").Cells(4, 3)
        For i = 1 To 3
            Worksheets("C Change").Activate
            Cells(27, 3).ClearContents
            Cells(28, 3).ClearContents
            Call transfer_CC_300I
            Worksheets("300 Index").Activate
            Call OutputGrowth
            Call transfer_300I_CC
            predictsph = calsph / Cells(Int(calage) + 6, 8) * predictsph
            Worksheets("C Change").Cells(predictrow, 3) = predictsph
        Next i
        Worksheets("C Change").Activate
        Cells(27, 3) = calage
        Cells(28, 3) = calsph
        
    End If
    
End Sub

Sub Prune_History()
'This subroutine predicts missing prune history
    Dim lastprht As Double, lastprsph As Double, std_prht(5) As Double, j As Long
        
    'Pruning history
    Worksheets("C Change").Activate
    std_prht(1) = 2.4
    std_prht(2) = 4.6
    std_prht(3) = 6
    std_prht(4) = 7.5
    std_prht(5) = 9
    
    For j = 5 To 1 Step -1
        If Cells(19 + j, 1) <> 0 Then
            If Cells(19 + j, 4) <> 0 Then
                lastprht = Cells(19 + j, 4)
            Else
                Cells(19 + j, 4) = std_prht(j)
                If lastprht < std_prht(j) Then Cells(19 + j, 4) = lastprht - 0.1
                lastprht = Cells(19 + j, 4)
            End If
            If Cells(19 + j, 2) = 0 Then
                'If prune age unknown find age when crown length = 4.5
                Cells(19 + j, 2) = Int(CalcAgefromMTH(SI, lastprht + 4.5) + 0.5)
            End If
            If Cells(19 + j, 3) <> 0 Then
                lastprsph = Cells(19 + j, 3)
            Else
                Cells(19 + j, 3) = lastprsph / 0.95
                lastprsph = Cells(19 + j, 3)
            End If
        End If
    Next j
End Sub

Function CalcAgefromMTH(SI As Double, mth As Double)
    CalcAgefromMTH = -Log(-(1 - Exp(-ha * 20)) * ((mth - 0.25) / (SI - 0.25)) ^ (1 / hb) + 1) / ha
End Function

Sub Batch_run_control()
'Predicts carbon in batch model using either:
'   standard method - independent estimates for each measurement
'   drift factor method - estimates a drift factor for PRAD plots with multiple measurements and reruns plots using this drift fatcor
    Dim numplots As Long, numrows As Long, nmeasurements As Long, plot As String, species As String, age As Double, I300 As Double, _
        Drift_factor As Double, plotnum As Long, row As Long, prev_plot As String, first_age As Double, first_I300 As Double
    
    If LCase(Worksheets("C_Change control").Cells(10, 2)) = "y" Then Worksheets("C Change").Cells(51, 3) = 0 'Set drift factor to 0 for first set of runs
    Call Batch_run
    
    If LCase(Worksheets("C_Change control").Cells(10, 2)) = "y" Then
        Worksheets("Plots").Activate
        numplots = Range(Cells(2, 1), Cells(2, 1).End(xlDown)).Rows.Count
        Worksheets("Plots Processed").Activate
        numrows = Range(Cells(2, 1), Cells(2, 1).End(xlDown)).Rows.Count
        prev_plot = ""
        nmeasurements = 1
        For row = 2 To numrows + 1
            plot = Worksheets("Plots Processed").Cells(row, 1)
            If plot = prev_plot Then
                nmeasurements = nmeasurements + 1
            Else
                If species = "PRAD" And nmeasurements > 1 Then
                    age = Worksheets("Plots Processed").Cells(row - 1, 2)
                    I300 = Worksheets("Plots Processed").Cells(row - 1, 3)
                    Drift_factor = (I300 - first_I300) / (age - first_age)
                    Worksheets("Plots").Activate
                    plotnum = WorksheetFunction.Match(prev_plot, Range(Cells(2, 1), Cells(numplots + 1, 1)), 0)
                    Worksheets("Plots").Cells(plotnum + 1, 18) = Drift_factor
                End If
                prev_plot = plot
                species = Worksheets("Plots Processed").Cells(row, 10)
                first_age = Worksheets("Plots Processed").Cells(row, 2)
                first_I300 = Worksheets("Plots Processed").Cells(row, 3)
                nmeasurements = 1
            End If
        Next row
        
        Call Batch_run
        
    End If
End Sub

Sub Batch_run()
'This subroutine predicts carbon in batch mode
    Dim startrow As Long, endrow As Long, outputrange As Object, _
        thinage As Double, thinsph As Double, pruneage As Double, prunesph As Double, pruneht As Double, _
        age As Double, sph As Double, sphaft As Double, mth As Double, vol As Double, volaft As Double, _
        voldead As Double, sheath_density As Double, ow_density As Double, tree_density As Double, _
        specindex As Boolean, plot As String, lastplot As String, plotrow As Long, _
        outrow As Long, latitude As Double, altitude As Double, _
        prodthin(200) As Boolean, SI As Double, I300 As Double, aage As Double, _
        stockline As Long, pruneline As Long, laststock As Double, i As Long, j As Long, _
        pspprsph As Double, pspprht As Double, numplots As Long, plotnum As Long, NR As Double, _
        row As Long, startplotrow As Long, indexage As Double, numrows As Long, yieldline As Long, _
        SoilC As Double, SoilN As Double, SoilOrganicP As Double, SoilBray2P As Double, _
        Early_survival As Double, Temp As Double, rotlth1 As Double, rotlth2 As Double, _
        GeneticAdj As Double, CoreDens As Double, CoreAge As Long, InnerRing As Long, OuterRing As Long, _
        drift As Double, predicted As Boolean, D200 As Double, rotlth As Double, _
        BA As Double, BA2 As Double, stock As Double, prsph As Double, prht As Double, _
        mortadd As Double, mortmult As Double, recordtype As String, _
        C1 As Double, C2 As Double, C As Double, _
        Initial_DryMat(5) As Double, Initial_N(5) As Double, Initial_P(5) As Double, _
        NutContTableRow As Long
        
    RunCChange = False
    If UCase(Sheets("C_Change control").Cells(8, 2)) = "Y" Then RunCChange = True
    RunNuBalM = False
    If UCase(Sheets("C Change").Cells(61, 3)) = "Y" Then RunNuBalM = True
    Application.ScreenUpdating = False
    Worksheets("C_Change control").Activate
    startrow = Cells(3, 2)
    endrow = Cells(4, 2)
    rotlth1 = Cells(5, 2)
    rotlth2 = Cells(6, 2)
    Sheets("C Change").Cells(58, 3) = "S"
    Sheets("C Change").Cells(59, 3) = "T"
    Sheets("C Change").Cells(5, 3) = rotlth1
    Sheets("C Change").Cells(6, 3) = rotlth2
    Sheets("C Change").Cells(61, 3) = Sheets("C_Change control").Cells(12, 2)

    Worksheets("Plots").Activate
    numplots = Range(Cells(2, 1), Cells(2, 1).End(xlDown)).Rows.Count
    
    Worksheets("Plots processed").Activate
    Range(Cells(2, 1), Cells(200000, 10)).ClearContents 'Clear any existing SI, I300 values
    Worksheets("C_Change Predictions").Range("A5:X200000").ClearContents
    Worksheets("Yield Tables").Range("A4:X200000").ClearContents
    Worksheets("Nutrient Content Tables").Range("A3:AD200000").ClearContents
    Worksheets("C_Change Output").Range("A1:AC200000").ClearContents
    Worksheets("C_Change Log").Range("A1:X200000").ClearContents
    
    Worksheets("PSP summary").Activate
    plot = "start"
    yieldline = 4
    plotrow = 2
    specindex = False
    plotrow = 2
    NutContTableRow = 3
   
    For row = startrow To endrow + 1
        
        'Display progress on status bar
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True ' makes sure that the statusbar is visible
        Application.StatusBar = "Processing row" + Str(row - startrow) + " out of" + Str(endrow - startrow + 1)
        Application.ScreenUpdating = False
        
        'Read next record
        Worksheets("PSP summary").Activate
        lastplot = plot
        plot = Cells(row, 1)
        
        'Test for new plot
        If (plot <> lastplot And row <> startrow) Or row = endrow + 1 Or _
            (LCase(Worksheets("C_Change control").Cells(11, 2)) <> "y" And recordtype = "M") Then
            
            'Get plot-level parameters for previous plot
            Worksheets("Plots").Activate
            plotnum = WorksheetFunction.Match(lastplot, Range(Cells(2, 1), Cells(numplots + 1, 1)), 0)
            species = Cells(plotnum + 1, 2)
            latitude = Cells(plotnum + 1, 4)
            altitude = Cells(plotnum + 1, 5)
            NR = Cells(plotnum + 1, 6)
            SoilC = Cells(plotnum + 1, 7)
            SoilN = Cells(plotnum + 1, 8)
            SoilOrganicP = Cells(plotnum + 1, 9)
            Early_survival = Cells(plotnum + 1, 10)
            Temp = Cells(plotnum + 1, 11)
            CoreDens = Cells(plotnum + 1, 12)
            CoreAge = Cells(plotnum + 1, 13)
            InnerRing = Cells(plotnum + 1, 14)
            OuterRing = Cells(plotnum + 1, 15)
            I300 = Cells(plotnum + 1, 16)
            SI = Cells(plotnum + 1, 17)
            drift = Cells(plotnum + 1, 18)
            mortadd = Cells(plotnum + 1, 19)
            mortmult = Cells(plotnum + 1, 20)
            
            For i = 1 To 5
                Initial_DryMat(i) = Cells(plotnum + 1, 20 + i)
                Initial_N(i) = Cells(plotnum + 1, 25 + i)
                Initial_P(i) = Cells(plotnum + 1, 30 + i)
            Next i
            
            'Enter plot-level parameters into "C Change" worksheet
            Worksheets("C Change").Activate
            Cells(3, 3) = species
            Range(Cells(33, 3), Cells(53, 3)).ClearContents
            If latitude <> 0 Then Cells(33, 3) = latitude
            If altitude <> 0 Then Cells(34, 3) = altitude
            If SoilC <> 0 Then Cells(35, 3) = SoilC
            If SoilN <> 0 Then Cells(36, 3) = SoilN
            If SoilOrganicP <> 0 Then Cells(37, 3) = SoilOrganicP
            If Early_survival <> 0 Then Cells(39, 3) = Early_survival
            If Temp <> 0 Then Cells(40, 3) = Temp
            If I300 <> 0 Then Cells(41, 3) = I300
            If SI <> 0 Then Cells(42, 3) = SI
            If CoreDens <> 0 Then Cells(45, 3) = CoreDens
            If CoreAge <> 0 Then Cells(46, 3) = CoreAge
            If InnerRing <> 0 Then Cells(47, 3) = InnerRing
            If OuterRing <> 0 Then Cells(48, 3) = OuterRing
            If drift <> 0 Then Cells(51, 3) = drift
            If mortadd <> 0 Then Cells(52, 3) = mortadd
            If mortmult <> 0 Then Cells(53, 3) = mortmult
            If NR <> 0 Then Cells(38, 3) = NR
             
            For i = 1 To 5
                Cells(63 + i, 3) = Initial_DryMat(i)
                Cells(63 + i, 4) = Initial_N(i)
                Cells(63 + i, 5) = Initial_P(i)
            Next i
            
            Call Yield_Table
            Application.ScreenUpdating = False
            
            If RunCChange Then
                Call Run_C_Change
                Application.ScreenUpdating = False
            
                'Copy carbon predictions to C_Change prediction worksheet
                aage = Worksheets("C Change").Cells(27, 3)
                Sheets("C Change").Range("G6:AB6").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Worksheets("C_Change Predictions").Activate
                Cells(5, 3).Select
                If Not IsEmpty(Cells(5, 3)) Then
                    Selection.End(xlDown).Offset(1, 0).Select
                End If
                ActiveSheet.Paste
                Application.CutCopyMode = False
                ActiveCell.Offset(0, -2) = lastplot
                If aage <> 0 Then ActiveCell.Offset(0, -1) = aage
    
                'Update nutrient contents table
                If RunNuBalM Then
                    Sheets("Nutrient Content Tables").Cells(NutContTableRow, 1) = lastplot
                    Sheets("Nutrient Content Tables").Cells(NutContTableRow, 2) = aage
                    For i = 1 To 200
                        If IsEmpty(Sheets("LP1OUT").Cells(2 + i, 1)) Then Exit For
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 3) = Sheets("LP1OUT").Cells(2 + i, 1) 'Age
                
                        'Dry weight
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 4) = Sheets("LP1OUT").Cells(2 + i, 4) '1yr foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 5) = Sheets("LP1OUT").Cells(2 + i, 5) + Sheets("LP1OUT").Cells(2 + i, 6)    '2yr+ foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 6) = Sheets("LP1OUT").Cells(2 + i, 7)   'live br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 7) = Sheets("LP1OUT").Cells(2 + i, 8)   'dead br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 8) = Sheets("LP1OUT").Cells(2 + i, 17)   'stem bark
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 9) = Sheets("LP1OUT").Cells(2 + i, 9)   'stem wood
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 10) = Sheets("LP1OUT").Cells(2 + i, 10) + Sheets("LP1OUT").Cells(2 + i, 15) 'live roots = coarse + fine roots
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 11) = Sheets("LP1OUT").Cells(2 + i, 13) 'stem litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 12) = Sheets("LP1OUT").Cells(2 + i, 11) + Sheets("LP1OUT").Cells(2 + i, 12) 'branch litter + needle litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 13) = Sheets("LP1OUT").Cells(2 + i, 14) 'coarse root litter
                
                        'Nitrogen content
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 14) = Sheets("Nitrogen").Cells(4 + i, 2) '1yr foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 15) = Sheets("Nitrogen").Cells(4 + i, 3) + Sheets("Nitrogen").Cells(4 + i, 4)    '2yr+ foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 16) = Sheets("Nitrogen").Cells(4 + i, 5)   'live br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 17) = Sheets("Nitrogen").Cells(4 + i, 6)   'dead br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 18) = Sheets("Nitrogen").Cells(4 + i, 15)   'stem bark
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 19) = Sheets("Nitrogen").Cells(4 + i, 7)   'stem wood
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 20) = Sheets("Nitrogen").Cells(4 + i, 8) + Sheets("Nitrogen").Cells(4 + i, 13) 'live roots = coarse + fine roots
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 21) = Sheets("Nitrogen").Cells(4 + i, 11) 'stem litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 22) = Sheets("Nitrogen").Cells(4 + i, 9) + Sheets("Nitrogen").Cells(4 + i, 10) 'branch + needle litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 23) = Sheets("Nitrogen").Cells(4 + i, 12) 'coarse root litter
                
                        'Phosphorus content
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 24) = Sheets("Phosphorus").Cells(4 + i, 2) '1yr foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 25) = Sheets("Phosphorus").Cells(4 + i, 3) + Sheets("Phosphorus").Cells(4 + i, 4)    '2yr+ foliage
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 26) = Sheets("Phosphorus").Cells(4 + i, 5)   'live br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 27) = Sheets("Phosphorus").Cells(4 + i, 6)   'dead br
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 28) = Sheets("Phosphorus").Cells(4 + i, 15)   'stem bark
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 29) = Sheets("Phosphorus").Cells(4 + i, 7)   'stem wood
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 30) = Sheets("Phosphorus").Cells(4 + i, 8) + Sheets("Phosphorus").Cells(4 + i, 13) 'live roots = coarse + fine roots
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 31) = Sheets("Phosphorus").Cells(4 + i, 11) 'stem litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 32) = Sheets("Phosphorus").Cells(4 + i, 9) + Sheets("Phosphorus").Cells(4 + i, 10) 'branch + needle litter
                        Sheets("Nutrient Content Tables").Cells(NutContTableRow, 33) = Sheets("Phosphorus").Cells(4 + i, 12) 'coarse root litter
                
                        NutContTableRow = NutContTableRow + 1
                    Next i
                End If
            End If
            
            'Output 300 Index & SI
            Worksheets("C Change").Activate
            I300 = Cells(41, 3)
            SI = Cells(42, 3)
            Worksheets("Plots processed").Activate
            Cells(plotrow, 1) = lastplot
            If Not specindex Then Worksheets("Plots processed").Cells(plotrow, 2) = aage
            Worksheets("Plots processed").Cells(plotrow, 3) = I300
            Worksheets("Plots processed").Cells(plotrow, 4) = SI
            For i = 1 To 5
                C1 = Worksheets("C Change").Cells(Int(aage) + 6, i + 15)
                C2 = Worksheets("C Change").Cells(Int(aage) + 7, i + 15)
                C = C1 + (aage - Int(aage)) * (C2 - C1)
                Worksheets("Plots processed").Cells(plotrow, i + 4) = C
            Next i
            Worksheets("Plots processed").Cells(plotrow, 10) = species
            plotrow = plotrow + 1
    
            'Copy 300 Index yield table to Yield tables worksheet
            Worksheets("300 Index").Activate
            Sheets("300 Index").Range("G5:R5").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Worksheets("Yield Tables").Activate
            Cells(4, 3).Select
            If Not IsEmpty(Cells(4, 3)) Then
                Selection.End(xlDown).Offset(1, 0).Select
            End If
            ActiveSheet.Paste
            Application.CutCopyMode = False
            ActiveCell.Offset(0, -2) = lastplot
            If Not specindex Then ActiveCell.Offset(0, -1) = aage
        
            'Produce full C_Change output if required
            If RunCChange And Worksheets("C_Change control").Cells(9, 2) = "Y" Then
                'Tranfer C_Change log file to Worksheet "C_Change Log"
                Worksheets("LP2OUT").Activate
                Sheets("LP2OUT").Range("A1:AA1").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Worksheets("C_Change Log").Activate
                Cells(1, 3).Select
                If Not IsEmpty(Cells(1, 1)) Then
                    Selection.End(xlDown).Offset(1, 0).Select
                End If
                ActiveSheet.Paste
                Application.CutCopyMode = False
                ActiveCell.Offset(0, -2) = lastplot
                If Not specindex Then ActiveCell.Offset(0, -1) = aage
        
                'Transfer complete C_Change output file to Worksheet "C_Change Output"
                Worksheets("LP1OUT").Activate
                Sheets("LP1OUT").Range("A2:AA2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.Copy
                Worksheets("C_Change Output").Activate
                Cells(1, 3).Select
                If Not IsEmpty(Cells(1, 1)) Then
                    Selection.End(xlDown).Offset(1, 0).Select
                End If
                ActiveSheet.Paste
                Application.CutCopyMode = False
                ActiveCell.Offset(0, -2) = lastplot
                If Not specindex Then ActiveCell.Offset(0, -1) = aage
            End If
       End If
       
       'If new plot, clear contents of previous plot
        If (plot <> lastplot Or row = startrow) And row <> endrow + 1 Then
            Worksheets("C Change").Activate
            Cells(3, 3).ClearContents
            Cells(4, 3).ClearContents
            Range(Cells(11, 1), Cells(15, 3)).ClearContents
            Range(Cells(20, 1), Cells(24, 4)).ClearContents
            Range(Cells(27, 3), Cells(30, 3)).ClearContents
            stockline = 11
        End If
        
        'Read next record for plot and copy to "C Change" worksheet
        Worksheets("PSP summary").Activate
        recordtype = Cells(row, 2)
        If recordtype = "E" Then
            stock = Cells(row, 4)
            Worksheets("C Change").Activate
            If stock <> 0 Then Cells(4, 3) = stock
        End If
        If recordtype = "M" Then
            age = Cells(row, 3)
            stock = Cells(row, 4)
            BA = Cells(row, 5)
            mth = Cells(row, 6)
            Worksheets("C Change").Activate
            Cells(27, 3) = age
            Cells(28, 3) = stock
            Cells(29, 3) = BA
            Cells(30, 3) = mth
        End If
        If recordtype = "TW" Then
            age = Cells(row, 3)
            stock = Cells(row, 4)
            Worksheets("C Change").Activate
            Cells(stockline, 1) = "W"
            If age <> 0 Then Cells(stockline, 2) = age
            If stock <> 0 Then Cells(stockline, 3) = stock
            stockline = stockline + 1
        End If
        If recordtype = "TP" Then
            age = Cells(row, 3)
            stock = Cells(row, 4)
            Worksheets("C Change").Activate
            Cells(stockline, 1) = "P"
            If age <> 0 Then Cells(stockline, 2) = age
            If stock <> 0 Then Cells(stockline, 3) = stock
            stockline = stockline + 1
        End If
        If recordtype = "P1" Then
            age = Cells(row, 3)
            prsph = Cells(row, 7)
            prht = Cells(row, 8)
            Worksheets("C Change").Activate
            Cells(20, 1) = 1
            If age <> 0 Then Cells(20, 2) = age
            If prsph <> 0 Then Cells(20, 3) = prsph
            If prht <> 0 Then Cells(20, 4) = prht
        End If
        If recordtype = "P2" Then
            age = Cells(row, 3)
            prsph = Cells(row, 7)
            prht = Cells(row, 8)
            Worksheets("C Change").Activate
            Cells(21, 1) = 2
            If age <> 0 Then Cells(21, 2) = age
            If prsph <> 0 Then Cells(21, 3) = prsph
            If prht <> 0 Then Cells(21, 4) = prht
        End If
        If recordtype = "P3" Then
            age = Cells(row, 3)
            prsph = Cells(row, 7)
            prht = Cells(row, 8)
            Worksheets("C Change").Activate
            Cells(22, 1) = 3
            If age <> 0 Then Cells(22, 2) = age
            If prsph <> 0 Then Cells(22, 3) = prsph
            If prht <> 0 Then Cells(22, 4) = prht
        End If
        If recordtype = "P4" Then
            age = Cells(row, 3)
            prsph = Cells(row, 7)
            prht = Cells(row, 8)
            Worksheets("C Change").Activate
            Cells(23, 1) = 4
            If age <> 0 Then Cells(23, 2) = age
            If prsph <> 0 Then Cells(23, 3) = prsph
            If prht <> 0 Then Cells(23, 4) = prht
        End If
        If recordtype = "P5" Then
            age = Cells(row, 3)
            prsph = Cells(row, 7)
            prht = Cells(row, 8)
            Worksheets("C Change").Activate
            Cells(24, 1) = 5
            If age <> 0 Then Cells(24, 2) = age
            If prsph <> 0 Then Cells(24, 3) = prsph
            If prht <> 0 Then Cells(24, 4) = prht
        End If
        
    Next row
    
    Worksheets("C_Change control").Activate
    Application.ScreenUpdating = True
    Application.StatusBar = False   'gives control of the statusbar back to the programme

End Sub

Sub transfer_CC_500I()
'This suboutine transfers information from "C Change" worksheet to "500 Index" worksheet
    Dim latitude As Double, altitude As Double, calage As Double, calsph As Double, calBA As Double, _
        calMTH As Double, I500 As Double, SI As Double, SoilC As Double, SoilN As Double, Temp As Double, _
        GeneticAdj As Double, CoreDens As Double, CoreAge As Long, InnerRing As Long, _
        OuterRing As Long, densitymodel As Long, j As Long, lastthinage As Double, _
        drift As Double, mortadd As Double, mortmult As Double
    
    Worksheets("C Change").Activate
    Initialsph = Cells(4, 3) * Cells(39, 3) / 100
    I500 = Cells(41, 3)
    SI = Cells(42, 3)
    latitude = Cells(33, 3)
    altitude = Cells(34, 3)
    rotlth1 = Cells(5, 3)
    rotlth2 = Cells(6, 3)
    If rotlth2 > rotlth1 Then
        rotlth = rotlth2
    Else
        rotlth = rotlth1
    End If
    calage = Cells(27, 3)
    calsph = Cells(28, 3)
    calBA = Cells(29, 3)
    calMTH = Cells(30, 3)
    SoilC = Cells(35, 3)
    SoilN = Cells(36, 3)
    Temp = Cells(40, 3)
    CoreDens = Cells(45, 3)
    CoreAge = Cells(46, 3)
    InnerRing = Cells(47, 3)
    OuterRing = Cells(48, 3)
    
    drift = Cells(51, 3)
    mortadd = Cells(52, 3)
    mortmult = Cells(53, 3)
    
    inthinrow = 11
    inprunerow = 20
    inyieldrow = 6
    
    Worksheets("500 Index").Activate
    Worksheets("500 Index").Cells(17, 10) = Worksheets("C Change").Cells(60, 3)  'Mode
    Range(Cells(2, 2), Cells(5, 2)).ClearContents
    Range(Cells(2, 4), Cells(5, 4)).ClearContents
    Range(Cells(2, 6), Cells(5, 6)).ClearContents
    Range(Cells(8, 3), Cells(10, 7)).ClearContents
    Range(Cells(14, 3), Cells(15, 7)).ClearContents
    Range(Cells(5, 10), Cells(10, 10)).ClearContents
    Range(Cells(5, 13), Cells(10, 13)).ClearContents
    
    If latitude <> 0 Then Cells(2, 4) = latitude
    If I500 <> 0 Then Cells(2, 2) = I500
    If SI <> 0 Then Cells(3, 2) = SI
    If SoilC <> 0 And SoilN <> 0 Then Cells(4, 4) = SoilC / (SoilN - 0.014)
    If Temp <> 0 Then Cells(3, 4) = Temp
    
    If CoreDens <> 0 Then Cells(5, 13) = CoreDens
    If CoreAge <> 0 Then Cells(6, 13) = CoreAge
    If InnerRing <> 0 Then Cells(7, 13) = InnerRing
    If OuterRing <> 0 Then Cells(8, 13) = OuterRing
            
    If mortmult <> 0 Then
        Cells(4, 2) = mortmult / 100
    Else
        Cells(4, 2).ClearContents
    End If
   
    Cells(3, 6) = rotlth
    Cells(2, 6) = Initialsph
    If calage <> 0 Then
        Cells(5, 10) = calage
        Cells(9, 10) = calsph
        Cells(6, 10) = calBA
        Cells(7, 10) = calMTH
    End If
    
    j = 0
    lastthinage = 0
    For i = 1 To 5
        Worksheets("C Change").Activate
        thinage = Cells(inthinrow - 1 + i, 2)
        thinsph = Cells(inthinrow - 1 + i, 3)
        Worksheets("500 Index").Activate
        If thinage = 0 Then Exit For
        Cells(14, 3 + j) = thinage
        Cells(15, 3 + j) = thinsph
        lastthinage = thinage
        j = j + 1
    Next i
    
    lastpruneage = 0
    j = 0
    For i = 1 To 5
        Worksheets("C Change").Activate
        pruneage = Cells(inprunerow - 1 + i, 2)
        prunesph = Cells(inprunerow - 1 + i, 3)
        pruneht = Cells(inprunerow - 1 + i, 4)
        If pruneage <> 0 Then
            Worksheets("500 Index").Activate
            If pruneage > lastpruneage Then j = j + 1
            Cells(8, 2 + j) = pruneage
            Cells(10, 2 + j) = prunesph
            Cells(9, 2 + j) = pruneht
            lastpruneage = pruneage
        End If
    Next i
End Sub

Sub transfer_500I_CC()
    'Copy predicted yield table from "500 Index" worksheet into "C Change" worksheet
    
    Dim i As Long, thinage As Double, earlyvol As Double
    Worksheets("C Change").Activate
  
    Range(Cells(6, 7), Cells(1000, 28)).ClearContents
    
    'Transfer age
    Worksheets("500 Index").Activate
    Range(Cells(23, 30), Cells(23 + rotlth, 30)).Copy
    Worksheets("C Change").Activate
    Cells(6, 7).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    'Transfer stocking
    Worksheets("500 Index").Activate
    Range(Cells(23, 31), Cells(23 + rotlth, 31)).Copy
    Worksheets("C Change").Activate
    Cells(6, 8).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    'Transfer volume
    Worksheets("500 Index").Activate
    Range(Cells(23, 34), Cells(23 + rotlth, 34)).Copy
    Worksheets("C Change").Activate
    Cells(6, 11).Select
    Selection.PasteSpecial Paste:=xlPasteValues
        
    'Transfer height
    Worksheets("500 Index").Activate
    If LCase(Cells(54, 3)) = "m" Then
        Range(Cells(23, 33), Cells(23 + rotlth, 33)).Copy
    Else
        Range(Cells(23, 32), Cells(23 + rotlth, 32)).Copy
    End If
    Worksheets("C Change").Activate
    Cells(6, 10).Select
    Selection.PasteSpecial Paste:=xlPasteValues
 
    'Transfer post-thin stocking and volume
    For i = 1 To 5
        thinage = Worksheets("500 Index").Cells(3 + i, 30)
        If thinage <> 0 Then
            Worksheets("C Change").Cells(thinage + 6, 9) = Worksheets("500 Index").Cells(3 + i, 31)
            Worksheets("C Change").Cells(thinage + 6, 12) = Worksheets("500 Index").Cells(3 + i, 32)
        End If
    Next i
     
    'Transfer sheath density
    Worksheets("500 Index").Activate
    Range(Cells(23, 36), Cells(23 + rotlth, 36)).Copy
    Worksheets("C Change").Activate
    Cells(6, 14).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    'Estimate volume lost to mortality
    Call Mortality_Volume

End Sub

Public Sub Mortality_Volume()
' This subroutine estimates the volume lost to mortality within each growth increment
    Dim numrows As Long, i As Long, sph1 As Double, vol1 As Double, sph2 As Double, vol2 As Double
    Worksheets("C Change").Activate
    numrows = Range(Cells(6, 7), Cells(6, 7).End(xlDown)).Rows.Count
    Cells(6, 13) = 0
    For i = 7 To numrows + 5
        sph1 = Cells(i - 1, 9)
        vol1 = Cells(i - 1, 12)
        If sph1 = 0 Then
            sph1 = Cells(i - 1, 8)
            vol1 = Cells(i - 1, 11)
        End If
        sph2 = Cells(i, 8)
        vol2 = Cells(i, 11)
        Cells(i, 13) = mort_vol(sph1, vol1, sph2, vol2)
    Next i
End Sub
