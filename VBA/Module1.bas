Attribute VB_Name = "Module1"
'*************************************************************************************************************
'
' This module contains the 300 Index Growth Model and a radiata pine wood density model.
'   Note that the controlling code for the Forest Carbon Predictor is in Module 7
'
'*************************************************************************************************************

'*****************************************************************************************************
'       300 Index Growth Model RADNAT version 1.05
'
'       This is an empirical stand-level growth model for radiata pine in New Zealand.
'
'
'       The original model, Version 1.01, was developed by Mark Kimberley in
'       February 2004.
'
'
'       Version 1.02 of the model was produced in September 2005 and incorporated
'       into the Radiata Pine Calculator Version 2. It had the following changes:
'
'       1. The function 'agezero' was changed. Agezero is a global variable
'           representing the age when MTH equals 1.4 metres. It is used in several
'           procedures. Version 1.01 used:
'               agezero = 8.6877 * Exp(-0.0539 * SI)
'           Version 1.02 and later versions use:
'               agezero = -Log(-(1 - Exp(-ha * 20)) * ((1.4 - 0.25) / (SI - 0.25)) ^
'                   (1 / hb) + 1) / ha
'           This is coded in the Function Calcagezero
'
'       2. The parameter db2 was set at 0.93 in Version 1.01. This was incorrect. It
'           was intended for use in a regional implementation of the model. However,
'           in the national version, db2 should equal one. This error was corrected
'           in Version 1.02.
'
'       3. In the function 'DBHmodel', stocking is limited to a maximum of 2000
'           stems/ha. This change was necessary because at extremely high stockings,
'           predicted growth of young trees was too low. This change only affects
'           predictions in stands growing at stockings greater than 2000 stems/ha.
'
'
'       Version 1.03, was produced in March 2006. It has considerable code
'       rearrangement and refinement to improve the efficiency of the model.
'       It includes the facility to run the carbon sequestration programme
'       C-Change. A wood density model is also added to this version. However, none
'       of the underlying growth model functions have changed although new volume and
'       mortality functions are added to the original lists of functions.
'
'       The major changes in Version 1.03 are as follows:
'
'       1. All code for the 300 Index model is in Module 1 along with the new wood
'           density model.
'
'       2. The code for running C-Change is in Module 2.
'
'       3. Input checking routines are now in Module 3.
'
'       4. All variables are declared explicitly in required modules and there have
'           been changes in the naming of some procedures.
'
'       5. The bisection method has been included as a separate function rather than
'           being embedded in individual procedures.
'
'       6. There is a new volume function. The original 2004 functions can still be
'           accessed in the spreadsheet by putting an 'x' beside the 'NZ 3-point taper
'           function'. However, the new function seems to be more stable and would
'           generally be preferred.
'
'       7. For very young trees, no conventional volume function provides good
'           predictions. To correct this, a new procedure 'Earlyield' is used to
'           predict volume when mean DBH is less than 2 cm.
'
'       8. There is a new and improved mortality function. The original 2004 function
'           can still be accessed by putting an 'x' beside the 'NZ 2004 model'. The new
'           model has the facility for the user to specify a base attritional mortality
'           level as a percentage.
'
'
'       Version 1.04 was programmed in March 2007 and includes three new features.
'
'       1. The 300 Index can be varied during a run. This enables a 300 Index drift
'           correction function to be applied, enabling regional variants of the model.
'
'       2. The mortality function developed for the P. M. Coop ias included. This is
'           an updated version of the Version 1.03 mortality function.
'
'
'       Version 1.05 was programmed in February 2008 and includes a number of changes
'
'       1. The Kimberley & Beets volume function (N.Z.J.For.Sci 37, 355-371 is included.
'           This is of the same form as the Version 1.03 function, but has updated coefficients.
'
'       2. Modifications to the model to eliminate bias in stands older than 30 years, in
'           stands less than 6 years old, and in stands with very high or low Site Index
'            as described in PMC Report No. 107.
'
'       3. The 300 Index can be varied during a run using a drift parameter (in m3/ha/yr2)
'           This enables a 300 Index drift correction function to be applied, enabling regional
'           variants of the model as described in PMC Report No. 107.
'
'       4. The mortality function developed in May 2007 (PMC Report No. ?) is included.
'           This is an updated version of the Version 1.03 mortality function. The user
'           can vary the standard predicted mortality using two parameters:
'           (i) an added % mortality rate which is added to the model prediction and
'              can be used to account for extra mortality not included in the model,
'               eg, windthrow.
'           (ii) a multiplicative % adjustment to account for regions with above or below
'               average mortality.
'
'       5. A minor problem with the 300 Index Model is that at extreme
'           stockings (>2000sph) and high Site Indices (>35m), the model predicts negligible
'           growth at mid-rotation ages, and can even produce negative DBH increments. A
'           simple fix eliminating negative increments was included. However, the model may
'           need some adjustment in the future to fully fix this problem.
'
'       6. The Tree List procedure developed in November 2007 (PMC Report No. 110)
'           was included (Subroutines Update_treelist and Run_treelist). Note that other
'           components of this model are in Module 3.
'
'       7. There was some minor rewriting of code. Global variables are now removed - all variables
'           except for the Implementation parameter are now local within each module.
'
'       8. The above changes refer to the code in this module (The basic single-plot 300 Index Model).
'           Other modules in this release have also been added or extensively modified.
'           These include Modules 5, 6 and 7 which contain the C_Change model.
'
'       In August 2010, this module was updated to include a new wood density model described in
'       the report for MFE: Kimberley and Beets April 2010 "A new model for predicting breast height
'       wood density by ring number in radiata pine". Also, the wood density function has been modified
'       by setting a minimum ring width limit of 1.5 mm.
'
'       Version 1.06 was programmed in June 2011. It includes modifications to the 300Index model
'       which include corrections for old ages and high stockings.
'
'       Version 1.07 was programmed in May 2012. It uses a delta of 0.0001 when checking for new prune
'           elements at each pruning lift in the subroutine Newlift to ensure that "phantom" pruned
'           elements are not retained
'
'       Version 1.08 was programmed in May 2012. The old age modifications introduced in Version 1.06
'           caused a kink in the BA growth curve. Version 1.08 corrects this using cubic interpolation
'           to predicr DBH between ages 20 and 40. This fix is contained in the new function DBHCalc.
'           There is also a modification to the way the drift parameter is appleid in function A200Start.
'           Drift adjustment now ceases at age 30 years (previously it ceased at age 45 years).
'
'       Version 1.09 was programmed in June 2014. Six new lines of code were included in the Thinning
'           subroutine to correct an error in the calculation of the thinning age shift which occured
'           when a late thinning was applied with a non-zero drift factor. Also the allowed range for
'           latitude was changed to 30-48.
'
'*****************************************************************************************************

Option Explicit     'Use Option Explicit except in calculator version

' Declare variables used in this project
Public implementation As Long, SI As Double, I300 As Double, ha As Double, hb As Double, D200 As Double, A200 As Double, N_MaxBA As Double, site_effect As Double, _
    DBHsqd_add_offset As Double, DBHsqd_mult_offset As Double, MTH_add_offset As Double, MTH_mult_offset As Double, _
    DBH_calibration_age As Double, MTH_calibration_age As Double

' Declare variables used in this module
Private adjage As Double, yieldline As Long, predicted As Boolean, agezero As Double, rotlth As Double, _
    startplotrow As Long, plotI300 As Double, SoilC As Double, SoilN As Double, Temp As Double, _
    GeneticAdj As Double, CoreDens As Double, CoreAge As Long, InnerRing As Long, OuterRing As Long, _
    drift As Double, sum_ba As Double, sum_height As Double, tl_prev_standN As Double, _
    tl_prev_standBA As Double, tl_prev_standmnheight As Double, tl_BA(10000) As Double, tl_standN As Double, _
    tl_standBA As Double, tl_standmnheight As Double, tl_predage As Double, tl_standage As Double, _
    tl_prev_standage As Double, tl_ntrees As Long, tl_survivalprob(10000), tl_prev_standDBH As Double, _
    tl_standDBH As Double, tl_dbh(10000) As Double, tl_height(10000) As Double, _
    tl_age As Double, tl_N As Double, tl_mnheight As Double, tl_mnba As Double, tl_MTH As Double, _
    tl_totBA As Double, N As Double, Thin As Long, Nshist As Long, outputline As Long, mth As Double, _
    Meanht(10) As Double, prht(10) As Double, prlag(10) As Double, ncum(10) As Double, maxage As Double, _
    Nlifts As Long, Nelements As Long, Nthins As Long, vol As Double, ThinLag(8) As Double, _
    sellag(10) As Double, adjageel(10) As Double, thinage As Boolean, nelement(10) As Double, age As Double, _
    dbh As Double, dbhsqd As Double, dbhelement(10) As Double, agethin(8) As Double, initiallag(8) As Double, _
    totalthinlag As Double, BA As Double, crlth As Double, shist_T(999) As Double, shist_N1(999) As Double, _
    shist_N2(999) As Double, shist_thincoeff(999) As Double, Initialstocking As Double, steplength As Double, _
    steps As Long, lift_T(999) As Double, lift_height(999) As Double, Npruned As Double, _
    Mortality(999) As Double, OUTPUT As Boolean, shist As Long, j As Long, el As Long, lineprinted As Boolean, _
    lift As Long, prevN As Double, total_prlag(10) As Double, prevprht As Double, _
    voltabarray(11) As String, DBH300 As Double, shist_thinratio(999) As Double, age_300Index, _
    agediff As Double, prevagediff As Double, lift_sph(10) As Double, v(11, 8) As Double, voltable As Long, _
    attrition As Double, pctmortadj As Double, X As Double, sdi As Double, lift_prunecoeff(10) As Double, _
    heightmodel As Long, mortmodel As Long, mnheight As Double, specage As Double, predage As Double, _
    altitude As Double, latitude As Double, indexage As Double, bias_young As Boolean, bias_old As Boolean, _
    bias_SI As Boolean
                                        
' Model coefficients used by Module 1 procedures

' Coefficients for February 2004 model
Private Const da1 As Double = 56.523, db1 As Double = -0.09045, dr As Double = 2.6416, _
    dl As Double = 28.1224, dc As Double = 1.4821, dbSI As Double = -0.00212, _
    dn As Double = 15.7581, dm As Double = -0.00455, dbdia As Double = -0.1325, _
    Ds As Double = 0.1702, dbsidia As Double = -0.0084, drsi As Double = 0.0209, _
    dr2 As Double = 0.8234, _
    pra As Double = 0.0934, prb As Double = 1.98, prc As Double = 0.2119, _
    ha0 As Double = -2.475, ha1 As Double = -0.01406, hb0 As Double = 0.33417, _
    hb1 As Double = 0.0104, hae0 As Double = -1.335, hae1 As Double = -0.03581, _
    hae2 As Double = -0.0006306, hbe0 As Double = 0.499, hbe1 As Double = 0.005059, _
    hNSWa As Double = -2.6842, hNSWb As Double = 0.7293, hNSWp As Double = -0.00176, _
    thb As Double = 0.5, thc As Double = -0.47, mortd As Double = 0.2493, _
    tha As Double = 0.5, db2 As Double = 1, thincoeff As Double = 0.784, _
    mortc As Double = 1.5, morte As Double = -0.0555, mortNSW = 0.869, _
    morta As Double = 0.000688, mortb As Double = -14.91, _
    mortp As Double = -44.691, mortq As Double = -4.611, _
    morts As Double = 3.901, mortt As Double = 1.3533, mortu As Double = 0.00246, _
    mortv As Double = -30.565, mortw As Double = 2.536, mortx As Double = 1.125, _
    morty As Double = 0.000438, morta1 As Double = 0.00206, mortb1 As Double = -46.3216, _
    mortc1 As Double = 3.1704, mortd1 As Double = 1.7477, morte1 As Double = -0.1631, _
    mortf1 As Double = 0.1991, _
    mort2007_a As Double = 0.000459, mort2007_b As Double = 0.974, mort2007_c As Double = 3.06, _
    mort2007_d As Double = 0.786, mort2007_f As Double = -0.037, mort2007_g As Double = 0.0371, _
    mort2007_h As Double = -0.32

Sub voltab()
    'Assign values to volume table coefficients
    v(1, 1) = 0.942            'Kimberley & Beets, 2007
    v(1, 2) = -1.161
    v(1, 3) = 0.317
    v(2, 1) = 0.989            'Kimberley, 2006
    v(2, 2) = -1.2752
    v(2, 3) = 0.3191
    v(3, 1) = 1.492912924           'Volume function 182
    v(3, 2) = -0.999113309
    v(3, 3) = 1.250753941
    v(3, 4) = -0.397037159
    v(3, 5) = 0.027218164
    v(3, 6) = -0.063166205
    v(3, 7) = 0.064609459
    v(3, 8) = -0.030665365
    v(4, 1) = 1.633105986           'Volume function 236
    v(4, 2) = -1.039327204
    v(4, 3) = 1.212696953
    v(4, 4) = -0.359131176
    v(4, 5) = 0.026454943
    v(4, 6) = -0.067457458
    v(4, 7) = 0.066992488
    v(4, 8) = -0.030528278
    v(5, 1) = 0.730448717           'Volume function 328
    v(5, 2) = -0.617440226
    v(5, 3) = 1.095616037
    v(5, 4) = -0.222220223
    v(5, 5) = 0.013858949
    v(5, 6) = -0.11022445
    v(5, 7) = 0.059157535
    v(5, 8) = -0.016942593
    v(6, 1) = 1.09857999           'Volume function 358
    v(6, 2) = -0.883862258
    v(6, 3) = 1.165375013
    v(6, 4) = -0.28047221
    v(6, 5) = 0.022081234
    v(6, 6) = -0.059261776
    v(6, 7) = 0.053187392
    v(6, 8) = -0.025226521
    v(7, 1) = 1.403009551           'Volume function 11
    v(7, 2) = -0.96392392
    v(7, 3) = 1.221046594
    v(7, 4) = -0.358337009
    v(7, 5) = 0.024975712
    v(7, 6) = -0.061374804
    v(7, 7) = 0.061895757
    v(7, 8) = -0.028672533
    v(8, 1) = 2.834246614           'Volume function 430
    v(8, 2) = -1.856804825
    v(8, 3) = 1.152097786
    v(8, 4) = -0.201346156
    v(8, 5) = -0.000721117
    v(8, 6) = 0.081503044
    v(8, 7) = 0.024428222
    v(8, 8) = 0.001938887
    v(9, 1) = 2.7023           'Volume function 3-point-taper
    v(9, 2) = -2.1301
    v(9, 3) = 1.3901
    v(9, 4) = -0.5056
    v(9, 5) = 0.0548
    v(9, 6) = 0.0991
    v(9, 7) = 0.1478
    v(9, 8) = -0.088
    v(10, 1) = 6.2733           'Volume function NSW1
    v(10, 2) = 0.1284
    v(10, 3) = -0.00097
    v(11, 1) = 2.1819           'Volume function NSW2
    v(11, 2) = 0.2504
    v(11, 3) = -0.00081
End Sub

Sub Inputparms()

' Input details of stand
' Macro written 5 Jan 2004 by Mark Kimberley
' Produces maxage, steplength, nsteps, SI, I300, Initialstocking, shist_T(999), shist_N(999), shist_pctmech(999), Nshist
    Dim i As Long, startline As Long, nlines As Long
    
    implementation = Cells(8, 6)    'Operating mode: 1=Standard mode, 2=Offset mode, 3=Index mode

    I300 = Cells(3, 3)
    SI = Cells(4, 3)
    Initialstocking = Cells(19, 3)
    drift = Cells(64, 6)
    
    ' Determine bias corrections to be used in BA model
    bias_old = False
    bias_young = False
    bias_SI = False
    If LCase(Cells(51, 6)) = "x" Then bias_old = True
    If LCase(Cells(52, 6)) = "x" Then bias_young = True
    If LCase(Cells(53, 6)) = "x" Then bias_SI = True
    
    maxage = Cells(47, 3)
    steplength = Cells(48, 3)
    If steplength < 0.01 Then steplength = 0.01
    
' Determine height model - 1 = NSW, 2 = Simple NZ, 3 = Environmental NZ

    heightmodel = heightmod()
    
' Determine mortality model - 1 = NSW, 2 = NZ, 3 = new NZ

    If LCase(Cells(68, 4)) = "x" Then
        mortmodel = 1
    ElseIf LCase(Cells(69, 4)) = "x" Then
        mortmodel = 2
    ElseIf LCase(Cells(70, 4)) = "x" Then
        mortmodel = 3
    ElseIf LCase(Cells(71, 4)) = "x" Then
        mortmodel = 5
    Else
        mortmodel = 6
    End If
        
    If mortmodel >= 4 Then
        If Cells(68, 6) <> 0 Then
            attrition = Cells(68, 6) / 100  'User input attritional mortality for new NZ mortality model
        ElseIf mortmodel = 4 Then
            attrition = mortu 'Default attritional mortality for interim 2006 NZ mortality model
        ElseIf mortmodel = 5 Then
            attrition = morta1 'Default attritional mortality for interim 2006 NZ mortality model
        Else
            attrition = 0 'Default attritional mortality for 2007 NZ mortality model
        End If
        If Cells(69, 6) <> 0 Then
            pctmortadj = Cells(69, 6) 'User input % adjustment for 2007 NZ mortality model
        Else
            pctmortadj = 0 'Default attritional mortality for 2007 NZ mortality model
        End If
    End If
        
' Read volume table array

    For i = 1 To 11
        voltabarray(i) = Cells(50 + i, 4)
    Next i
    
' Obtain volume table number from volume table array
    If LCase(voltabarray(1)) = "x" Then
        voltable = 1
    ElseIf LCase(voltabarray(2)) = "x" Then
        voltable = 2
    ElseIf LCase(voltabarray(3)) = "x" Then
        voltable = 3
    ElseIf LCase(voltabarray(4)) = "x" Then
        voltable = 4
    ElseIf LCase(voltabarray(5)) = "x" Then
        voltable = 5
    ElseIf LCase(voltabarray(6)) = "x" Then
        voltable = 6
    ElseIf LCase(voltabarray(7)) = "x" Then
        voltable = 7
    ElseIf LCase(voltabarray(8)) = "x" Then
        voltable = 8
    ElseIf LCase(voltabarray(9)) = "x" Then
        voltable = 9
    ElseIf LCase(voltabarray(10)) = "x" Then
        voltable = 10
    ElseIf LCase(voltabarray(11)) = "x" Then
        voltable = 11
    End If
    
  
    steps = maxage / steplength
    Call calcheightcoeff(SI)

' Zeroize stocking & pruning history arrays
    For shist = 1 To 17
        shist_T(shist) = 0
        shist_N1(shist) = 0
        shist_N2(shist) = 0
        shist_thinratio(shist) = 0 'zero if pre/post DBH ratio must be calculated
        shist_thincoeff(shist) = 0
        Mortality(shist) = 0
    Next shist

    For lift = 1 To 5
        lift_T(lift) = 0
        lift_height(lift) = 0
        lift_sph(lift) = 0
        lift_prunecoeff(lift) = 0
    Next lift
        
' Read stocking & thinning history
    startline = 19
    nlines = 17
    
    Nshist = 0
    For shist = 1 To nlines
        If Not IsEmpty(Cells(startline + shist, 2)) Then
            shist_T(shist) = Cells(startline + shist, 2)
            shist_N1(shist) = Cells(startline + shist, 3)
            shist_N2(shist) = Cells(startline + shist, 4)
            shist_thincoeff(shist) = Cells(startline + shist, 5)
            shist_thinratio(shist) = Cells(startline + shist, 6)
            Nshist = shist
        End If
    Next shist
    
    ' Calculate mortalities
        Call mort
        
    ' Read pruning history
    startline = 39

    Nlifts = 0
    For lift = 1 To 5
        If Not IsEmpty(Cells(startline + lift, 2)) Then
            lift_T(lift) = Cells(startline + lift, 2)
            lift_height(lift) = Cells(startline + lift, 3)
            lift_sph(lift) = Cells(startline + lift, 4)
            If lift_sph(lift) = 0 Then lift_sph(lift) = 10000
                'If pruned stocking not specified, assume all trees are pruned
            lift_prunecoeff(lift) = Cells(startline + lift, 5)
            Nlifts = lift
        End If
    Next lift
    
End Sub

Sub OutputGrowth()
Attribute OutputGrowth.VB_ProcData.VB_Invoke_Func = "a\n14"
' Input I300, SI & stand history and predict growth

    Dim outputrange As Object
    If checkinput_site = False Then Exit Sub
    If checkinput_SI = False Then Exit Sub
    If checkinput_htfn = False Then Exit Sub
    If checkinput_initialstock = False Then Exit Sub
    If checkinput_stocking = False Then Exit Sub
    If checkinput_prune = False Then Exit Sub
    If checkinput_fellage = False Then Exit Sub
    If checkinput_steplth = False Then Exit Sub
    If checkinput_volfn = False Then Exit Sub
    If checkinput_mortfn = False Then Exit Sub
    OUTPUT = True
    Worksheets("300 Index").Activate
    Set outputrange = Range(Cells(5, 7), Cells(150, 70))
    outputrange.ClearContents
    Call Inputparms
    Call voltab
    If implementation = 2 Then
        Call CalcOffsets
        Call Inputparms
        OUTPUT = True
    End If
    Call Growth(OUTPUT)
    Call earlyield
    Call mortvol
    Call density
End Sub

Sub Calc300Index()
Attribute Calc300Index.VB_ProcData.VB_Invoke_Func = "i\n14"
' Input plot measurement data, SI, and stand history, calculate 300 Index and predict growth
    Dim BA300 As Double, age300 As Double, Stock300 As Double, MTH300 As Double, _
        Vol300 As Double
    
    Worksheets("300 Index").Activate
    If checkinput_I300 = False Then Exit Sub
    If checkinput_SI = False Then Exit Sub
    If checkinput_htfn = False Then Exit Sub
    If checkinput_initialstock = False Then Exit Sub
    If checkinput_stocking = False Then Exit Sub
    If checkinput_prune = False Then Exit Sub
    If checkinput_fellage = False Then Exit Sub
    If checkinput_steplth = False Then Exit Sub
    If checkinput_volfn = False Then Exit Sub
    If checkinput_mortfn = False Then Exit Sub
    Call Inputparms
    Call voltab
    age300 = Cells(7, 3)
    Stock300 = Cells(8, 3)
    maxage = age300
    steps = maxage / steplength
    MTH300 = CalcMTH(SI, age300)
    DBH300 = Cells(9, 3)
    If DBH300 = 0 Then
        BA300 = Cells(10, 3)
        If BA300 <> 0 Then
            DBH300 = CalcDBHfromBA(BA300, Stock300)
        Else
            Vol300 = Cells(11, 3)
            DBH300 = CalcDBHfromBA(calcBAfromVol(MTH300, Vol300, Stock300), Stock300)
        End If
    End If
    Call Index300
    Cells(3, 3) = I300
    Call OutputGrowth
    
End Sub

Sub Index300()
' Calculate 300 Index using Bisection Method
    Dim nsteps As Long, j1 As Long
    'Use 1.328 and 60 as lower and upper limits for calculating 300 Index using bisection method
    I300 = Bisection(1.328, 60, 14, 1, 0, 0, 0, 0)
End Sub

Sub Growth(OUTPUT)
' Predict growth from given stand parameters. Output predictions if required

    N = Initialstocking         'N contains stocking before thinning
    shist = 1
    Thin = 0
    lift = 0
    age = 0
    Nelements = 1
    A200 = CalcA200start(age, I300, SI)
    dbh = 0
    'The following loop was included in October 2007 and is necessary to allow for
    '   checking for negative diameter increments in subroutine Diameter
    For el = 1 To 10
        dbhelement(el) = 0
    Next el
    vol = 0.0000064 * N     'Volume of seedling at planting (Beets)
    BA = 0
    mth = 0.25
    mnheight = 0.25
    nelement(1) = N
    ncum(1) = N
    prht(1) = 0
    prlag(1) = 0
    totalthinlag = 0
    sellag(1) = 0
    total_prlag(1) = 0
    outputline = 5
    lineprinted = False
    If implementation <> 2 Then 'If implementation does not require offsets then set them to neutral values
        DBHsqd_add_offset = 0
        DBHsqd_mult_offset = 1
        MTH_add_offset = 0
        MTH_mult_offset = 1
        DBH_calibration_age = 0
        MTH_calibration_age = 0
    End If
    If OUTPUT Then Call OutStep 'Print age zero
    
    For j = 1 To steps
        tl_prev_standDBH = dbh  'Stand parameters at previous step used for tree list projection
        tl_prev_standN = N
        tl_prev_standBA = BA
        tl_prev_standmnheight = mnheight
        tl_prev_standage = age
        age = age + steplength
        A200 = CalcA200start(age, I300, SI)
        Call stock
        Call Height
        Call Ageshifts
        Call Diameter
        Call VolBA
        If OUTPUT And (age - Int(age) < 0.001 Or age - Int(age) > 0.999) _
            Or Abs(age - maxage) < 0.001 Then Call OutStep
        If shist_N2(shist) <> 0 And age >= shist_T(shist) - 0.001 Then
            If OUTPUT And Not lineprinted Then Call OutStep
            Thin = Thin + 1
            Call thinning
            If OUTPUT Then
                Call OutThin
                Call OutElements
            End If
        End If
        If lift < Nlifts And age >= lift_T(lift + 1) - 0.001 Then
            If OUTPUT And Not lineprinted Then Call OutStep
            lift = lift + 1
            Call Newlift
            If OUTPUT Then
                Call OutPrune
                Call OutElements
            End If
        End If
        If shist < Nshist And age >= shist_T(shist) - 0.001 Then shist = shist + 1
        If lineprinted Then
            outputline = outputline + 1
            lineprinted = False
        End If
    Next j
    
End Sub
        
Sub OutStep()
' Output yield for a single prediction iteration
    Dim offset_corrected_dbh As Double, offset_corrected_MTH As Double
    Worksheets("300 Index").Activate
    Cells(outputline, 7) = WorksheetFunction.Round(age, 2)
    Cells(outputline, 8) = N
    
    If age < MTH_calibration_age Then
        offset_corrected_MTH = mth * MTH_mult_offset
    Else
        offset_corrected_MTH = mth + MTH_add_offset
    End If
    If age < DBH_calibration_age Then
        offset_corrected_dbh = Sqr(dbh ^ 2 * DBHsqd_mult_offset)
    Else
        offset_corrected_dbh = Sqr(dbh ^ 2 + DBHsqd_add_offset)
    End If
    
    Cells(outputline, 10) = offset_corrected_MTH
    Cells(outputline, 16) = offset_corrected_dbh
    Cells(outputline, 14) = CalcBAfromDBH(offset_corrected_dbh, N)
    Cells(outputline, 12) = CalcVol(offset_corrected_MTH, CalcBAfromDBH(offset_corrected_dbh, N), N)
    Cells(outputline, 18) = calcMeanht(offset_corrected_MTH, N)
    
    Call OutElements
    Cells(outputline, 43) = DBHmodel(A200, SI, 20, N)
        'Calculate DBH at age 20 for use by branch model
    lineprinted = True

End Sub

Sub OutElements()
' Output stocking, mean DBH and volume of each pruned element
    Range(Cells(outputline, 19), Cells(outputline, 36)).ClearContents
    For el = 1 To Nelements
        Cells(outputline, 16 + el * 3) = nelement(el)
        Cells(outputline, 16 + el * 3 + 1) = dbhelement(el)
        Cells(outputline, 16 + el * 3 + 2) = _
            CalcVol(mth, CalcBAfromDBH(dbhelement(el), nelement(el)), nelement(el))
        ' Volume of each element is only approximate
    Next el
'    Cells(outputline, 42) = totalthinlag
'    For el = 1 To Nelements
'        Cells(outputline, 42 + el) = prlag(el)
'        Cells(outputline, 46 + el) = sellag(el)
'        Cells(outputline, 50 + el) = adjageel(el)
'    Next el
'    Cells(outputline, 55) = totalthinlag
'    Cells(outputline, 56) = adjage
End Sub
    

Sub OutThin()
' Output stand parameters following a thinning (Stocking, DBH, Volume, BA, and element data)
    Dim offset_corrected_dbh As Double, offset_corrected_MTH As Double
    Worksheets("300 Index").Activate
    If age < MTH_calibration_age Then
        offset_corrected_MTH = mth * MTH_mult_offset
    Else
        offset_corrected_MTH = mth + MTH_add_offset
    End If
    If age < DBH_calibration_age Then
        offset_corrected_dbh = Sqr(dbh ^ 2 * DBHsqd_mult_offset)
    Else
        offset_corrected_dbh = Sqr(dbh ^ 2 + DBHsqd_add_offset)
    End If
    
    Cells(outputline, 9) = N
    Cells(outputline, 17) = offset_corrected_dbh
    Cells(outputline, 15) = CalcBAfromDBH(offset_corrected_dbh, N)
    Cells(outputline, 13) = CalcVol(offset_corrected_MTH, CalcBAfromDBH(offset_corrected_dbh, N), N)
    Cells(outputline, 18) = calcMeanht(offset_corrected_MTH, N)
    For el = 1 To Nelements
        Cells(outputline, 16 + el * 3) = nelement(el)
        Cells(outputline, 16 + el * 3 + 1) = dbhelement(el)
        Cells(outputline, 16 + el * 3 + 2) = _
            CalcVol(mth, CalcBAfromDBH(dbhelement(el), nelement(el)), nelement(el))
    Next el
    Cells(outputline, 43) = DBHmodel(A200, SI, 20, N)
        'Calculate DBH at age 20 for use by branch model
    
End Sub

Sub OutPrune()
' Output crown length for a pruning lift
    Worksheets("300 Index").Activate
    Cells(outputline, 11) = crlth
    If lift_sph(lift) = 10000 Then Cells(outputline, 37) = N _
        Else Cells(outputline, 37) = lift_sph(lift)
    Cells(outputline, 38) = lift_height(lift)
    Cells(outputline, 43) = DBHmodel(A200, SI, 20, N)
        'Calculate DBH at age 20 for use by branch model
    
End Sub

Sub mort()

' Calculate mortalities
    Dim prevage As Double
    
    'mortrate = 0.57            'Old mortality rate coefficient (%mortality per year)

    If shist_T(Nshist) < maxage Then
        Nshist = Nshist + 1
        shist_T(Nshist) = maxage
    End If
    
    prevage = 0
    prevN = Initialstocking
    For shist = 1 To Nshist
        If IsEmpty(shist_N1(shist)) Or shist_N1(shist) = 0 Then
'            mortality(shist) = mortrate 'Old mortality model
            Mortality(shist) = -1   'Will need to calculate mortality rate at each step
        Else
            Mortality(shist) = 100 * Log(prevN / shist_N1(shist)) / _
                (shist_T(shist) - prevage)
        End If
        prevage = shist_T(shist)
        If IsEmpty(shist_N2(shist)) Or shist_N2(shist) = 0 Then
            prevN = shist_N1(shist)
        Else
            prevN = shist_N2(shist)
        End If
    Next shist

End Sub

Sub stock()

' Generate stocking using mortality function
' Macro written 7 Oct 2003 by Mark Kimberley
'   Produces N (total stocking), Nelement(el) (stocking in each element)
'   & Ncum(el) (cumulative stocking in each element)
'   Last updated June 2007

    Dim mortrate As Double, X As Double, el As Long
    
    prevN = N
    If Mortality(shist) >= 0 Then
        N = prevN / Exp(Mortality(shist) * steplength / 100)
    ElseIf mortmodel = 1 Then
        mortrate = mortNSW
        N = prevN / Exp(mortrate * steplength / 100)
    ElseIf mortmodel = 2 Then
        If dbh = 0 Then
            mortrate = 0
        Else
            X = Exp(mortb + morte * SI + mortc * (Log(N) + mortd * Log(dbh ^ 2)))
            mortrate = (morta + (1 - morta) * X / (1 + X)) * 100
            N = prevN / Exp(mortrate * steplength / 100)
        End If
    ElseIf mortmodel = 3 Then
        If dbh = 0 Then
            mortrate = 0
        Else
            X = Exp(mortv + mortw * (Log(N) + mortx * Log(dbh)))
            mortrate = (morty + (1 - morty) * X / (1 + X)) * 100
            N = prevN / Exp(mortrate * steplength / 100)
        End If
    ElseIf mortmodel = 4 Then
        If dbh = 0 Then
            mortrate = 0
        Else
            X = Exp(mortp + mortq * I300 / SI + morts * (Log(N) + mortt * Log(dbh)))
            mortrate = (attrition + (1 - attrition) * X / (1 + X)) * 100
            N = prevN / Exp(mortrate * steplength / 100)
        End If
    ElseIf mortmodel = 5 Then
        If dbh = 0 Then
            mortrate = 0
        Else
            X = Exp(mortb1 + morte1 * I300 + mortf1 * SI + mortc1 * (Log(N) + mortd1 * Log(dbh)))
            mortrate = (attrition + (1 - attrition) * X / (1 + X)) * 100
            N = prevN / Exp(mortrate * steplength / 100)
        End If
    ElseIf mortmodel = 6 Then
        If dbh = 0 Then
            mortrate = 0
        Else
            sdi = Exp(mort2007_f * I300 + mort2007_g * SI + Log(N) + mort2007_d * Log(dbh / 100) + _
                mort2007_h * (Log(dbh / 100)) ^ 2) / 1000
            mortrate = attrition * 100 + 100 * (1 + pctmortadj / 100) * (mort2007_a + mort2007_b * sdi ^ mort2007_c)
            If mortrate > 95 Then mortrate = 95
            If mortrate < 0 Then mortrate = 0
            N = prevN * (1 - mortrate / 100) ^ steplength
        End If
    End If
    For el = 1 To Nelements
        nelement(el) = nelement(el) * N / prevN
        ncum(el) = ncum(el) * N / prevN
                'Apply mortality evenly to each element
    Next el
    
End Sub

Sub Height()

' Predict MTH and Mean Height of each element for a given age, SI and stocking
' Macro written 7 Oct 2003 by Mark Kimberley
'   Produces MTH and Meanht

    mth = CalcMTH(SI, age)
    mnheight = calcMeanht(mth, N)
    Meanht(1) = calcMeanht(mth, nelement(1))
    For el = 2 To Nelements
        Meanht(el) = (ncum(el) * calcMeanht(mth, ncum(el)) _
            - ncum(el - 1) * calcMeanht(mth, ncum(el - 1))) / (ncum(el) - ncum(el - 1))
    Next el
    
End Sub

Sub Ageshifts()

' Calculate pruning and thinning time shifts for each element
' Macro written 24 February 2004 by Mark Kimberley
'   Produces prlag(el), totalthinlag, thinlag(th), adjage, adjageel(el)
    Dim th As Long, el As Long, timesincethin As Double
    
    For el = 1 To Nelements
        If prht(el) > 0 Then prlag(el) = prlag(el) + 0.3 * steplength
        If prlag(el) > total_prlag(el) Then prlag(el) = total_prlag(el)
    Next el
    
    totalthinlag = 0
    For th = 1 To Thin
        timesincethin = age - agethin(th)
        ThinLag(th) = initiallag(th) + _
            WorksheetFunction.Min(initiallag(th), tha) * thb * (1 - Exp(thc * timesincethin))
        totalthinlag = totalthinlag + ThinLag(th)
    Next th
    
    adjage = 0
    For el = 1 To Nelements
        adjageel(el) = age - prlag(el) - sellag(el) - totalthinlag
        adjage = adjage + adjageel(el) * nelement(el)
    Next el
    adjage = adjage / N

End Sub

Sub Newlift()
' Calculate element means following a pruning lift
    Dim dbhb4pr As Double, prunecoeff As Double
    
    If lift_sph(lift) + 0.0001 < nelement(1) Then
        Nelements = Nelements + 1       ' New element required
        For el = Nelements To 2 Step -1     'Move elements up one
            prht(el) = prht(el - 1)
            nelement(el) = nelement(el - 1)
            ncum(el) = ncum(el - 1)
            total_prlag(el) = total_prlag(el - 1)
            prlag(el) = prlag(el - 1)
            dbhelement(el) = dbhelement(el - 1)
            Meanht(el) = Meanht(el - 1)
            adjageel(el) = adjageel(el - 1)
            sellag(el) = sellag(el - 1)
        Next el
        prht(1) = lift_height(lift)
        nelement(1) = lift_sph(lift)
        nelement(2) = nelement(2) - nelement(1)
        ncum(1) = nelement(1)
        Meanht(1) = calcMeanht(mth, nelement(1))
        crlth = Meanht(1) - prht(1)
        dbhb4pr = dbhelement(1)
        'Obtain pruning selection coefficient
        If lift_prunecoeff(lift) <> 0 Then
            prunecoeff = lift_prunecoeff(lift)
        Else
            prunecoeff = thincoeff  'Use default thinning coefficient for pruning selection if not specified
        End If
        dbhelement(1) = dbhelement(1) * (nelement(1) / ncum(2)) ^ ((prunecoeff - 1) / 2)
        dbhelement(2) = Sqr((ncum(2) * dbhelement(2) ^ 2 - nelement(1) * dbhelement(1) ^ 2) _
            / nelement(2))
        sellag(1) = sellag(1) + adjageel(1) - CalcAge(dbhelement(1), A200, N, SI)
        sellag(2) = sellag(2) + adjageel(2) - CalcAge(dbhelement(2), A200, N, SI)
        adjageel(1) = age - prlag(1) - totalthinlag - sellag(1)
        adjageel(2) = age - prlag(2) - totalthinlag - sellag(2)
        
        total_prlag(1) = total_prlag(2) + _
            pra * (prht(1) ^ prb - prht(2) ^ prb) * Exp(-prc * crlth)
    Else            'No new element required - prune existing first element
        prevprht = prht(1)
        prht(1) = lift_height(lift)
        crlth = Meanht(1) - prht(1)
        total_prlag(1) = total_prlag(1) + _
            pra * (prht(1) ^ prb - prevprht ^ prb) * Exp(-prc * crlth)
    End If
    
End Sub

Sub Diameter()
' Macro written Jan 2004 by Mark Kimberley
' Calculate mean DBH of each element
    Dim prevdbh_el As Double
    N = ncum(Nelements)
    dbhsqd = 0
    For el = 1 To Nelements
        prevdbh_el = dbhelement(el)
'        dbhelement(el) = DBHmodel(A200, SI, adjageel(el), N)
        dbhelement(el) = CalcDBH(I300, SI, adjageel(el), N)
        'Change made in  October 2007 - check that DBH increment is not negative
        If dbhelement(el) < prevdbh_el Then dbhelement(el) = prevdbh_el
        dbhsqd = dbhsqd + nelement(el) * dbhelement(el) ^ 2
    Next el
    dbh = Sqr(dbhsqd / N)
End Sub
    
Sub thinning()
' Macro written Jan 2004 by Mark Kimberley
' Predict BA, DBH, Volume, and stocking following thinning, overall and for each pruned element after thinning
' Produces N, Nelement(el), DBH(el), ThinLag

    Dim kcoeff As Double, ratio1 As Double, ratio2 As Double, prevdbh As Double, _
        thinN As Double, el As Long, prevNel As Double, prevNcum, _
        current_thinratio As Double

    prevN = N
    prevdbh = dbh
    
    'Obtain thinning coefficient
    If shist_thincoeff(shist) <> 0 Then
        kcoeff = shist_thincoeff(shist)
    Else
        kcoeff = thincoeff  'Default thinning coefficient for waste thinning
    End If
        
    'Thinning from below
    thinN = prevN - shist_N2(shist)
    For el = Nelements To 1 Step -1
        If thinN + 0.0001 >= nelement(el) Then 'Add 0.0001 to eliminate bug
            thinN = thinN - nelement(el)
            Nelements = Nelements - 1   'Thin out entire element
        Else
            prevNel = nelement(el)
            prevNcum = ncum(el)
            nelement(el) = nelement(el) - thinN     'Thin element
            ncum(el) = ncum(el) - thinN
            If el <> 1 Then
                dbhelement(el) = dbhelement(el) * _
                    (ncum(el) ^ ((kcoeff + 1) / 2) - ncum(el - 1) ^ ((kcoeff + 1) / 2)) * (prevNcum - ncum(el - 1)) / _
                    ((prevNcum ^ ((kcoeff + 1) / 2) - ncum(el - 1) ^ ((kcoeff + 1) / 2)) * (ncum(el) - ncum(el - 1)))
            Else
                dbhelement(el) = dbhelement(el) * (ncum(el) / prevNcum) ^ ((kcoeff - 1) / 2)
            End If
            Nelements = el
            Exit For
        End If
    Next el
    
    'Calculate DBH after thinning
    N = ncum(Nelements)
    dbhsqd = 0
    For el = 1 To Nelements
        dbhsqd = dbhsqd + nelement(el) * dbhelement(el) ^ 2
    Next el
    dbh = Sqr(dbhsqd / N)
    
    'Specified diameter ratio
    If shist_thinratio(shist) <> 0 And prevdbh <> 0 Then
        current_thinratio = dbh / prevdbh
        For el = 1 To Nelements
            dbhelement(el) = dbhelement(el) * shist_thinratio(shist) / current_thinratio
        Next el
        dbh = dbh * shist_thinratio(shist) / current_thinratio
    End If
    
    'Calculate volume and BA after thinning
    Call VolBA
    
    'Calculate selection and thinning age shifts at time of thinning
    If prevdbh = 0 Then
        initiallag(Thin) = 0
    Else
        'The following 6 lines of code were added in June 2014
            'These correctd an error in the caclulation of the thinning lag which occured when a late thinnings was
            'applied with a non-zero drift factor
        A200 = CalcA200start(adjage, I300, SI)
        initiallag(Thin) = adjage - CalcAge(prevdbh, A200, N, SI)
        A200 = CalcA200start(adjage - initiallag(Thin), I300, SI)
        initiallag(Thin) = adjage - CalcAge(prevdbh, A200, N, SI)
        A200 = CalcA200start(adjage - initiallag(Thin), I300, SI)
        initiallag(Thin) = adjage - CalcAge(prevdbh, A200, N, SI)
    End If
    ThinLag(Thin) = initiallag(Thin)
    agethin(Thin) = age
    totalthinlag = totalthinlag + initiallag(Thin)
    For el = 1 To Nelements
        If dbhelement(el) = 0 Then
            sellag(el) = 0
        Else
            sellag(el) = age - prlag(el) - totalthinlag - CalcAge(dbhelement(el), A200, N, SI)
        End If
        adjageel(el) = age - prlag(el) - totalthinlag - sellag(el)
    Next el
    adjage = 0
    For el = 1 To Nelements
        adjage = adjage + adjageel(el) * nelement(el)
    Next el
    adjage = adjage / N

    'Calculate mean height after thinning
    mnheight = calcMeanht(mth, N)
    
End Sub

Sub VolBA()
' Calculate Volume and BA from DBH, MTH and stocking
    BA = CalcBAfromDBH(dbh, N)
    vol = CalcVol(mth, BA, N)
End Sub

Sub siteIndex()
Attribute siteIndex.VB_ProcData.VB_Invoke_Func = "j\n14"
' Use bisection method to calculate SI using Mina's model
    Dim HAge As Double, HMTH As Double

    If checkinput_htage = False Then Exit Sub
    If checkinput_htfn = False Then Exit Sub
       
    HAge = Cells(14, 3)
    HMTH = Cells(15, 3)
    If HAge = 20 Then
        SI = HMTH
    Else
        
        ' Determine height model - 1 = NSW, 2 = Simple NZ, 3 = Environmental NZ

        heightmodel = heightmod()
        SI = Bisection(5, 60, 15, 2, HMTH, HAge, 0, 0)
    End If
    Cells(4, 3) = SI
End Sub

Sub calcheightcoeff(SI)
' Calculate coefficients for height model
    If heightmodel = 1 Then
        ha = Exp(hNSWa)
        hb = 1 / (hNSWb + hNSWp * SI)
    ElseIf heightmodel = 2 Then
        ha = Exp(ha0 + ha1 * SI)
        hb = 1 / (hb0 + hb1 * SI)
    Else
        ' latitude = Cells(3, 6)
    ' Corrected Latitude issue
        latitude = Abs(Cells(3, 6))
        altitude = Cells(4, 6)
        ha = Exp(hae0 + hae1 * latitude + hae2 * altitude)
        
        latitude = Abs(Cells(3, 6))
        altitude = Cells(4, 6)
        ha = Exp(hae0 + hae1 * latitude + hae2 * altitude)
        
        hb = 1 / (hbe0 + hbe1 * SI)
    End If
End Sub

Function CalcDBHfromBA(BA, N) As Double
' Calculate DBH from BA and stocking
            
    CalcDBHfromBA = Sqr(1.273 * BA / N) * 100

End Function

Function CalcBAfromDBH(dbh, N) As Double
' Calculate BA from DBH and stocking
            
    CalcBAfromDBH = N / 1.273 * (dbh / 100) ^ 2

End Function

Function calcBAfromVol(mth, vol, stock) As Double
' Calculate BA from MTH, Volume & Stocking using appropriate stand-level volume function
        
    If vol <= 0 Or mth <= 1.6 Or stock <= 0 Then
        calcBAfromVol = 0
    ElseIf voltable = 1 Or voltable = 2 Then
        calcBAfromVol = vol / (mth * (v(voltable, 1) * (mth - 1.4) ^ v(voltable, 2) + v(voltable, 3)))
    ElseIf voltable = 10 Or voltable = 11 Then
        calcBAfromVol = vol / (v(voltable, 1) + v(voltable, 2) * mth + v(voltable, 3) * stock)
    Else
        calcBAfromVol = Exp(v(voltable, 1) + _
            v(voltable, 2) * Log(mth) + _
            v(voltable, 3) * Log(vol) + _
            v(voltable, 4) * Log(stock) + _
            v(voltable, 5) * Log(stock) * Log(stock) + _
            v(voltable, 6) * Log(mth) * Log(mth) + _
            v(voltable, 7) * Log(mth) * Log(stock) + _
            v(voltable, 8) * Log(vol) * Log(stock))
    End If
    
End Function

Function CalcVol(mth, BA, stock) As Double
' Calculate Volume from MTH and BA using stand-level volume function
    
    If BA <= 0 Or mth <= 1.6 Or stock <= 0 Then
        CalcVol = 0
    ElseIf voltable = 1 Or voltable = 2 Then
        CalcVol = mth * BA * (v(voltable, 1) * (mth - 1.4) ^ v(voltable, 2) + v(voltable, 3))
    ElseIf voltable = 10 Or voltable = 11 Then
        CalcVol = BA * (v(voltable, 1) + v(voltable, 2) * mth + v(voltable, 3) * stock)
    Else
        CalcVol = Exp(-(v(voltable, 1) + v(voltable, 2) * Log(mth) _
        + v(voltable, 4) * Log(stock) + v(voltable, 5) * Log(stock) ^ 2 _
        + v(voltable, 6) * Log(mth) ^ 2 + v(voltable, 7) * Log(mth) * Log(stock) _
        - Log(BA)) / (v(voltable, 3) + v(voltable, 8) * Log(stock)))
    End If
    
End Function

Function CalcMTH(SI, HAge) As Double
' Calculate MTH from SI and Age using Mina's height function

    Call calcheightcoeff(SI)
    CalcMTH = 0.25 + (SI - 0.25) * ((1 - Exp(-ha * HAge)) / (1 - Exp(-ha * 20))) ^ hb
        
End Function

Function Calcagezero() As Double
' Calculate age when MTH=1.4

    Calcagezero = -Log(-(1 - Exp(-ha * 20)) * ((1.4 - 0.25) / (SI - 0.25)) ^ (1 / hb) + 1) / ha
'    Calcagezero = 8.6877 * Exp(-0.0539 * SI)   'This is the original 2004 code
    
End Function

Function calcMeanht(mth, stock) As Double
' Mean height model (derived from stocking trials)
    Const A As Double = 0.07, B As Double = -0.00399
    If Excel.WorksheetFunction.IsNumber(mth) And Excel.WorksheetFunction.IsNumber(stock) Then
        calcMeanht = mth * (1 - A * (1 - Exp(B * (stock - 100))))
    End If
End Function

Function MH2MTH(MH, stock)
    ' calculate MTH from MH and SPH
    Const A As Double = 0.07, B As Double = -0.00399
    If Excel.WorksheetFunction.IsNumber(mth) And Excel.WorksheetFunction.IsNumber(stock) Then
        MH2MTH = (1 / MH * (1 - A * (1 - Exp(B * (stock - 100))))) ^ -1
    End If
End Function

Function CalcDBH(I300 As Double, SI As Double, age As Double, stock As Double)
    'This function fixes a kink in the BA curve which appears at about age 28 years.
    'DBH is predicted using cubic interpolation between ages 20 and 40 years.
    Dim DBH1 As Double, DBH2 As Double, DBH3 As Double, DBH4 As Double, Y0 As Double, Y1 As Double, _
        Y0p As Double, Y1p As Double, A As Double, B As Double, C As Double, D As Double
    If age <= 20 Or age >= 40 Then
        A200 = CalcA200start(age, I300, SI)
        CalcDBH = DBHmodel(A200, SI, age, stock)
    Else
        A200 = CalcA200start(19.5, I300, SI)
        DBH1 = DBHmodel(A200, SI, 19.5, stock)
        A200 = CalcA200start(20.5, I300, SI)
        DBH2 = DBHmodel(A200, SI, 20.5, stock)
        A200 = CalcA200start(39.5, I300, SI)
        DBH3 = DBHmodel(A200, SI, 39.5, stock)
        A200 = CalcA200start(40.5, I300, SI)
        DBH4 = DBHmodel(A200, SI, 40.5, stock)
        Y0 = (DBH1 + DBH2) / 2
        Y1 = (DBH3 + DBH4) / 2
        Y0p = (DBH2 - DBH1)
        Y1p = (DBH4 - DBH3)
        A = Y0
        B = Y0p
        D = (2 * (Y0 + Y0p * 20 - Y1) + 20 * (Y1p - Y0p)) / (20 ^ 3)
        C = (Y1p - Y0p - 3 * D * 20 ^ 2) / (2 * 20)
        CalcDBH = A + B * (age - 20) + C * (age - 20) ^ 2 + D * (age - 20) ^ 3
    End If
End Function

Function DBHmodel(A200, SI, age, stock) As Double
' Predict DBH at given age and stocking using 300 Index model
'   A200 is DBH at 200 sph and age 30 with no pruning or thinning
    Dim stk As Double, A As Double, B As Double, _
        q As Double, P As Double, qq As Double

    ' agezero is age when MTH=1.4
    agezero = Calcagezero()
    site_effect = A200 / da1 - 1
    stk = stock
    
    A = da1 * (1 + site_effect)
    B = db2 * (db1 + dbSI * (SI - 28) + dbdia * site_effect + dbsidia * (SI - 28) * site_effect)
    If B > -0.05 Then B = -0.05         'Check that b is within reasonable bounds
    If age < agezero Then
        DBHmodel = 0
    Else
        D200 = OldAgeCorrection(age, agezero, B) * A * ((1 - Exp(B * (age - agezero))) / (1 - Exp(B * (30 - agezero)))) ^ dc
        'Modify q to eliminate unnatural behaviour at stockings near 200 sph, March 2011
        If stk > 220 Then
            qq = (Log(stk) - Log(200)) ^ dr2
        Else
            qq = 2 * (Log(220) - Log(200)) ^ dr2 - (Log(242) - Log(stk)) ^ dr2
        End If
        q = dr * (1 + drsi * (SI - 28)) * qq
'        q = dr * (1 + drsi * (SI - 28)) * Sgn(stk - 200) * (Abs(Log(stk) - Log(200))) ^ dr2
        P = dl + dm * stk + dn * site_effect
        DBHmodel = D200 - q * Log(1 + Exp(Ds * (D200 - P)))
        'High stocking correction, March 2011 - Apply if BA decreasing with stocking
        If stk > 250 Then
            If dBA_dN(D200, P, q, stk) <= 0 Then
                N_MaxBA = MaxBAStocking(D200, site_effect, SI, stk)
                q = dr * (1 + drsi * (SI - 28)) * Sgn(N_MaxBA - 200) * (Abs(Log(N_MaxBA) - Log(200))) ^ dr2
                P = dl + dm * N_MaxBA + dn * site_effect
                DBHmodel = (D200 - q * Log(1 + Exp(Ds * (D200 - P)))) * Sqr(N_MaxBA / stk)
            End If
        End If
    End If
    If DBHmodel < 0 Then DBHmodel = 0
        
End Function

Sub testdBA_dN()
    Dim test As Double
    test = dBA_dN(57.81, 4.33, 7.49, 5200)
End Sub

Function dBA_dN(D200, P, q, N) As Double
'Derivative of predicted BA wrt Stocking
    Dim dp_dN As Double, dq_dN As Double, dD_dN As Double, D As Double
    dp_dN = dm
    dq_dN = q * dr2 / N / (Log(N) - Log(200))
    dD_dN = -Ds * D200 * dq_dN + Ds * P * dq_dN + Ds * q * dp_dN
    D = approxDBH(D200, P, q)
    If D < 0 Then
        dBA_dN = 0
    Else
        dBA_dN = D * (D + 2 * N * dD_dN)
    End If
End Function

Sub testmaxBAstock()
    Dim maxN As Double
    maxN = MaxBAStocking(57.81, -0.036992, 31.69, 5200)
End Sub

Function MaxBAStocking(D200, site_effect, SI, N) As Double
'Obtain the stocking that produces maximum predicted BA
    Dim NA As Double, NB As Double, NC As Double, FA As Double, FB As Double, _
        FC As Double, j As Long, P As Double, q As Double, D As Double
    NA = 250
    q = dr * (1 + drsi * (SI - 28)) * Sgn(NA - 200) * (Abs(Log(NA) - Log(200))) ^ dr2
    P = dl + dm * NA + dn * site_effect
    FA = dBA_dN(D200, P, q, NA)
    NB = N
    q = dr * (1 + drsi * (SI - 28)) * Sgn(NB - 200) * (Abs(Log(NB) - Log(200))) ^ dr2
    P = dl + dm * NB + dn * site_effect
    FB = dBA_dN(D200, P, q, NB)
    For j = 1 To 13
        NC = (NA + NB) / 2
        q = dr * (1 + drsi * (SI - 28)) * Sgn(NC - 200) * (Abs(Log(NC) - Log(200))) ^ dr2
        P = dl + dm * NC + dn * site_effect
        FC = dBA_dN(D200, P, q, NC)
        If FA * FC < 0 Then
            NB = NC
            FB = FC
        Else
            NA = NC
            FA = FC
        End If
    Next j
    MaxBAStocking = NC
    
End Function

Function approxDBH(D200, P, q) As Double
    approxDBH = D200 - q * Ds * (D200 - P)
End Function

Sub testOldAgeCorr()
    Dim test As Double
    test = OldAgeCorrection(40, 1.19, -0.1459)
End Sub

Function OldAgeCorrection(age, agez, B)
    Dim T As Double
'    Const a1 As Double = 1, a2 As Double = 0.003061989, a3 As Double = 1.140404981, a4 As Double = 2.407818773, _
        a5 As Double = 20.82973409, a6 As Double = -0.072897
'    Const a1 As Double = 1, a2 As Double = 0.011674652, a3 As Double = 1.497700593, a4 As Double = 0.8, _
        a5 As Double = 19.49383991, a6 As Double = -0.072897
'    Const a1 As Double = 1, a2 As Double = 0.008757209, a3 As Double = 1.363084086, a4 As Double = 1, _
        a5 As Double = 19.99756941, a6 As Double = -0.072897
'    Const a1 As Double = 1, a2 As Double = 0.004497048, a3 As Double = 1.187782959, a4 As Double = 1.725660042, _
'        a5 As Double = 20.39169105, a6 As Double = -0.072897
'    Const a1 As Double = 1, a2 As Double = 0.001176493, a3 As Double = 0.967256067, a4 As Double = 4.347880931, _
'        a5 As Double = 25
    Const a1 As Double = 1, a2 As Double = 0.001473784, a3 As Double = 0.973636099, a4 As Double = 4.350585474, _
        a5 As Double = 25
'    t = b * (age - agez) / a6 - a5
    T = (age - agez) - a5
    If T < 0 Then T = 0
    OldAgeCorrection = 1 + a4 * (1 - Exp(-a2 * T)) ^ a3
'    OldAgeCorrection = 1
End Function

Function CalcAge(dbh As Double, A200 As Double, stock As Double, SI As Double) As Double
' Calculate Age from DBH, A200 (DBH at 200sph, age 30years), Stocking & SI using Bisection Method
    CalcAge = Bisection(0.001, 150, 15, 3, A200, SI, stock, dbh)
End Function

Function CalcA200(dbh As Double, age As Double, stock As Double, SI As Double) As Double
' Calculate A200 from DBH, Age, Stocking & SI using Bisection Method
    agezero = Calcagezero()
    CalcA200 = Bisection(10, 150, 20, 4, SI, age, stock, dbh)
End Function

Function CalcA200start(age As Double, I300 As Double, SI As Double) As Double
'Calculate A200 from the 300 Index and SI
    Dim BA300_30 As Double, DBH300_30 As Double, adjI300 As Double, i300adjustment As Double, _
        b_adj As Double, c_adj As Double, k1_adj As Double, k2_adj As Double, k3_adj As Double, _
        k4_adj As Double, k_adj As Double
                
    adjI300 = I300  'Standard model
    b_adj = 0.0206 / 19.488
    c_adj = -0.0182 / 100 / 19.488
    k1_adj = 25
    k2_adj = 55
    k3_adj = 215.97
    k4_adj = -0.05532
    k_adj = k3_adj * Exp(k4_adj * I300)
    
    '300 Index bias correction for ages<6.77
    If bias_young And age < 6.77 Then
        i300adjustment = 180.5 * adjI300 ^ (-3.256) * (age - 6.77) ^ 2
        If i300adjustment > 5 Then i300adjustment = 5
        adjI300 = adjI300 + i300adjustment

    'Original age>25 year bias correction
'    ElseIf bias_old And age >= 25 And age < 60 Then
'        adjI300 = adjI300 * (1 + 0.00042 * (age - 25) ^ 2)
'    ElseIf bias_old And age >= 60 Then
'        adjI300 = adjI300 * (-0.255 + 0.0295 * age)
    End If

    '300 Index drift correction for low and high SI
    If bias_SI Then
        If SI < 25 And SI >= 15 Then adjI300 = adjI300 * (30 - 0.02 * (25 - SI) * (age - 28.6)) / 30
        If SI < 15 Then adjI300 = adjI300 * (30 - 0.2 * (age - 28.6)) / 30
        If SI > 35 And SI <= 45 Then adjI300 = adjI300 * (30 - 0.02 * (SI - 35) * (age - 28.6)) / 30
        If SI > 45 Then adjI300 = adjI300 * (30 - 0.2 * (age - 28.6)) / 30
    End If
    
    '300 Index drift correction
    If age < 30 Then
        adjI300 = adjI300 * (30 + drift * (age - 28.6)) / 30
'    ElseIf age < 45 Then
'        adjI300 = adjI300 * (30 + drift * (45 - age) / 15 * (age - 28.6)) / 30
    Else
        adjI300 = adjI300
    End If
    
    BA300_30 = calcBAfromVol(CalcMTH(SI, 30), adjI300 * 30, 300)
    DBH300_30 = CalcDBHfromBA(BA300_30, 300)
    CalcA200start = CalcA200(DBH300_30, 28.7, 300, SI)
End Function

Function heightmod() As Long
' Determine height model - 1 = NSW, 2 = Simple NZ, 3 = Environmental NZ
        If LCase(Cells(64, 4)) = "x" Then
            heightmod = 1
        ElseIf (IsEmpty(Cells(3, 6)) Or IsEmpty(Cells(4, 6))) Or _
            (Abs(Cells(3, 6)) < 30 Or Abs(Cells(3, 6)) > 48) Then   'Test whether latitude within NZ range
            heightmod = 2
        Else
            heightmod = 3
        End If
End Function

Function Bisection(xlower As Double, xupper As Double, niterations As Long, _
    fnno As Long, p1 As Double, p2 As Double, p3 As Double, p4 As Double)
'Find when function number fnno equals zero using the bisection method
    
    Dim xA As Double, xB As Double, xC As Double, FA As Double, FB As Double, _
        FC As Double, j As Long
    xA = xlower
    FA = fn(xA, fnno, p1, p2, p3, p4)
    xB = xupper
    FB = fn(xB, fnno, p1, p2, p3, p4)
    For j = 1 To niterations
        xC = (xA + xB) / 2
        FC = fn(xC, fnno, p1, p2, p3, p4)
        If FA * FC < 0 Then
            xB = xC
            FB = FC
        Else
            xA = xC
            FA = FC
        End If
    Next j
    Bisection = xC
    
End Function

Function fn(X As Double, fnno As Long, p1 As Double, p2 As Double, p3 As Double, p4 As Double)
'Gives the function to be zeroised using the bisection subroutine
    If fnno = 1 Then
        I300 = X
        Call Growth(False)
        fn = DBH300 - dbh
    ElseIf fnno = 2 Then
        fn = p1 - CalcMTH(X, p2)    'p1=MTH, p2=age
    ElseIf fnno = 3 Then
        fn = p4 - DBHmodel(p1, p2, X, p3) 'p1=A200, p2=SI, p3=stock, p4=dbh
    ElseIf fnno = 4 Then
        fn = p4 - DBHmodel(X, p1, p2, p3) 'p1=SI, p2=Age, p3=stock, p4=dbh
    End If
End Function

Sub earlyield()
' Correct early volume predictions in yield tables
' It is assumed that early volume growth of individual trees is proportional to Age^2.7 (based on data from VMAN)
' This function is applied when DBH is less than 2 cm
' Age zero volume from planted seedlings supplied by Beets
' Volume weightings for these ages were derived using VMAN
    Dim initialvol As Double, age As Double, treevolinc As Double, i As Long, _
        dbh As Double, prevdbh As Double, j As Long, k As Double, T As Double
    Worksheets("300 Index").Activate
    initialvol = 0.0000064  'Volume of seedling at planting (m3) - Beets
    dbh = 0
    For i = 5 To 20
        age = Cells(i, 7)
        prevdbh = dbh
        dbh = Cells(i, 16)
        If dbh >= 2 And prevdbh < 2 Then
            treevolinc = Cells(i, 12) / Cells(i, 8) - initialvol
            k = treevolinc / (age ^ 2.7)  'parameter k ensures that volume is correct when T=age
            For j = i - 1 To 5 Step -1
                T = Cells(j, 7)
                If T < 0 Then T = 0
                Cells(j, 12) = (initialvol + k * T ^ 2.7) * Cells(j, 8)
                If Cells(j, 9) <> 0 Then
                    Cells(j, 13) = (initialvol + k * T ^ 2.7) * Cells(j, 9)
                End If
            Next j
        End If
    Next i
End Sub

Sub mortvol()
' This subroutine estimates the volume lost to mortality within each growth increment
    Dim numrows As Long, i As Long, sph1 As Double, sph2 As Double, vol1 As Double, vol2 As Double
    Worksheets("300 Index").Activate
    numrows = Range(Cells(5, 7), Cells(5, 7).End(xlDown)).Rows.Count
    Cells(5, 39) = 0
    For i = 6 To numrows + 4
        sph1 = Cells(i - 1, 9)
        vol1 = Cells(i - 1, 13)
        If sph1 = 0 Then
            sph1 = Cells(i - 1, 8)
            vol1 = Cells(i - 1, 12)
        End If
        sph2 = Cells(i, 8)
        vol2 = Cells(i, 12)
        Cells(i, 39) = mort_vol(sph1, vol1, sph2, vol2)
    Next i
End Sub

Function mort_vol(sph1, vol1, sph2, vol2) As Double
' Predict volume lost to mortality within a growth increment
        mort_vol = (vol1 + vol2) * 0.5 * ((sph1 / sph2) ^ 0.541 - 1)
End Function

Sub density()
    'This subroutine predicts wood density of each year growth sheath, and of the stem
        'densityinfo = 1 Estimate density from environmental information
        'densityinfo = 2 Estimate density from BH outerwood core
        'densityinfo = 3 Estimate density from BH density for specified range of rings
        'densityinfo = 4 Assume constant average density within stem
        'densityinfo = 5 No density given - assume national average
    
    Dim numrows As Long, i As Long, age As Long, od As Double, vol As Double, _
        prevvol As Double, densityinfo As Long, CoreDBH As Double, Coreline As Long, _
        outdensring As Double, stocking As Double, outdens_ringwidth As Double, _
        ringwidth As Double, Wcal As Double, prevdbh As Double, currdbh As Double, _
        currage As Double, prevage As Double, densitymodel As Long, InnerDBH As Double, _
        OuterDBH As Double, first_ringwidth As Double, ring As Double

    ' Input density information
    SoilC = Cells(75, 4)
    SoilN = Cells(76, 4)
    Temp = Cells(77, 4)
    CoreDens = Cells(78, 4)
    CoreAge = Cells(79, 4)
    InnerRing = Cells(80, 4)
    OuterRing = Cells(81, 4)
    GeneticAdj = Cells(82, 4)
    densitymodel = Cells(83, 4)
    
    ' Determine type of density information
    If SoilC <> 0 And SoilN <> 0 And Temp <> 0 Then
        densityinfo = 1 'Estimate density from environmental information
    ElseIf CoreDens <> 0 And CoreAge <> 0 Then
        densityinfo = 2 'Estimate density from BH outerwood core
    ElseIf CoreDens <> 0 And InnerRing <> 0 And OuterRing <> 0 Then
        densityinfo = 3 'Estimate density from BH density for specified range of rings
    ElseIf CoreDens <> 0 Then
        densityinfo = 4 'Assume constant average density within stem
    Else
        densityinfo = 5 'No density given - assume national average
    End If
        
    Worksheets("300 Index").Activate
    numrows = Range(Cells(5, 7), Cells(5, 7).End(xlDown)).Rows.Count
    agezero = Calcagezero()
    
    If densityinfo = 1 Then  'Predict OW core density at age 26 years from environmental variables
        CoreAge = 26
'        For i = 1 To numrows
'            stocking = Cells(i + 4, 8)  'Stocking at age 20 years is used to adjust density
'            If Cells(i + 4, 7) = 20 Then
'                Exit For
'            End If
'        Next i
        
        'New density model uses stocking of 250 stems/ha based on WQI Benchmarking trial
        stocking = 250
        
        CoreDens = outdens26(SoilC, SoilN, Temp, stocking, GeneticAdj)
        outdensring = 18.95 - 0.024 * SI
        Wcal = 10.19 + 0.0893 * I300 - 0.255 * SI + 0.00373 * SI * SI - 0.00339 * I300 * SI
    End If
        
'    If densityinfo <= 2 Then  'Find OW core inner and outer rings
'        For i = 1 To numrows
'            CoreDBH = Cells(i + 4, 16)
'            If Cells(i + 4, 7) = CoreAge Then
'                Coreline = i + 4
'                Exit For
'            End If
'        Next i
'        OuterRing = CoreAge - agezero + 0.5
'        For i = 1 To Coreline - 5
'            If Cells(Coreline - i, 16) < CoreDBH - 10 Then 'Assuming core is 5 cm long, find age when DBH is 10 cm less
'                If CoreDBH - 10 - Cells(Coreline - i, 16) < Cells(Coreline - i + 1, 16) - (CoreDBH - 10) Then
'                    InnerRing = OuterRing - i
'                Else
'                    InnerRing = OuterRing - i + 1
'                End If
'                'If densityinfo=2 then obtain calibration ring width from core
'                If densityinfo = 2 Then Wcal = 10 * (Cells(Coreline, 16) - Cells(Coreline - i, 16)) / 2 / _
'                    (Cells(Coreline, 7) - Cells(Coreline - i, 7))
'                Exit For
'            End If
'            InnerRing = 0
'        Next i
'    End If
'
'    'If densityinfo=3 then obtain calibration ring width from core
'    If densityinfo = 3 Then
'        OuterDBH = Cells(numrows + 4, 16)
'        For i = 1 To numrows
'            If Int(Cells(i + 4, 7) - agezero) = InnerRing Then InnerDBH = Cells(i + 4, 16)
'            If Int(Cells(i + 4, 7) - agezero) = OuterRing Then OuterDBH = Cells(i + 4, 16)
'        Next i
'        Wcal = 10 * (OuterDBH - InnerDBH) / 2 / (OuterRing - InnerRing)
'    End If
'
'    If densityinfo <= 3 Then
'        'Find average age of density core
'        outdensring = (OuterRing + InnerRing) / 2
'    End If
'
'    If densityinfo = 5 Then
'        CoreDens = 470
'        outdensring = 23
'    End If
    
    
    'Find first ring width of first complete ring
    
    For i = 6 To numrows + 4
        If Cells(i - 1, 17) <> 0 Then
            prevdbh = Cells(i - 1, 17)
        Else
            prevdbh = Cells(i - 1, 16)
        End If
        If prevdbh <> 0 Then
            currdbh = Cells(i, 16)
            prevage = Cells(i - 1, 7)
            currage = Cells(i, 7)
            first_ringwidth = 10 * (currdbh - prevdbh) / (currage - prevage) / 2
            Exit For
        End If
    Next i
    
    'Predict sheath and whole stem density at each age
    
    For i = 5 To numrows + 4
        age = Cells(i, 7)
        vol = Cells(i, 12)
        
        'Obtain ring width (mm)
        If i = 5 Then
            ringwidth = first_ringwidth
        Else
            If Cells(i - 1, 17) <> 0 Then
                prevdbh = Cells(i - 1, 17)
            Else
                prevdbh = Cells(i - 1, 16)
            End If
            If prevdbh = 0 Then
                ringwidth = first_ringwidth
            Else
                currdbh = Cells(i, 16)
                prevage = Cells(i - 1, 7)
                currage = Cells(i, 7)
                ringwidth = 10 * (currdbh - prevdbh) / (currage - prevage) / 2
            End If
        End If
      
'        'BH density of outer 5 rings
'        If densityinfo = 4 Then
'            Cells(i, 42) = CoreDens / 1000
'        Else
'            If age > 2 Then
'                If densitymodel = 2 Then
'                    Cells(i, 42) = outdens(age - 2, ringwidth, CoreDens, outdensring, Wcal) / 1000
'                Else
'                    Cells(i, 42) = old_outdens(age - 2, CoreDens, outdensring) / 1000
'                End If
'            Else
'                If densitymodel = 2 Then
'                    Cells(i, 42) = outdens(0, ringwidth, CoreDens, outdensring, Wcal) / 1000
'                Else
'                    Cells(i, 42) = old_outdens(0, CoreDens, outdensring) / 1000
'                End If
'            End If
'        End If
        
        'Growth sheath density
        If densityinfo = 4 Then
            Cells(i, 44) = CoreDens / 1000
            Cells(i, 40) = CoreDens / 1000
        Else
            If densitymodel = 2 Then
                ring = age - agezero
                If ring < 1 Then ring = 1
                od = outdens(ring, ringwidth, CoreDens, outdensring, Wcal)
'               od = outdens(age, ringwidth, CoreDens, outdensring, Wcal)
            Else
                od = old_outdens(age, CoreDens, outdensring)
            End If
            Cells(i, 44) = od / 1000
            Cells(i, 40) = sheathdens(od, age) / 1000
        End If
        
'        'Whole stem density
'        If i = 5 Then
'            Cells(5, 41) = Cells(5, 40)
'        Else
'            Cells(i, 41) = (Cells(i, 40) * (vol - prevvol) + Cells(i - 1, 41) * prevvol) / vol
'        End If
'        If Cells(i, 13) <> 0 Then
'            prevvol = Cells(i, 13)
'        Else
'            prevvol = vol
'        End If
    Next i

End Sub

Function outdens(ring, ringwidth, refdens, refring, refwidth) As Double
    ' Predict density for a breast height ring from reference density at a reference ring
    Dim S As Double, ringwidth_adj As Double
    S = (refdens - 477.8 + 46.2 * Log(refwidth) + 84.8 * Exp(-0.258 * refring)) / _
        (1 - 46.2 * 0.0042 * Log(refwidth))
    ringwidth_adj = ringwidth
    If ringwidth_adj < 1.5 Then ringwidth_adj = 1.5 'minimum ring width for density model is 1.5 mm
    outdens = 477.8 + S - 46.2 * (1 + 0.0042 * S) * Log(ringwidth_adj) - 84.8 * Exp(-0.258 * ring)
End Function

'2007 wood density function
Function old_outdens(ring, refdens, refring) As Double
    ' Predict density for a breast height ring from reference density at a reference ring
    Dim B As Double
    Const A As Double = 332.2, C As Double = 0.0193, g As Double = 0.0809, k As Double = 23.8, _
        D As Double = 10.94, validation_adjustment As Double = 0.968
    If refring < k Then
        B = (refdens - A * validation_adjustment) / (refring - C * refring ^ 2)
    Else
        B = (refdens - A * validation_adjustment) / (D + g * refring)
    End If
    If ring < k Then
        old_outdens = A * validation_adjustment + B * (ring - C * ring ^ 2)
    Else
        old_outdens = A * validation_adjustment + B * (D + g * ring)
    End If
End Function

Function sheathdens(outdens, age) As Double
    ' Predict density of growth sheath from breast height density at corresponding ring
    Dim age1 As Double
    Const S As Double = 1.33415953, T As Double = -0.0108173186465, u As Double = -0.000963837, v As Double = 0.000061770373226, _
        w As Double = 0.00002435373794
    age1 = age
    If age > 30 Then age1 = 30 + (age - 30) * 0.33
    If age > 40 Then age1 = 34
    sheathdens = (S + T * age1 + u * outdens + v * age1 ^ 2 + w * age1 * outdens) * outdens
End Function

Function outdens26(SoilC, SoilN, Temp, stocking, GeneticAdj) As Double
    ' Predict outerwood BH density at age 26 years for given soil N & C, mean temperature, and stocking
    Dim outdens_26_250 As Double, CN_adj As Double
    'The following parameters have no genetic adjustment
    Const P As Double = 143, q As Double = 15.9, R As Double = 4.1, A As Double = 332.2, _
        z As Double = 18.64, validation_adjustment As Double = 0.968
    If SoilN <= 0.014 Then CN_adj = 50 Else CN_adj = SoilC / (SoilN - 0.014)
    If CN_adj > 50 Then CN_adj = 50
    outdens_26_250 = P + q * Temp + R * CN_adj
    outdens26 = (A * validation_adjustment + (outdens_26_250 * validation_adjustment - A * validation_adjustment) * _
        (z + Sqr(stocking)) / (z + Sqr(250))) * (1 + GeneticAdj / 100)
End Function

Sub CalcOffsets()
    ' Derive offsets for MTH and DBH from a measurement.
    ' For use when user specifies 300 Index and Site Index together with a Stand Measurement (Mode 2).
    
    Dim BA300 As Double, Stock300 As Double, MTH300 As Double, Vol300 As Double
    
    Call Inputparms
    Call voltab
    DBH_calibration_age = Cells(7, 3)
    Stock300 = Cells(8, 3)
    DBH300 = Cells(9, 3)
    If DBH300 = 0 Then
        BA300 = Cells(10, 3)
        If BA300 <> 0 Then
            DBH300 = CalcDBHfromBA(BA300, Stock300)
        Else
            Vol300 = Cells(11, 3)
            DBH300 = CalcDBHfromBA(calcBAfromVol(MTH300, Vol300, Stock300), Stock300)
        End If
    End If
    MTH_calibration_age = Cells(14, 3)
    MTH300 = Cells(15, 3)
    
    OUTPUT = False
    
    maxage = DBH_calibration_age
    steps = maxage / steplength
    Call Growth(OUTPUT)
    DBHsqd_add_offset = DBH300 ^ 2 - dbh ^ 2
    DBHsqd_mult_offset = DBH300 ^ 2 / dbh ^ 2
    
    maxage = MTH_calibration_age
    steps = maxage / steplength
    Call Growth(OUTPUT)
    MTH_add_offset = MTH300 - mth
    MTH_mult_offset = MTH300 / mth

End Sub


