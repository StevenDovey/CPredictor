Attribute VB_Name = "Module5"
'*************************************************************************************************************
'
'   This module contains C_Change code translated from FORTRAN code in January 2008 by Mark Kimberley
'
'*************************************************************************************************************

Option Explicit

Public line5 As Long, line6 As Long, line7 As Long, line8 As Long, line10 As Long, col5 As Long, col6 As Long, col7 As Long, col8 As Long, col10 As Long

Dim X(16) As Double, ID As Long, INPX As Long, GCLMAX As Double, X3MAX As Double, X4MAX As Double, _
    X5MAX As Double, X6MAX As Double, HTBGC As Double, GCL As Double, BA As Double, BASP(100) As Double, _
    BAN(10) As Double, XTRACF(10) As Double, XTRACT(10) As Double, S As Double, LA1Y As Double, LA2Y As Double, LA3Y As Double, _
    ADIST(11) As Double, SPHAN(10) As Double, AY1A As Double, AY1B As Double, AY1C As Double, PNA As Double, _
    PNB As Double, pBA As Double, PBB As Double, PBC As Double, CONS As Double, PRNHT(10) As Double, GCFRE As Double, _
    HT As Double, SPHA As Double, AGECK As Double, TLA As Double, SSABAR As Double, SDBAR As Double, _
    B31 As Double, B32 As Double, AM1 As Double, AM2 As Double, BM1 As Double, BM2 As Double, SFINT As Double, _
    SFSPE As Double, RNAPA As Double, RNAPB As Double, RNAPC As Double, TMA As Double, TMB As Double, _
    dens As Double, VMA As Double, VMB As Double, VMC As Double, VMD As Double, INDEAD As Long, _
    INPNAP As Long, HINC As Double, T As Double, MAXL1Y, TVOL As Double, vol As Double, _
    BABGC As Double, XN(16), XP(16), TBRCH As Double, FNA As Double, FNB As Double, SNA As Double, SNB As Double, _
    BNA As Double, BNB As Double, YESN As Double, HTSP(100) As Double, TVOLSP(100) As Double, _
    SSA As Double, SBA As Double, CRA As Double, PMAI As Double, SX3 As Double, NLIMITED As Long, _
    DIAGN(100) As Double, AGEND As Double, CNRAT As Double, XTRACC(10) As Double, XTR As Double, XTRN As Double, XTRP As Double, _
    XTF As Double, XTFN As Double, XTFP As Double, DMINI As Double, FLOSS As Double, CSHRUB As Double, ICOUNT As Long, _
    RADIA As Double, DENWH(100) As Double, DENINC(100) As Double, _
    F(16, 16) As Double, XMN1(16) As Double, XNMN1(16) As Double, XPMN1(16) As Double, _
    UPNAG As Double, UPPAG As Double, UPNBG As Double, UPPBG As Double, UPNFF As Double, UPPFF As Double, _
    DecayN As Double, DecayP As Double, DeficitN As Double, DeficitP As Double, PROPORTION_STEM_LOSS As Double, _
    PROPORTION_BRANCH_LOSS As Double, PROPORTION_ATTACHEDBRANCH_LOSS As Double, _
    PROPORTION_COARSEROOT_LOSS As Double, PROPORTION_NEEDLE_LOSS As Double, _
    XN14THINLOSS As Double, XP14THINLOSS As Double, _
    SoilC As Double, SoilN As Double, SoilOrganicP As Double, SoilBray2P As Double, stem_P_reduction As Double
    
Sub DRYMAT()

'C******************************************************************************
'C
'C   VERSION NUMBER IS = 3.1   EDIT THIS NUMBER HERE AND IN SUBROUTINE OUTPUT.
'C   Version 3.1 includes harvest extraction options that relate to LTSP1 trial series
'C   VERSION 2.0 INCORPORATES DENSITY FROM SITE FERTILITY DATA, WITH GROSS
'C   AND NET SP VOLUME INPUT CONVERTED TO MASS.
'C   DENSITY CAN BE WHOLE STEM (DENSREG 5) OR BY RING DENSREG 6). THE LATTER
'C   IS EXPECTED TO PROVIDE SUPERIOR PREDICTIONS OF CARBON SEQUESTRATION COMPARED
'C   TO THE FORMER.

'C PURPOSE
'C   DRYMAT IS A COMPARTMENT MODEL FOR DYNAMICALLY SIMULATING THE DRY MATTER
'C   AND NUTRIENT CONTENT OF MANAGED RADIATA PINE STANDS OVER A ROTATION. THE
'C   MODEL CONSISTS OF CONSERVATIVE FLOW EQUATIONS WHICH ARE SOLVED
'C   SIMULTANEOUSLY, ON AN ANNUAL BASIS, USING STATE-DETERMINED TRANSITION
'C   EQUATIONS GIVEN BY:
'C
'C dX(T) / DT = Z(T) + A(T) * X(T)
'C
'C   WHERE t IS TIME AND X IS AN ARRAY OF COMPARTMENT WEIGHTS (TONNES DRY
'C   MATTER/HA). THE ELEMENTS OF THE ARRAY ARE DEFINED AS:
'C
'C        X(3) = 0-1 YEAR OLD NEEDLE
'C        X(4) = 1-2 YEAR OLD NEEDLE
'C        X(5) = 2+ YEAR OLD NEEDLE
'C        X(6) = LIVE BRANCH
'C        X(7) = DEAD BRANCH
'C        X(8) = STEM WOOD
'C        X(9) = COARSE ROOT
'C        X(10)= NEEDLE LITTER
'C        X(11)= BRANCH LITTER
'C        X(12)= STEM LITTER
'C        X(13)= COARSE ROOT LITTER
'C        X(14)= LIVE FINE ROOT
'C        X(15)= FINE ROOT LITTER
'C        X(16)= STEM BARK
'C
'C   Z IS A TIME VARYING FORCING DETERMINING GROWTH RATE:
'C
'C          Z  = NET PRIMARY PRODUCTION (TONNES/HA/YEAR)
'C
'C   Y ALLOWS FOR TIME VARYING OUTPUTS IN TERMS OF DRY MATTER FROM THE STAND:
'C
'C        Y(t) = DRY MATTER LOSS (TONNES/HA/YEAR)
'C
'C   A(t) IS A MATRIX OF TIME VARYING RATE COEFFICIENTS (OF DIMENSION 16)
'C
'C   NOTE THAT X(1) AND X(2) ARE ABSTRACT AND SHOULD NOT BE INCLUDED IN
'C   SUMMATIONS OF STAND DRY MATTER CONTENT, X(1) BEING EQUIVALENT TO Z, AND
'C   X(2) BEING EQUIVALENT TO Z - MORTALITY FLOWS.
'C
'C*****************************************************************************
'C
'C   PROGRAM STRUCTURE:
'C
'C   DRYMAT IS WRITTEN IN FORTRAN AND IS COMPRISED OF A MAIN PROGRAM AND
'C   SUBROUTINES. AN OUTLINE OF THE PROGRAM STRUCTURE IS PRESENTED HERE,
'C   WITH MORE DETAILS PROVIDED WITHIN EACH SUBROUTINE. MAIN ALLOWS THE INPUT
'C   OF SIMULATION INFORMATION AND ALSO CONTROLS VARIOUS ASPECTS OF THE
'C   SIMULATION WHICH ARE GROUPED IN SUBROUTINES:
'C
'C   DISTUR: SETS UP INITIAL CONDITIONS AND POST-DISTURBANCE CONDITIONS
'C           FOLLOWING PRUNING, THINNING, AND HARVESTING.
'C   PHOTO:  DETERMINES NET PRIMARY PRODUCTION
'C   DMFMAT: GENERATES FLOW MATRIX, F WHICH PARTITIONS PRIMARY PRODUCTION
'C           TO LIVE TREE COMPONENTS, SIMULATES MORTALITY, AND CALCULATES
'C           LOSSES FROM STAND ON AN ANNUAL BASIS.
'C   DMAMAT: COMPUTES COEFFICIENTS OF RATE MATRIX (A) FROM FLOW MATRIX (F).
'C   EULER:  COMPUTES CURRENT SYSTEM STATE, USING STATE TRANSITION EQUATIONS.
'C   STRUCT: DETERMINES STAND STRUCTURAL FEATURES REQUIRED BY OTHER
'C           SUBROUTINES.
'C   OUTPUT: PRINTS OUT SIMULATION RESULTS.
'C
'C****************************************************************************
'C
'C   PROGRAM MAIN
'C
'C PURPOSE:
'C
'C READS   'INPUT.DAT' AND THEREBY CONTROLS CALL TO ALL SUBROUTINES.
'C   INFORMATION IN THE CONTROL FILE, IS PRINTED TO LP2OUT.DAT WHICH
'C   SHOULD BE CHECKED TO ENSURE THAT THE SIMULATION RUN HAS BEEN CORRECTLY
'C   SPECIFIED.

    Dim TITLE As String, TLAM(12) As Double, A(16, 16) As Double, Y(16) As Double, _
        ZNAP(100) As Double, _
        F128SP(100) As Double, F82SP(100) As Double, SPHSP(100) As Double, WDAGE(100) As Double, NAP As Double, _
        NRUNS As Long, i As Long, j As Long, YESF As Double, YESX As Double, YESA As Double, _
        PARTF As Double, PARTB As Double, PARTS As Double, PARTRC As Double, DENSREG As Double, HTREG As Double, _
        PA As Double, NDIST As Long, DT As Double, ITT As Long, DMLOSI As Double, DMLOSD As Double, _
        POTCL As Double, TNAPAG As Double, DMPNY As Double, DRA As Double, _
        DRB As Double, DRC As Double, DRHA As Double, RADIUS As Double, X8HOLD As Double, ATD As Double, _
        IATD As Long, X8GAIN As Double, AVSSA As Double, AVBSA As Double, AVCRA As Double, _
        MATEMP As Double

    line5 = 1
    line6 = 1
    line7 = 5
    line8 = 5
    line10 = 1
    col5 = 1
    col6 = 1
    col7 = 1
    col8 = 1
    col10 = 1
    Sheets("LP1OUT").Activate
    Range(Cells(1, 1), Cells(64000, 70)).ClearContents
    Sheets("LP2OUT").Activate
    Range(Cells(1, 1), Cells(64000, 70)).ClearContents
    Sheets("Nitrogen").Activate
    Range(Cells(5, 1), Cells(64000, 70)).ClearContents
    Sheets("Phosphorus").Activate
    Range(Cells(5, 1), Cells(64000, 70)).ClearContents

'C
'C     READ CONTROL INFORMATION, TITLE, MATRICES TO BE INPUT OR OUTPUT
'C
    
    Call inout(5, NRUNS)
    Call newline(5)
    Call inout(10, "CARDTYPE A:")
    Call inout(10, NRUNS)
    Call inout(10, "SIMULATION RUNS SPECIFIED")
    Call newline(10)
      
    ICOUNT = 0
1004: ICOUNT = ICOUNT + 1
'C  INITIALISE ARRAYS
    For i = 1 To 100
        TVOLSP(i) = 0
        BASP(i) = 0
        dens = 0        ' is this a bug?
        HTSP(i) = 0
    Next i
    For i = 1 To 16
        For j = 1 To 16
            A(i, j) = 0
            F(i, j) = 0
        Next j
    Next i

    BABGC = 0
    Call inout(5, TITLE)
    Call newline(5)
    Call inout(5, YESF)
    Call inout(5, YESA)
    Call inout(5, YESX)
    Call inout(5, YESN)
    Call inout(5, NLIMITED)
    Call inout(5, INPNAP)
    Call inout(5, INPX)
    Call inout(5, INDEAD)
    Call newline(5)
    
    Call inout(10, "CARDTYPE B: THE TITLE IS")
    Call newline(10)
    Call inout(10, TITLE)
    Call newline(10)
    
    Call inout(10, "CARDTYPE C: MATRICES TO BE OUTPUT. YESF/YESA/YESX")
    Call inout(10, YESF)
    Call inout(10, YESA)
    Call inout(10, YESX)
    Call newline(10)
    
    Call inout(10, "NUTRIENT SIMULATED IF YESN=1.0. YESN= ")
    Call inout(10, YESN)
    Call newline(10)
    
    Call inout(10, "INPNAP/INPX/INDEAD")
    Call inout(10, INPNAP)
    Call inout(10, INPX)
    Call inout(10, INDEAD)
    Call newline(10)
    
    Call inout(10, "GROWTH ADJUSTED WHEN NLIMITED=1. NLIMITED=")
    Call inout(10, NLIMITED)
    Call newline(10)
      
    If NLIMITED = 1 Then
'redundant code        Call inout(5, PARTF)
'redundant code        Call inout(5, PARTB)
'redundant code        Call inout(5, PARTS)
'redundant code        Call inout(5, PARTRC)
'redundant code        Call newline(5)
        
'redundant code        Call inout(10, "PARTITIONING COEF TO FAS,BRCH,STM,ROOTS-COURSE")
'redundant code        Call inout(10, PARTF)
'redundant code        Call inout(10, PARTB)
'redundant code        Call inout(10, PARTS)
'redundant code        Call inout(10, PARTRC)
'redundant code        Call newline(10)
    End If
    
    ID = 1
'C     READ MANAGEMENT INFORMATION : AGE SIMULATION TO END,
'C     INTEGRATION INTERVAL DT
'C    STAND INITIAL HEIGHT,NUMBER OF DISTURBANCES INCLUDING; AGE,
'C    STOCKING,BASAL AREA,PRUNING HEIGHT,AND EXTRACTION INTENTION
      
    Call inout(5, AGEND)
    Call inout(5, DENSREG)
    Call inout(5, HTREG)
    Call inout(5, PA)
    Call inout(5, NDIST)
    Call newline(5)
    
    For i = 1 To 10
        Call inout(5, ADIST(i))
        Call inout(5, SPHAN(i))
        Call inout(5, BAN(i))
        Call inout(5, PRNHT(i))
        Call inout(5, XTRACT(i))
        Call inout(5, XTRACC(i))
        Call inout(5, XTRACF(i))
        Call newline(5)
    Next i
      
    DT = 1
    HT = 0.2
    i = NDIST + 1
    ADIST(i) = 10000
    Call inout(10, "CARDTYPE D: CEASE AT AGE AGEND=")
    Call inout(10, AGEND)
    Call newline(10)
    
    Call inout(10, "DENSITY_REGION/HT/PHYSIOL_AGE/NDIST")
    Call inout(10, DENSREG)
    Call inout(10, HT)
    Call inout(10, PA)
    Call inout(10, NDIST)
    Call newline(10)

    Call inout(10, "CARDTYPE E: ONE CARD FOR EACH OF")
    Call inout(10, NDIST)
    Call inout(10, "DISTURBANCES")
    Call newline(10)
    
    Call inout(10, "ADIST")
    Call inout(10, "SPHAN")
    Call inout(10, "BAN")
    Call inout(10, "PRNHT")
    Call inout(10, "XTRACT")
    Call inout(10, "XTRACC")
    Call inout(10, "XTRACF")
    Call newline(10)
    For i = 1 To NDIST
        Call inout(10, ADIST(i))
        Call inout(10, SPHAN(i))
        Call inout(10, BAN(i))
        Call inout(10, PRNHT(i))
        Call inout(10, XTRACT(i))
        Call inout(10, XTRACC(i))
        Call inout(10, XTRACF(i))
        Call newline(10)
    Next i
    
'C    ANNUALLY PROCEED THROUGH SUBROUTINES 2 TO 7 ADDITIONALLY PASSING
'C    THROUGH SUBROUTINE 1 AT INITIAL TIME AND AT AGE OF EACH
'C    MANAGEMENT OPERATION,AND OPTIONALLY THROUGH NUTRIENT.
'C    PHOTO IS CALLED ONLY ONCE IF ARRAY OF NAP VALUES SUPPLIED AS MEASUREMENT
'C    DATA RATHER THAN BEING SIMULATED
    T = ADIST(1)

'C If DENSITY REGION IS GREATER THAN 4, AND SIMULATING GROWTH, THEN READ AGE
'C AND CORRESPONDING DENSITY OF WHOLE STEM AND ANNUAL INCREMENT BY AGE
'redundant code    If INPNAP = 2 And DENSREG > 4 Then
'redundant code        For i = 1 To 100
'redundant code            Call inout(5, WDAGE(i))
'redundant code            Call inout(5, DENWH(i))
'redundant code            Call inout(5, DENINC(i))
'redundant code        Next i
'redundant code    End If
    
'C ITT IS TIME (AGE) CONVERTED TO ARRAY POSITION -eg AT TIME ZERO IS POSITION 1.
    ITT = Int(T) + 1
    If INPNAP = 2 And DENSREG = 5 Then dens = DENWH(ITT)

'C    READ INITIAL CONDITIONS OR SUBSEQUENT PERTURBATION CONDITIONS
1001: Call DISTUR
'C    FOR FIRST ROTATION, CALCULATE SITE PREP EFFECT ON INITIAL (PREPLANT)
'C    DRY MATTER. SECOND+ ROTATION SHRUBS ALL BECOMES LITTER WITH NO SITE PREP
'C    LOSSES.
    If T = 0 And INPNAP = 2 Then vol = X(8) / DENINC(1)
    If ICOUNT > 1 Then FLOSS = 0
    If T = 0 Then DMLOSI = DMINI * FLOSS
    If T = 0 Then DMLOSD = DMINI - DMLOSI

'C    OUTPUT INITIAL CONDITIONS OR SUBSEQUENT PERTURBATION CONDITIONS
  
      
    Call OUTPUT(A, F, X, Y, NAP, T, ADIST, ID, YESA, YESF, YESX, YESN, TITLE, PRNHT, _
    SPHA, BA, HT, GCL, TLA, AGECK, XTR, XTF, ICOUNT, DMLOSI, DMLOSD, DMINI, dens, _
    TVOL, vol, BABGC, GCFRE, XN, UPNAG, UPPAG, UPNBG, UPPBG, XTRN, XTRP, POTCL, CSHRUB)

'C
'C  IF HARVEST DATE ATTAINED, THEN PROCEED TO END OF CURRENT ROTATION
'C
    If T = AGEND Then GoTo 1000
    ID = ID + 1
    AGECK = 0
1002: If INPNAP = 1 And T <> ADIST(1) Then GoTo 1003

'C    OBTAIN NET ANNUAL PRODUCTION FOR YEAR USING PHOTO
      
    Call PHOTO(INPNAP, NAP, T, LA1Y, LA2Y, LA3Y, RNAPA, RNAPB, RNAPC, SPHA, HTSP, TVOLSP, _
        PNA, PNB, AY1A, AY1B, AY1C, SX3, ADIST, X, TNAPAG, X3MAX, S, B31, B32, GCL, HT, _
        GCLMAX, GCFRE, TLAM, DMPNY, PMAI, DENSREG, DRA, DRB, DRC, CNRAT, DRHA, ZNAP, _
        BASP, F128SP, F82SP, SPHSP, RADIA, DENWH, DENINC)

'C    COMPUTE FLOWS BY COMPARTMENT, AND REALLOCATING EXISTING
'C    STANDING STOCKS AS NECESSARY, THEREBY GENERATING THE FLOW MATRIX, F

    MATEMP = Worksheets("C Change").Cells(40, 3)
    If MATEMP = 0 Then MATEMP = 12
        'Check MAT and if not specified assume default of 12 degrees
    SoilC = Worksheets("C Change").Cells(35, 3)
    SoilN = Worksheets("C Change").Cells(36, 3)
    SoilOrganicP = Worksheets("C Change").Cells(37, 3)
        'Soil variables required for %N and %P component models
    stem_P_reduction = 0
        'Reduction in stem %P from Puruki fertile site standard

1003: Call DMFMAT(X, NAP, F, Y, T, LA1Y, PNA, PNB, pBA, PBB, PBC, CONS, GCL, X3MAX, ID, GCFRE, _
        BA, SPHA, X6MAX, TMA, TMB, VMA, VMB, VMC, VMD, INDEAD, HINC, MAXL1Y, _
        TNAPAG, S, SX3, B31, B32, HT, RADIUS, INPNAP, PA, _
        NLIMITED, PARTF, PARTB, PARTS, PARTRC, ZNAP, F128SP, F82SP, MATEMP)

'C    ADJUST FLOWS WHEN N LIMITED
'    If NLIMITED = 1 Then Call ADFMATN(DIAGN, F, T, NAP, ADJFB, ADJFM)              bug

'C    CALCULATE RATE MATRIX FROM FLOW MATRIX
    Call DMAMAT(A, F, X, Y)

'C    IF N IS BEING SIMULATED STORE X AND XN AS XMN1 AND XNMN1.
    If YESN = 0 Then GoTo 1007
    For i = 1 To 16
        XMN1(i) = X(i)
        XNMN1(i) = XN(i)
        XPMN1(i) = XP(i)
    Next i

'C    PROCEED WITH NUMERICAL INTEGRATION
1007: X8HOLD = X(8)

    Call EULER(A, X, NAP, T)

'C  DENSITY ESTIMATE (T/M**3) BY REGION 1=LOW; 2=MEDIUM; 3=HIGH.
    ATD = T + 1
    If DENSREG <= 3 Then dens = (DRC + DRA * (1 - Exp(DRB * ATD))) / 1000

'C ITT IS TIME (AGE) CONVERTED TO ARRAY POSITION -eg AT TIME ZERO IS POSITION 1.
    IATD = Int(ATD) + 1
    If DENSREG = 5 Then dens = DENWH(IATD)
    X8GAIN = X(8) - X8HOLD

'C    SIMULATE STAND STRUCTURE
      
    Call STRUCT(AVSSA, AVBSA, AVCRA, POTCL, SPHSP, DENSREG, X8GAIN)

'C  OPTIONALLY SIMULATE STAND NUTRIENT CONTENT AND NUTRIENT UPTAKE
    If YESN = 1 Then Call NUTRIENT(F, X, XMN1, XN, XP, XNMN1, XPMN1, UPNAG, UPPAG, UPNBG, UPPBG, T, TBRCH, Y)
    
'C   OUTPUT MATRICES(OPTIONALLY),COMPARTMENT STATES AND STAND STRUCTURE
      
    Call OUTPUT(A, F, X, Y, NAP, T, ADIST, ID, YESA, YESF, YESX, YESN, TITLE, PRNHT, _
    SPHA, BA, HT, GCL, TLA, AGECK, XTR, XTF, ICOUNT, DMLOSI, DMLOSD, DMINI, dens, _
    TVOL, vol, BABGC, GCFRE, XN, UPNAG, UPPAG, UPNBG, UPPBG, XTRN, XTRP, POTCL, CSHRUB)

     
'C    CHECK IF SIMULATION AGE IS EXCEEDED
    If T = AGEND Then GoTo 1001
'C    CHECK IF STAND MANAGEMENT OPERATION INTENDED PRIOR TO NEXT GROWTH
'C SEQUENCE
    If T = ADIST(ID) Then GoTo 1001
    GoTo 1002
1000:  If ICOUNT < NRUNS Then GoTo 1004

End Sub
      

'C******************************************************************************
'C
'C
Sub DISTUR()

'C PURPOSE:
'C
'C    DISTURB READS IN INITIAL CONDITIONS TOGETHER WITH CORRESPONDING
'C    COEFFICIENTS OF PROCESSES UNDER THE DIRECT CONTROL OF MANAGEMENT
'C    ALSO SIMULATES (OR OPTIONALLY WILL READ IN) PERTURBATION STATES
'C    GIVEN THE STAND MANAGEMENT PRESCRIPTION AND AND STAND STRUCTURE
'C    INFORMATION SIMULATED IN SUBROUTINE STRUCTURE
      
    Dim CR(4) As Double, AVERX(8) As Double, IT As Long, XTRFOL As Double, XTRBRC As Double, _
        XTRSTM As Double, IROT As Long, SOILKMG As Double, Cmin As Double, TKjeldN As Double, i As Long, _
        CRNRED As Double, A As Double, B As Double, C As Double, CA As Double, RELCRM As Double, _
        TVX3 As Double, TVX4 As Double, TVX5 As Double, TVX6 As Double, DEADCL As Double, DEADN As Double, DEADP As Double, _
        TVX7 As Double, TVXN7 As Double, TVXP7 As Double, BASET As Double, TVX8 As Double, TVX9 As Double, TVX14 As Double, _
        TVX16 As Double, TVXN3 As Double, TVXP3 As Double, TVXN4 As Double, TVXP4 As Double, TVXN5 As Double, TVXP5 As Double, _
        TVXN6 As Double, TVXP6 As Double, TVXN8 As Double, TVXP8 As Double, TVXN9 As Double, TVXP9 As Double, _
        TVXN14 As Double, TVXP14 As Double, TVXN16 As Double, TVXP16 As Double, XTRFOLN As Double, XTRFOLP As Double, _
        XTRBRN As Double, XTRBRP As Double, XTRSTMN As Double, XTRSTMP As Double, XTF As Double, _
         XTFN As Double, XTFP As Double
    AGECK = 1
    IT = Int(T)
    XTR = 0
    XTF = 0
    XTFN = 0
    XTFP = 0
    XTRFOL = 0
    XTRBRC = 0
    XTRSTM = 0
    XN14THINLOSS = 0
    XP14THINLOSS = 0
    
    If ID <> 1 Then GoTo 101
'C INITIALISE LEAF AREA COEFFICIENTS
    AY1A = 0.009617
    AY1B = -0.08771
    AY1C = 1.7122
'C    ENTER HEIGHT GROWTH COEFFICIENTS, CONSUMPTION, NEEDLE RETENTION, SOLAR
'C    INCIDENT RADIATION, AND ROTATION NUMBER (IROT).
'C    ROTATION 1 REINITIALISES FOREST FLOOR TO ZERO BETWEEN RUNS, HIGHER DOES NOT.
    Call inout(5, B31)
    Call inout(5, B32)
    Call inout(5, S)
    Call inout(5, CONS)
    Call inout(5, SX3)
    Call inout(5, RADIA)
    Call inout(5, IROT)
    Call inout(5, SOILKMG)
    Call inout(5, Cmin)
    Call inout(5, TKjeldN)
    Call newline(5)
    
'C Note that in C_Change (reverse) mode, inputs from eg the 300 Index model in all cases
'C over-write values generated in forward mode given here.
    Call inout(10, "CARDTYPE G: HT GROWTH COEFF:")
    Call inout(10, B31)
    Call inout(10, B32)
    Call inout(10, "HT AGE 20:=")
    Call inout(10, S)
    Call newline(10)
    Call inout(10, "FRACTION CONSUMPTION 1 YR FAS:")
    Call inout(10, CONS)
    Call inout(10, "FRACTION RETENTION OLDER FAS:")
    Call inout(10, SX3)
    Call inout(10, "RADIATION:")
    Call inout(10, RADIA)
    Call inout(10, "ROTATION NO")
    Call inout(10, IROT)
    Call newline(10)
    Call inout(10, "SOIL C=")
    Call inout(10, Cmin)
    Call inout(10, "SOIL Kjeldahl N=")
    Call inout(10, TKjeldN)
    Call newline(10)
    
    HINC = 0
    If PRNHT(1) > 0 Then GCFRE = HT - PRNHT(1)
'C    ASSUME CANOPY AGE EQUAL TO STAND AGE(TRUE IF STAND UNPRUNED INITIALLY)
    TBRCH = T

'C    ENTER COEFFICIENTS DESCRIBING MAXIMUM ONE YEAR NEEDLE MASS AS
'C    A FUNCTION OF STAND AGE.
'C    ALSO INPUT STEM FORM COEFFICIENTS(not required for C_Change), AND BRANCH MAXIMUM WEIGHT
'C    COEFFICIENTS AS A FUNCTION OF STOCKING, AND FASCICLE RETENTION
'C COEFFICIENT
'CP      READ(5,*)AM1,AM2,SFINT,SFSPE,BM1,BM2,SX3
    AM1 = 10.5
    AM2 = -0.0301704
    SFINT = 0.000725
    SFSPE = 0.156
    BM1 = 181.04
    BM2 = 0.3


'C     ENTER INITIAL CONDITIONS and ALSO OPTIONALLY
'C    REINITIALISES STATES AS INPUT DATA (OPTION 1) AFTER EACH
'C    MANAGEMENT OPERATION

'C     ALTERNATIVELY, IF INPX=2 CAN ENTER AVERAGE TREE COMPONENT WEIGHTS
'C     IN KG'S; ZEROS FOR OTHER MODEL COMPARTMENTS IS ASSUMED FOR ROTATION 1.
'C     PREPLANT VEGETATION AND FRACTION LOST IMMEDIATELY IS ALSO GIVEN NOW.
    For i = 1 To 8
        Call inout(5, AVERX(i))
    Next i
    Call inout(5, DMINI)
    Call inout(5, FLOSS)
    
'C If in second+ rotation, DMINI is shrubs CARBON X 2 at end of previous
'C rotation, and FLOSS will be assumed to be zero.
      
    If ICOUNT > 1 Then DMINI = CSHRUB * 2
    Call inout(10, "CRDTYPE K: AVERAGE TREE COMPONENT WEIGHTS(KG),STARTING AT 'X3':=")
    Call newline(10)
    For i = 1 To 8
        Call inout(10, AVERX(i))
    Next i
    Call newline(10)
    
    Call inout(10, "INITIAL VEGETATION DM, AND SITE PRP LOSS, RESP")
    Call inout(10, DMINI)
    Call inout(10, FLOSS)
    Call newline(10)
    For i = 1 To 7
        X(i + 2) = AVERX(i) * SPHAN(1) / 1000
    Next i
    X(14) = AVERX(8) * SPHAN(1) / 1000
    X(8) = AVERX(6) * SPHAN(1) * 0.75 / 1000
    X(16) = AVERX(6) * SPHAN(1) * 0.25 / 1000
    X(1) = 0
    X(2) = 0
    
'C  Calculate nitrogen content at start of rotation
    If YESN = 1 Then
        For i = 3 To 9
            XN(i) = X(i) * Nconc(i, 0) * 10
            XP(i) = X(i) * Pconc(i, 0) * 10
        Next i
        XN(14) = X(14) * Nconc(14, 0) * 10
        XN(16) = X(16) * Nconc(16, 0) * 10
        XP(14) = X(14) * Pconc(14, 0) * 10
        XP(16) = X(16) * Pconc(16, 0) * 10
'        XN(8) = AVERX(6) * SPHAN(1) * 0.75 / 1000
'        XN(16) = AVERX(6) * SPHAN(1) * 0.25 / 1000
        XN(1) = 0
        XN(2) = 0
        XTRN = 0
        XP(1) = 0
        XP(2) = 0
        XTRP = 0
    End If
    
'C  SET FOREST FLOOR WEIGHTS TO SPECIFIED INITIAL CONDITIONS IF A FIRST ROTATION SITE
'C  FOREST FLOOR MASS IN SECOND ROTATION IS BASED ON PREVIOUS RUN
    If IROT = 1 Then Call inout(5, X(10))
    If IROT = 1 Then Call inout(5, X(11))
    If IROT = 1 Then Call inout(5, X(12))
    If IROT = 1 Then Call inout(5, X(13))
    If IROT = 1 Then Call inout(5, X(15))
    
'C  Set nitrogen & phosphorous contents of litter pools to specified intial conditions if first rotation
    If IROT = 1 Then Call inout(5, XN(10))
    If IROT = 1 Then Call inout(5, XN(11))
    If IROT = 1 Then Call inout(5, XN(12))
    If IROT = 1 Then Call inout(5, XN(13))
    If IROT = 1 Then Call inout(5, XN(15))
    If IROT = 1 Then Call inout(5, XP(10))
    If IROT = 1 Then Call inout(5, XP(11))
    If IROT = 1 Then Call inout(5, XP(12))
    If IROT = 1 Then Call inout(5, XP(13))
    If IROT = 1 Then Call inout(5, XP(15))
   
    Call newline(5)

'C    NUTRIENT COMPARTMENTS(KG/HA)
'C      IF(YESN.EQ.0.0)GO TO 230
'C
    If X(1) = 0 Then X(1) = 5
    If X(2) = 0 Then X(2) = 5
    If PRNHT(ID) <> -1 Then HTBGC = PRNHT(ID)
    If PRNHT(ID) <> -1 Then GCL = HT - PRNHT(ID)
    If SPHAN(ID) <> -1 Then SPHA = SPHAN(ID)
    If BAN(ID) <> -1 Then BA = BAN(ID)
    LA1Y = (AY1A * X(3) ^ 2 + AY1B * X(3) + AY1C) * X(3)
    LA2Y = (AY1A * X(4) ^ 2 + AY1B * X(4) + AY1C) * X(4)
    LA3Y = (AY1A * X(5) ^ 2 + AY1B * X(5) + AY1C) * X(5)
    TLA = LA1Y + LA2Y + LA3Y
    X3MAX = AM1 * Exp(AM2 * (T + 1))
    MAXL1Y = (AY1A * X3MAX ^ 2 + AY1B * X3MAX + AY1C) * X3MAX
    X6MAX = BM1 * (1 / SPHA) ^ BM2

104: If SPHA = 0 Then HT = 0
    If SPHA = 0 Then GCL = 0
    If SPHA = 0 Then vol = 0
    If SPHA = 0 Then Exit Sub

'C REPLACE VOL WITH GROWTH MODEL (300 INDEX) VOLUMES - REMEMBER TO ALLOW FOR MULTIPLE LINES
    If TVOLSP(IT + ID) > 0 Then vol = TVOLSP(IT + ID)
703: Exit Sub

'C    OPTION 2
'C    REINITIALISE STATES BY SIMULATION, FOLLOWING EACH MANAGEMENT OPERATION.
101:
'C
'C    EFFECTS OF PRUNING ON CROWN AND LITTER COMPARTMENTS
    If PRNHT(ID) = -1 Then GoTo 103
'C    LENGTH OF GREEN CROWN REMOVED
    CRNRED = PRNHT(ID) - HTBGC
'CP      READ(5,*)A,B,C
    A = 1.072
    B = -3.747
    C = 2.909
 
'CP      WRITE(10,41)A,B,C
'CP 41   FORMAT(40H CARDTYPE L: CROWN VERT. DISTBN COEFFS:=,/3F10.6)
    If CRNRED <= 0 Then GoTo 102
    If GCLMAX = 0 Then GCLMAX = HT
    If T <= 6 Then GCLMAX = HT
    GCL = GCL - CRNRED
    If GCL < GCFRE Then GCFRE = GCL
'C    CROWN AGE (ASSUME CONSTANT HT INCREMENT OVER THIS AGE)
    TBRCH = GCL / HINC
'C    RELATIVE CROWN LENGTH REMAINING, RELCRM
    CA = 0
    For i = 1 To 3
        RELCRM = (GCL - HINC * CA) / GCLMAX
        CA = CA + 1
'C    PREDICT FRACTION OF  CROWN REMAINING AFTER PRUNING BUT PRIOR
'C    TO THINNING
        If RELCRM > 0 Then GoTo 219
        CR(i) = 0
        GoTo 220
219:    CR(i) = A * (1 - Exp(B * RELCRM)) ^ C
220: Next i
    CR(4) = (GCL / GCLMAX) ^ 2

'C    NEW INITIAL CONDITIONS AFTER PRUNING
'C    DRY MATTER COMPARTMENTS(T/HA)
    TVX3 = X(3) * CR(1)
    TVX4 = X(4) * CR(2)
    TVX5 = X(5) * CR(3)
    TVX6 = X(6) * CR(4)
'C RESET LEAF AREA COEFFICIENTS FOR PRUNED STAND
'C AY1A = 0
'C AY1A = 0.04134
'C AY1A = 1.2988

    LA1Y = (AY1A * TVX3 ^ 2 + AY1B * TVX3 + AY1C) * TVX3
    LA2Y = (AY1A * TVX4 ^ 2 + AY1B * TVX4 + AY1C) * TVX4
    LA3Y = (AY1A * TVX5 ^ 2 + AY1B * TVX5 + AY1C) * TVX5
    TLA = LA1Y + LA2Y + LA3Y

    X(10) = X(10) + X(3) + X(4) + X(5) - (TVX3 + TVX4 + TVX5)
    X(11) = X(11) + X(7) + X(6) - TVX6
    X(3) = TVX3
    X(4) = TVX4
    X(5) = TVX5
    X(6) = TVX6
    X(7) = 0

'C    NUTRIENT COMPARTMENTS(KG/HA)
    If YESN = 0 Then GoTo 109
    TVXN3 = TVX3 * Nconc(3, T) * 10
    TVXN4 = TVX4 * Nconc(4, T) * 10
    TVXN5 = TVX5 * Nconc(5, T) * 10
    TVXN6 = TVX6 * Nconc(6, T) * 10
    XN(10) = XN(10) + XN(3) + XN(4) + XN(5) - (TVXN3 + TVXN4 + TVXN5)
    XN(11) = XN(11) + XN(7) + XN(6) - TVXN6
    XN(3) = TVXN3
    XN(4) = TVXN4
    XN(5) = TVXN5
    XN(6) = TVXN6
    XN(7) = 0
    
    TVXP3 = TVX3 * Pconc(3, T) * 10
    TVXP4 = TVX4 * Pconc(4, T) * 10
    TVXP5 = TVX5 * Pconc(5, T) * 10
    TVXP6 = TVX6 * Pconc(6, T) * 10
    XP(10) = XP(10) + XP(3) + XP(4) + XP(5) - (TVXP3 + TVXP4 + TVXP5)
    XP(11) = XP(11) + XP(7) + XP(6) - TVXP6
    XP(3) = TVXP3
    XP(4) = TVXP4
    XP(5) = TVXP5
    XP(6) = TVXP6
    XP(7) = 0
      
    GoTo 109
'C
'C    EFFECTS OF GREEN CROWN PRUNING ON SYSTEM COMPARTMENTS COMPLETED.
'C
'C    ASSUMING THAT DEAD BRANCH/M STEM IS CONSTANT WITH HEIGHT,
'C    EFFECT OF DEAD ZONE PRUNING ON DEAD BRANCH AND LITTER
'C     COMPARTMENTS IS :
102: If HTBGC = PRNHT(ID - 1) Then GoTo 109
    DEADCL = X(7) / (HTBGC - PRNHT(ID - 1))
    If YESN = 1 Then DEADN = XN(7) / (HTBGC - PRNHT(ID - 1))
    If YESN = 1 Then DEADP = XP(7) / (HTBGC - PRNHT(ID - 1))
    If CRNRED < 0 Then GoTo 107
    TVX7 = 0
    TVXN7 = 0
    TVXP7 = 0
    GoTo 108
107: TVX7 = DEADCL * (HTBGC - PRNHT(ID))
    If YESN = 1 Then TVXN7 = DEADN * (HTBGC - PRNHT(ID))
    If YESN = 1 Then TVXP7 = DEADP * (HTBGC - PRNHT(ID))
108: X(11) = X(11) + X(7) - TVX7
    X(7) = TVX7
    If YESN = 1 Then GoTo 109
    XN(11) = XN(11) + XN(7) - TVXN7
    XN(7) = TVXN7
    XP(11) = XP(11) + XP(7) - TVXP7
    XP(7) = TVXP7
109: HTBGC = HT - GCL

'C    EFFECTS OF PRUNING ON SYSTEM COMPARTMENTS COMPLETED
'C
'C    EFFECTS OF THINNING ON CROWN, STEM AND LITTER COMPARTMENTS
'C
'C    STANDING COMPARTMENTS ARE REDUCED BY THINNING IN DIRECT
'C    PROPORTION TO THE REDUCTION IN STAND BASAL AREA OR PREFERRABLY INPUT VOLUME.
'C    IF A PRODUCTION THINNING, A SPECIFIED FRACTION OF STEM MATTER IS HARVESTED,
'C    AND A SPECIFIED FRACTION OF CROWN MATTER ARE ADDED TO LITTER COMPARTMENTS
103: If BAN(ID) = -1 Then GoTo 104
    If BAN(ID) > BA Then
        Call inout(10, "BASAL AREA AFTER THINNING EXCEEDS SIMULATED VALUE AT THINNING AGE:")
        Call inout(10, ADIST(ID))
        Call newline(10)
    End If

'C  CLEARFELLING REDUCES BASAL AREA AND HENCE ALL TREE COMPARTMENTS TO ZERO
    If SPHA = 0 Then BAN(ID) = 0
'C COULD WARN HERE IF THINNING SPECIFIED (BAN = -2.0, BUT SPHA IS SAME.
    If SPHA > 0 And BAN(ID) = -2 Then BAN(ID) = (BA / SPHA) * SPHAN(ID) * 1.2

'C NOTE THAT THINNING EFFECT ON POOLS IS PROPORTIONAL TO VOLUME (NOT BA)
    If TVOLSP(IT) > 0 Then BAN(ID) = BA * (TVOLSP(IT + ID) / TVOLSP(IT + ID - 1))

    BASET = 0
    If BA > 0 Then GoTo 2345
'redundant code    BA = SPHA
'redundant code    BAN(ID) = SPHAN(ID) * 1.2
'redundant code    BASET = 1
'C    TREE COMPARTMENTS
2345: TVX3 = (X(3) / BA) * BAN(ID)
    TVX4 = (X(4) / BA) * BAN(ID)
    TVX5 = (X(5) / BA) * BAN(ID)
    TVX6 = (X(6) / BA) * BAN(ID)
    TVX7 = (X(7) / BA) * BAN(ID)
    TVX8 = (X(8) / BA) * BAN(ID)
    TVX9 = (X(9) / BA) * BAN(ID)
    TVX14 = (X(14) / BA) * BAN(ID)
    TVX16 = (X(16) / BA) * BAN(ID)
    vol = (vol / BA) * BAN(ID)
'C
'C    FOREST FLOOR COMPARTMENTS
'C    NOTE: HARVESTED MATTER DOES NOT ENTER FOREST FLOOR LITTER COMPARTMENTS
'C    HARVESTED MATTER DEPENDS ON INPUT PARAMETERS, XTRACC (FRACTION CROWN REMOVED), XTRACT (FRACTION STEM REMOVED),
'C    XTRACF (FRACTION FOREST FLOOR REMOVED)
'C    If XTRACC = -1, THEN ALL FOLIAGE SLASH IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT

    If XTRACC(ID) > 0 Then XTRFOL = (X(3) + X(4) + X(5) - (TVX3 + TVX4 + TVX5)) * XTRACC(ID)
    If XTRACC(ID) = -1 Then XTRFOL = (X(3) + X(4) + X(5) - (TVX3 + TVX4 + TVX5)) * XTRACC(ID)
    X(10) = X(10) + X(3) + X(4) + X(5) - (TVX3 + TVX4 + TVX5 + XTRFOL)

'C    FOR BRANCHES, ASSUME LIVE PLUS DEAD MATTER REMOVED IN SAME PROPORTION FOR EXTRACTION PURPOSES
'C    STUMPS ARE LEFT ON-SITE IF XTRACT AT CLEARFELL IS SET TO 0.85 (SET THIS HIGHER TO EXTRACT AG STUMP)
'C    ROOTS ALL RETAINED ON SITE (TO SPECIFY ROOT STOCK REMOVAL, WILL NEED TO SPECIFY THE FRACTION OF COARSE ROOT SYSTEM REMOVED AND ADD CODE!)
    If XTRACC(ID) > 0 Then XTRBRC = (X(6) + X(7) - (TVX6 + TVX7)) * XTRACC(ID)
'C    If XTRACC = -1, THEN ALL BRANCH SLASH IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT
    If XTRACC(ID) = -1 Then XTRBRC = (X(6) + X(7) - (TVX6 + TVX7)) * XTRACC(ID)
    X(11) = X(11) + X(6) + X(7) - (TVX6 + TVX7 + XTRBRC)
    If XTRACT(ID) > 0 Then XTRSTM = (X(8) + X(16) - TVX8 - TVX16) * XTRACT(ID)
    X(12) = X(12) + X(8) + X(16) - (TVX8 + TVX16 + XTRSTM)
    X(13) = X(13) + X(9) - TVX9
    X(15) = X(15) + X(14) - TVX14
'C    SLASH REMOVED DURING HARVESTING IS SUM OF EXTRACTED CROWN AND STEM DRY MATTER
    XTR = XTRFOL + XTRBRC + XTRSTM
'C    FOREST FLOOR REMOVED DURING SITE PREPARATION CONTAINS ORIGINAL FF PLUS HARVEST RESIDUES ON FF
    If XTRACF(ID) > 0 Then XTF = (X(10) + X(11) + X(12)) * XTRACF(ID)
'C    FF DM POOLS AFTER HARVESTING AND SITE PREPARATION WITH FF EXTRACTION SPECIFIED
    If XTRACF(ID) > 0 Then X(10) = X(10) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then X(11) = X(11) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then X(12) = X(12) * (1 - XTRACF(ID))
    
    LA1Y = (LA1Y / BA) * BAN(ID)
    LA2Y = (LA2Y / BA) * BAN(ID)
    LA3Y = (LA3Y / BA) * BAN(ID)
    TLA = LA1Y + LA2Y + LA3Y

'C     TREE COMPARTMENTS

    X(3) = TVX3
    X(4) = TVX4
    X(5) = TVX5
    X(6) = TVX6
    X(7) = TVX7
    X(8) = TVX8
    X(9) = TVX9
    X(14) = TVX14
    X(16) = TVX16

'C    NUTRIENT COMPARTMENTS
    If YESN = 0 Then GoTo 110
    TVXN3 = (XN(3) / BA) * BAN(ID)
    TVXN4 = (XN(4) / BA) * BAN(ID)
    TVXN5 = (XN(5) / BA) * BAN(ID)
    TVXN6 = (XN(6) / BA) * BAN(ID)
    TVXN7 = (XN(7) / BA) * BAN(ID)
    TVXN8 = (XN(8) / BA) * BAN(ID)
    TVXN9 = (XN(9) / BA) * BAN(ID)
    TVXN14 = (XN(14) / BA) * BAN(ID)
    TVXN16 = (XN(16) / BA) * BAN(ID)
    
    TVXP3 = (XP(3) / BA) * BAN(ID)
    TVXP4 = (XP(4) / BA) * BAN(ID)
    TVXP5 = (XP(5) / BA) * BAN(ID)
    TVXP6 = (XP(6) / BA) * BAN(ID)
    TVXP7 = (XP(7) / BA) * BAN(ID)
    TVXP8 = (XP(8) / BA) * BAN(ID)
    TVXP9 = (XP(9) / BA) * BAN(ID)
    TVXP14 = (XP(14) / BA) * BAN(ID)
    TVXP16 = (XP(16) / BA) * BAN(ID)
    
    If XTRACC(ID) > 0 Then XTRFOLN = (XN(3) + XN(4) + XN(5) - (TVXN3 + TVXN4 + TVXN5)) * XTRACC(ID)
    If XTRACC(ID) > 0 Then XTRFOLP = (XP(3) + XP(4) + XP(5) - (TVXP3 + TVXP4 + TVXP5)) * XTRACC(ID)
    
'C    If XTRACC = -1, THEN ALL FOLIAGE SLASH NITROGEN IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT
    If XTRACC(ID) = -1 Then XTRFOLN = (XN(3) + XN(4) + XN(5) - (TVXN3 + TVXN4 + TVXN5)) * XTRACC(ID)
 'C    If XTRACC = -1, THEN ALL FOLIAGE SLASH PHOSPHORUS IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT
    If XTRACC(ID) = -1 Then XTRFOLP = (XP(3) + XP(4) + XP(5) - (TVXP3 + TVXP4 + TVXP5)) * XTRACC(ID)

    XN(10) = XN(10) + XN(3) + XN(4) + XN(5) - (TVXN3 + TVXN4 + TVXN5 + XTRFOLN)
    XP(10) = XP(10) + XP(3) + XP(4) + XP(5) - (TVXP3 + TVXP4 + TVXP5 + XTRFOLP)
    If XTRACC(ID) > 0 Then XTRBRN = (XN(6) + XN(7) - (TVXN6 + TVXN7)) * XTRACC(ID)
    If XTRACC(ID) > 0 Then XTRBRP = (XP(6) + XP(7) - (TVXP6 + TVXP7)) * XTRACC(ID)
'C    If XTRACC = -1, THEN ALL BRANCH SLASH NITROGEN IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT
    If XTRACC(ID) = -1 Then XTRBRN = (XN(6) + XN(7) - (TVXN6 + TVXN7)) * XTRACC(ID)
'C    If XTRACC = -1, THEN ALL BRANCH SLASH PHOSPHORUS IS RETAINED ON SITE - AS IN LTSP1 DOUBLE SLASH TREATMENT
    If XTRACC(ID) = -1 Then XTRBRP = (XP(6) + XP(7) - (TVXP6 + TVXP7)) * XTRACC(ID)
    XN(11) = XN(11) + XN(6) + XN(7) - (TVXN6 + TVXN7 + XTRBRN)
    XP(11) = XP(11) + XP(6) + XP(7) - (TVXP6 + TVXP7 + XTRBRP)
    If XTRACT(ID) > 0 Then XTRSTMN = (XN(8) + XN(16) - TVXN8 - TVXN16) * XTRACT(ID)
    If XTRACT(ID) > 0 Then XTRSTMP = (XP(8) + XP(16) - TVXP8 - TVXP16) * XTRACT(ID)
    XN(12) = XN(12) + XN(8) + XN(16) - (TVXN8 + TVXN16 + XTRSTMN)
    XN(13) = XN(13) + XN(9) - TVXN9
    XP(12) = XP(12) + XP(8) + XP(16) - (TVXP8 + TVXP16 + XTRSTMP)
    XP(13) = XP(13) + XP(9) - TVXP9
'C    EXTRACT FOREST FLOOR NUTRIENTS
'C    FOREST FLOOR N AND P REMOVED, AND RETAINED ON SITE
    If XTRACF(ID) > 0 Then XTFN = (XN(10) + XN(11) + XN(12)) * XTRACF(ID)
    If XTRACF(ID) > 0 Then XN(10) = XN(10) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then XN(11) = XN(11) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then XN(12) = XN(12) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then XTFP = (XP(10) + XP(11) + XP(12)) * XTRACF(ID)
    If XTRACF(ID) > 0 Then XP(10) = XP(10) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then XP(11) = XP(11) * (1 - XTRACF(ID))
    If XTRACF(ID) > 0 Then XP(12) = XP(12) * (1 - XTRACF(ID))
    
'C    SUM POOLS EXTRACTED WHEN HARVESTING
    XTRN = XTRFOLN + XTRBRN + XTRSTMN + XTFN
    XTRP = XTRFOLP + XTRBRP + XTRSTMP + XTFP
    
    XN14THINLOSS = XN(14) - TVXN14
    XN(3) = TVXN3
    XN(4) = TVXN4
    XN(5) = TVXN5
    XN(6) = TVXN6
    XN(7) = TVXN7
    XN(8) = TVXN8
    XN(9) = TVXN9
    XN(14) = TVXN14
    XN(16) = TVXN16

    XP14THINLOSS = XP(14) - TVXP14
    XP(3) = TVXP3
    XP(4) = TVXP4
    XP(5) = TVXP5
    XP(6) = TVXP6
    XP(7) = TVXP7
    XP(8) = TVXP8
    XP(9) = TVXP9
    XP(14) = TVXP14
    XP(16) = TVXP16

110: SPHA = SPHAN(ID)
    BA = BAN(ID)
    If BASET = 1 Then BA = 0
    If BASET = 1 Then BAN(ID) = 0
    GoTo 104

End Sub



'C******************************************************************************

Sub EULER(A, X, NAP, T)

'C
'C PURPOSE:
'C
'C  CALCULATES NEW STATE OF SYSTEM
'C
'C  STATEMENT FUNCTIONS FOR STATE TRANSITION EQUATIONS

    Dim j As Long, z As Double, DT As Double, NSOL As Long, DX1 As Double, DX2 As Double, DX3 As Double, _
        DX4 As Double, DX5 As Double, DX6 As Double, DX7 As Double, DX8 As Double, DX9 As Double, DX10 As Double, _
        DX11 As Double, DX12 As Double, DX13 As Double, DX14 As Double, DX15 As Double, DX16 As Double

    j = Int(T)
    z = NAP * 3
    If DT = 0 Then DT = 1
    NSOL = 1 / DT
    For j = 1 To NSOL
'C  BEGIN NUMERICAL INTEGRATION
      DX1 = DT * (z + A(1, 1) * X(1))
      DX2 = DT * (A(2, 1) * X(1) + A(2, 2) * X(2))
      DX3 = DT * (A(3, 2) * X(2) + A(3, 3) * X(3))
      DX4 = DT * (A(4, 3) * X(3) + A(4, 4) * X(4))
      DX5 = DT * (A(5, 4) * X(4) + A(5, 5) * X(5))
      DX6 = DT * (A(6, 2) * X(2) + A(6, 6) * X(6))
      DX7 = DT * (A(7, 6) * X(6) + A(7, 7) * X(7))
      DX8 = DT * (A(8, 2) * X(2) + A(8, 8) * X(8))
      DX9 = DT * (A(9, 2) * X(2) + A(9, 9) * X(9))
      DX10 = DT * (A(10, 3) * X(3) + A(10, 4) * X(4) + A(10, 5) * X(5) + A(10, 10) * X(10))
      DX11 = DT * (A(11, 7) * X(7) + A(11, 11) * X(11))
      DX12 = DT * (A(12, 8) * X(8) + A(12, 16) * X(16) + A(12, 12) * X(12))
      DX13 = DT * (A(13, 9) * X(9) + A(13, 13) * X(13))
      DX14 = DT * (A(14, 2) * X(2) + A(14, 14) * X(14))
      DX15 = DT * (A(15, 14) * X(14) + A(15, 15) * X(15))
      DX16 = DT * (A(16, 2) * X(2) + A(16, 16) * X(16))
      X(1) = X(1) + DX1
      X(2) = X(2) + DX2
      X(3) = X(3) + DX3
      X(4) = X(4) + DX4
      X(5) = X(5) + DX5
      X(6) = X(6) + DX6
      X(7) = X(7) + DX7
      X(8) = X(8) + DX8
      X(9) = X(9) + DX9
      X(10) = X(10) + DX10
      X(11) = X(11) + DX11
      X(12) = X(12) + DX12
      X(13) = X(13) + DX13
      X(14) = X(14) + DX14
      X(15) = X(15) + DX15
      X(16) = X(16) + DX16
    Next j
'C  END NUMERICAL INTEGRATION
      
End Sub

'C******************************************************************************

Sub OUTPUT(A, F, X, Y, NAP, T, ADIST, ID, YESA, YESF, YESX, YESN, TITLE, PRNHT, _
    SPHA, BA, HT, GCL, TLA, AGECK, XTR, XTF, ICOUNT, DMLOSI, DMLOSD, DMINI, dens, _
    TVOL, vol, BABGC, GCFRE, XN, UPNAG, UPPAG, UPNBG, UPPBG, XTRN, XTRP, POTCL, CSHRUB)

'C
'C PURPOSE:
'C
'C  CONVERTS DRY MATTER TO CARBON, ASSUMING 50% CARBON CONTENT.
'C  SUMMARISES COMPARTMENT DATA AND OUTPUTS SIMULATION RESULTS.
'C  CURRENTLY, UNDERSTOREY DEVELOPMENT, BASED ON A SIMPLE EMPIRICAL EQUATION,
'C  IS ALSO SIMULATED IN THIS SUBROUTINE.

    Dim CX(16) As Double, VERS As Double, TMP1 As Double, TMP2 As Double, DMDEF As Double, CTREES As Double, _
        PRVEGR As Double, IS_ As Long, CSTAND As Double, CXTR As Double, CFAS As Double, CROOTL As Double, _
        CROOTD As Double, CSTEM As Double, XTMP1 As Double, XTMP2 As Double, i As Long, j As Long


'C SPECIFY VERSION NUMBER HERE AS VERS
      
    VERS = 2

'C  Differs from version 1.01 only in backward SP run mode now (dim 100)
'C  with gross/net volume input data.

'C  PREDICT SHRUB AND HERB CARBON CONTENT (ABOVE AND BELOW GROUND) IN
'C  KAINGAROA FOREST AS A FUNCTION OF RADIATA PINE STAND AGE. (OLIVER
'C  AND BULLOCK 1986, PROJECT RECORD NO: 1174

'C  PREPLANT VEGETATION REMAINING AT TIME T IS MINUS THAT LOST FROM SYSTEM
'C  DURING SITE PREP. PLUS ANY REMAINING AFTER DELAYED LOSS (DMLOSI, DMLOSD),
'C  ACCORDING TO DEFAULT DECOMPOSITION RATE (USE BRCH = O.18)
'C IF ZERO PREPLANT VEG SPECIFIED, A DEFAULT VALUE = DMDEF IS USED.
    TMP1 = 0
    TMP2 = 0
    DMDEF = 0
    If DMINI > 0 Then DMDEF = 0
    CTREES = 0
    PRVEGR = DMLOSD * Exp(-0.18 * T)
    CSHRUB = (PRVEGR + DMDEF + 0.1778 * T + 0.0064 * T ^ 2) / 2

'C  CONVERT RADIATA PINE DRY MATTER TO CARBON (ASSUME 50% C CONTENT).
    For IS_ = 3 To 16
        CX(IS_) = X(IS_) / 2
        CTREES = CTREES + CX(IS_)
    Next IS_
    
    CSTAND = CTREES + CSHRUB

    CXTR = XTR / 2

'C  SUMMARISE CARBON COMPONENTS FOR OUTPUT
    CFAS = CX(3) + CX(4) + CX(5)
    CROOTL = CX(9) + CX(14)
    CROOTD = CX(13) + CX(15)
    CSTEM = CX(8) + CX(16)
    If T > ADIST(1) Then GoTo 110
    If ICOUNT > 1 Then GoTo 110
    If ICOUNT = 1 Then Call newline(6)
    If ICOUNT = 1 Then
        Call inout(10, "    DRYMAT  VERS")
        Call inout(10, VERS)
        Call newline(10)
    End If
    
'C      IF(ICOUNT.EQ.1)WRITE(8,30)VERS
      
    Call inout(10, TITLE)
    Call newline(10)
    Call inout(6, "AGE")
    Call inout(6, "X1")
    Call inout(6, "X2")
    Call inout(6, "X3")
    Call inout(6, "X4")
    Call inout(6, "X5")
    Call inout(6, "X6")
    Call inout(6, "X7")
    Call inout(6, "X8")
    Call inout(6, "X9")
    Call inout(6, "X10")
    Call inout(6, "X11")
    Call inout(6, "X12")
    Call inout(6, "X13")
    Call inout(6, "X14")
    Call inout(6, "X15")
    Call inout(6, "X16")
    Call inout(6, "SPH")
    Call inout(6, "BA")
    Call inout(6, "HT")
    Call inout(6, "GCL")
    Call inout(6, "VOL")
    Call inout(6, "DENS")
    Call inout(6, "LAI")
    Call inout(6, "PRNHT")
    Call inout(6, "EXT")
    Call inout(6, "CSHRUB")
    Call newline(6)

110: If AGECK = 0 Then GoTo 120
    If YESX = 1 Then
        Call inout(6, T)
        Call inout(6, XTMP1)
        Call inout(6, XTMP2)
        For i = 3 To 16
            Call inout(6, X(i))
        Next i
        Call inout(6, SPHA)
        Call inout(6, BA)
        Call inout(6, HT)
        Call inout(6, GCL)
        Call inout(6, vol)
        Call inout(6, dens)
        Call inout(6, TLA)
        Call inout(6, PRNHT(ID))
        Call inout(6, XTR)
        Call inout(6, CSHRUB)
        Call newline(6)
    End If
    
'       IF(YESN.EQ.1.0)WRITE(6,17)T,(XN(I),I=3,15),XTRN
       If YESN = 1 Then
        Call inout(7, T)
        Call inout(7, XN(3))
        Call inout(7, XN(4))
        Call inout(7, XN(5))
        Call inout(7, XN(6))
        Call inout(7, XN(7))
        Call inout(7, XN(8))
        Call inout(7, XN(9))
        Call inout(7, XN(10))
        Call inout(7, XN(11))
        Call inout(7, XN(12))
        Call inout(7, XN(13))
        Call inout(7, XN(14))
        Call inout(7, XN(15))
        Call inout(7, XN(16))
        Call inout(7, XN(3) + XN(4) + XN(5) + XN(6) + XN(7) + XN(8) + XN(9) + XN(14) + XN(16))
        Call inout(7, XN(10) + XN(11) + XN(12) + XN(13))
        Call inout(7, XN(3) + XN(4) + XN(5) + XN(6) + XN(7) + XN(8) + XN(9) + XN(14) + XN(16) + XN(10) + XN(11) + XN(12) + XN(13))
        Call inout(7, 0)
        Call inout(7, 0)
        Call inout(7, 0)
        Call inout(7, 0)
        Call inout(7, -XN14THINLOSS)
        Call inout(7, XTRN)
        Call newline(7)
    
        Call inout(8, T)
        Call inout(8, XP(3))
        Call inout(8, XP(4))
        Call inout(8, XP(5))
        Call inout(8, XP(6))
        Call inout(8, XP(7))
        Call inout(8, XP(8))
        Call inout(8, XP(9))
        Call inout(8, XP(10))
        Call inout(8, XP(11))
        Call inout(8, XP(12))
        Call inout(8, XP(13))
        Call inout(8, XP(14))
        Call inout(8, XP(15))
        Call inout(8, XP(16))
        Call inout(8, XP(3) + XP(4) + XP(5) + XP(6) + XP(7) + XP(8) + XP(9) + XP(14) + XP(16))
        Call inout(8, XP(10) + XP(11) + XP(12) + XP(13))
        Call inout(8, XP(3) + XP(4) + XP(5) + XP(6) + XP(7) + XP(8) + XP(9) + XP(14) + XP(16) + XP(10) + XP(11) + XP(12) + XP(13))
        Call inout(8, 0)
        Call inout(8, 0)
        Call inout(8, 0)
        Call inout(8, 0)
        Call inout(8, -XP14THINLOSS)
        Call inout(8, XTRP)
        Call newline(8)
    End If
    
'C      IF(YESX.EQ.1.0)WRITE(8,27)T,CSTAND,CSHRUB,CFAS,(CX(I),I=6,7),
'C     1CSTEM,CROOTL,CX(10),CX(11),CX(12),CROOTD,SPHA,BA,
'C     1HT,GCL,VOL,TLA,PRNHT(ID),CXTR
      
    Exit Sub
120: If YESX = 1 Then
        Call inout(6, T)
        For i = 1 To 16
            Call inout(6, X(i))
        Next i
        Call inout(6, SPHA)
        Call inout(6, BA)
        Call inout(6, HT)
        Call inout(6, GCL)
        Call inout(6, vol)
        Call inout(6, dens)
        Call inout(6, TLA)
        Call inout(6, TMP1)
        Call inout(6, TMP2)
        Call inout(6, CSHRUB)
        Call newline(6)
    End If
    
'C      IF(YESX.EQ.1.0)WRITE(8,27)T,CSTAND,CSHRUB,CFAS,(CX(I),I=6,7),
'C     1CSTEM,CROOTL,CX(10),CX(11),CX(12),CROOTD,SPHA,BA,
'C     1HT,GCL,VOL,TLA
'      IF(YESN.EQ.1.0)WRITE(6,7)T,UPNAG,UPNBG,(XN(I),I=3,15)
        
    If YESN = 1 Then
        Call inout(7, T)
        Call inout(7, XN(3))
        Call inout(7, XN(4))
        Call inout(7, XN(5))
        Call inout(7, XN(6))
        Call inout(7, XN(7))
        Call inout(7, XN(8))
        Call inout(7, XN(9))
        Call inout(7, XN(10))
        Call inout(7, XN(11))
        Call inout(7, XN(12))
        Call inout(7, XN(13))
        Call inout(7, XN(14))
        Call inout(7, XN(15))
        Call inout(7, XN(16))
        Call inout(7, XN(3) + XN(4) + XN(5) + XN(6) + XN(7) + XN(8) + XN(9) + XN(14) + XN(16))
        Call inout(7, XN(10) + XN(11) + XN(12) + XN(13))
        Call inout(7, XN(3) + XN(4) + XN(5) + XN(6) + XN(7) + XN(8) + XN(9) + XN(14) + XN(16) + XN(10) + XN(11) + XN(12) + XN(13))
        Call inout(7, UPNAG)
        Call inout(7, UPNBG)
        Call inout(7, UPNFF)
        Call inout(7, DecayN)
        Call inout(7, DeficitN)
        Call newline(7)
    
        Call inout(8, T)
        Call inout(8, XP(3))
        Call inout(8, XP(4))
        Call inout(8, XP(5))
        Call inout(8, XP(6))
        Call inout(8, XP(7))
        Call inout(8, XP(8))
        Call inout(8, XP(9))
        Call inout(8, XP(10))
        Call inout(8, XP(11))
        Call inout(8, XP(12))
        Call inout(8, XP(13))
        Call inout(8, XP(14))
        Call inout(8, XP(15))
        Call inout(8, XP(16))
        Call inout(8, XP(3) + XP(4) + XP(5) + XP(6) + XP(7) + XP(8) + XP(9) + XP(14) + XP(16))
        Call inout(8, XP(10) + XP(11) + XP(12) + XP(13))
        Call inout(8, XP(3) + XP(4) + XP(5) + XP(6) + XP(7) + XP(8) + XP(9) + XP(14) + XP(16) + XP(10) + XP(11) + XP(12) + XP(13))
        Call inout(8, UPPAG)
        Call inout(8, UPPBG)
        Call inout(8, UPPFF)
        Call inout(8, DecayP)
        Call inout(8, DeficitP)
        Call newline(8)
    End If
      
    j = Int(T)
    j = j - 1
'C      IF(YESA.EQ.1.0)WRITE(8,12)((A(I,L),L=1,14),I=1,14),
'C     1NAP,Y(2),Y(7),Y(10),Y(11),Y(12),Y(13)
      
End Sub


'C******************************************************************************

Sub PHOTO(INPNAP, NAP, T, LA1Y, LA2Y, LA3Y, RNAPA, RNAPB, RNAPC, SPHA, HTSP, TVOLSP, _
    PNA, PNB, AY1A, AY1B, AY1C, SX3, ADIST, X, TNAPAG, X3MAX, S, B31, B32, GCL, HT, _
    GCLMAX, GCFRE, TLAM, DMPNY, PMAI, DENSREG, DRA, DRB, DRC, CNRAT, DRHA, ZNAP, _
    BASP, F128SP, F82SP, SPHSP, RADIA, DENWH, DENINC)

'C
'C PURPOSE

'C    THE FORCING FUNCTION, Z IS EQUIVALENT TO NET ANNUAL PRODUCTION, NAP
'C    OPTIONS ARE TO TO EITHER READ MEASURED STEM ANNUAL PRODUCTION, ZNAP
'C    AND CONVERT TO Z;
'C    OR ALTERNATIVELY, SIMULATE Z FROM MODELLED LEAF AREA INDEX.

    Dim PMIMM(12) As Double, RAD(12) As Double, XIMM(12) As Double, AGESP(100) As Double, PNM(12) As Double, _
        ADJPM(12) As Double, TWT(100) As Double, TWTG(100) As Double, VOLGSP(100) As Double, j As Long, _
        NYRSD As Double, i As Long, TT As Long, DENSAGE As Double, ITT As Long, ILL As Long, _
        A As Double, B As Double, C As Double, D As Double, TP As Double, PSTMW As Double, PROOT As Double

    j = Int(T)
    If INPNAP <> 1 Then Exit Sub
'C Note convention - variable names followed by SP are from an external model (eg 300 INDEX)
    Call inout(5, NYRSD)
    Call newline(5)
    For i = 1 To NYRSD
        Call inout(5, AGESP(i))
        Call inout(5, SPHSP(i))
        Call inout(5, HTSP(i))
        Call inout(5, TVOLSP(i))
        Call inout(5, VOLGSP(i))
        Call inout(5, BASP(i))
        Call inout(5, DENWH(i))
        Call inout(5, DENINC(i))
        Call newline(5)
    Next i
    
'C  1  FORMAT(I2/,  (8F10.4))
    Call inout(10, " CARDTYPES N AND O: NYRS OF DATA =")
    Call inout(10, NYRSD)
    Call newline(10)
    Call inout(10, "AGE")
    Call inout(10, "SPH")
    Call inout(10, "HT")
    Call inout(10, "LIVEVOL")
    Call inout(10, "GROSSVOL")
    Call inout(10, "BA")
    Call inout(10, "STEMDENS")
    Call inout(10, "RINGDENS")
    Call newline(10)
    For i = 1 To NYRSD
        Call inout(10, AGESP(i))
        Call inout(10, SPHSP(i))
        Call inout(10, HTSP(i))
        Call inout(10, TVOLSP(i))
        Call inout(10, VOLGSP(i))
        Call inout(10, BASP(i))
        Call inout(10, DENWH(i))
        Call inout(10, DENINC(i))
        Call newline(10)
    Next i

'C USING ANNUAL WHOLE STEM DENSITY VALUES, CONVERT TOTAL STEM VOLUME UNDER BARK TO
'C STEM WEIGHT UNDER BARK, AND CALCULATE STEM ANNUAL DM INCREMENT BY DIFFERENCE
'C Note that this calculation is not used if the incremental wood density is known.
    For i = 1 To NYRSD
'C COMPUTE WHOLE STEM DENSITY ESTIMATE (T/M**3) BY REGION: 1=LOW; 2=MEDIUM;
'C 3=HIGH; 5=SITE FERTILITY BASED DENSITY
        TT = AGESP(i)
        If DENSREG <= 3 Then DENSAGE = (DRC + DRA * (1 - Exp(DRB * TT))) / 1000
        ITT = Int(TT)
        If DENSREG = 5 Then DENSAGE = DENWH(ITT)
        If DENSREG = 6 Then GoTo 3003
'C GROSS AND NET VOLUMES ARE ASSUMED TO HAVE SAME WHOLE STEM DENSITY
'redundant code        TWTG(i) = DENSAGE * VOLGSP(i)
'redundant code        TWT(i) = DENSAGE * TVOLSP(i)
    Next i

'C CALCULATE NET (F82SP)AND GROSS (ZNAP) STEM WOOD INCREMENT
'redundant code    ILL = 0
'redundant code    For i = 1 To NYRSD - 1
'redundant code        If AGESP(i + 1) <> AGESP(i) Then ILL = ILL + 1
'redundant code        If AGESP(i + 1) <> AGESP(i) Then F128SP(ILL) = TWTG(i + 1) - TWT(i + 1)
'redundant code        If (AGESP(i + 1) <> AGESP(i)) Then F82SP(ILL) = TWTG(i + 1) - TWT(i)
'redundant code        If AGESP(i + 1) <> AGESP(i) Then ZNAP(ILL) = TWTG(i + 1) - TWT(i)
'redundant code    Next i
'redundant code   GoTo 3006

'C CALCULATE GROSS STEM DM INCREMENT FROM RING DENSITY (T/M**3)- SITE FERTILITY BASED
'C This calculation method is preferred.
3003: ILL = 0
    For i = 1 To NYRSD - 1
        If AGESP(i + 1) <> AGESP(i) Then ILL = ILL + 1
        If AGESP(i + 1) <> AGESP(i) Then F128SP(ILL) = (VOLGSP(i + 1) - TVOLSP(i + 1)) * DENWH(i + 1)
        If AGESP(i + 1) <> AGESP(i) Then F82SP(ILL) = (VOLGSP(i + 1) - TVOLSP(i)) * DENINC(i + 1)
        If AGESP(i + 1) <> AGESP(i) Then ZNAP(ILL) = (VOLGSP(i + 1) - TVOLSP(i)) * DENINC(i + 1)
    Next i
3006:

'C Convert stemwood dm production (ZNAP) to above ground production, and NAP.
'C (NOTE: CHRONOLOGICAL AGE ASSUMED).
'C Data from Webber and Madgwick (1983) NZJFSci vol 13(2); Madgwick (1985)
'C NZJFSci vol 15(3); Beets and Pollock (1987),
'C NZJFSci vol 17(2/3).
'C Note, estimated value applies to age at end of increment period. eg partitioning
'C from age 0-1 years given by equation at age 1.
      
'C PARTITIONING TO STEM WOOD (UNDER BARK).(logistic equation).

    A = 65
    B = -2.9473
    C = 8.3314
    D = 28.5213
    TP = 0
      
    For i = 1 To ILL
        TP = TP + 1
        PSTMW = ((A - D) / (1 + (TP / C) ^ B) + D) / 100
        TNAPAG = ZNAP(i) / PSTMW
'C Assume that 30% of total annual production is partitioned to roots.

        PROOT = 0.3
        ZNAP(i) = TNAPAG / (1 - PROOT)
    Next i
    
End Sub


'C******************************************************************************

Sub DMFMAT(X, NAP, F, Y, T, LA1Y, PNA, PNB, pBA, PBB, PBC, CONS, GCL, X3MAX, ID, GCFRE, _
    BA, SPHA, X6MAX, TMA, TMB, VMA, VMB, VMC, VMD, INDEAD, HINC, MAXL1Y, _
    TNAPAG, S, SX3, B31, B32, HT, RADIUS, INPNAP, PA, _
    NLIMITED, PARTF, PARTB, PARTS, PARTRC, ZNAP, F128SP, F82SP, MATEMP)
    
    Dim j As Long, PT As Double, PROOT As Double, BG As Double, CVD As Double, dbh As Double, _
        DBHMN1 As Double, ROOTS As Double, PREVRG2 As Double, PREVRTS As Double, ROOTSG2 As Double, _
        A As Double, B As Double, C As Double, D As Double, PFOL As Double, PSTMW As Double, PSTMWB As Double, _
        STMMORT As Double, PREP As Double, PBRRP As Double, X8MAX As Double, X8NEW As Double, X16NEW As Double, _
        RWB As Double, X816NEW As Double, X816MRT As Double, RCW As Double, TMBIV As Double, RCS As Double, _
        X3RET As Double, X4RET As Double, X3LOSS As Double, X4LOSS As Double, RAT As Double, TMORT As Double, _
        DECAY_CONSTANT_STEM As Double, DECAY_CONSTANT_BRANCH As Double, DECAY_CONSTANT_COARSEROOT As Double, _
        DECAY_CONSTANT_NEEDLES As Double

'C PURPOSE:
'C
'C   NET PRIMARY PRODUCTION, Z IS PARTITIONED INTO FLOWS(F), TO BE
'C   ADDED TO LIVING COMPONENTS. ALSO, MORTALITY FLOWS ARE SIMULATED. LOSSES
'C   (Y) ARE ASSUMED TO EXIT FROM THE STAND, WHEN IN PRACTICE SOME LOSSES WILL
'C   BE ENTERING THE SOIL COMPARTMENT - (WHICH IS NOT INSTALLED IN THIS VERSION)

'C  THE TO/FROM CONVENTION IS USED, EG. F(10,4) IS FLOW TO NEEDLE LITTER FROM
'C  1-2 YEAR OLD NEEDLE.
      
'C Set time, TP to end of increment period for partitioning functions.

    j = Int(T)

'C  SET NAP TO APPROPRIATE ZNAP (TIME)
    If INPNAP = 1 Then NAP = ZNAP(j + 1)
'C   CURRENT PHYSIOLOGICAL (PT) AGE IS CHRONOLOGICAL AGE (T) PLUS INITIAL
'C   PHYSIOLOGICAL AGE (PA)
    PT = T + 1 + PA
    F(2, 1) = NAP * 2
'C    (IF OPTION INPNAP EQUALS 3 THEN IN PHOTO, RESID FOR RESP,ROOT GROWTH
'C     AND ERROR)
'Redundant code    If NLIMITED = 1 Then GoTo 20
'C   Root partitioning estimated. Root was added to TNAPAG in PHOTO to give NAP,
'C   and is subtracted from NAP here to give TNAPAG again, for use by above
'C   ground partitioning functions.
      
'C  PARTITION NAP TO ROOTS, WHEN TREES ARE TOO SMALL FOR DBH METHOD
'C  USE 30% ASSUMPTION FOR YOUNG TREES
      
    PROOT = 0.3
    BG = NAP * PROOT
'c  this version assumes roots are 30% - but code below allows for Jackson et
'c      IF(BA.LT.10.0)F(9,2) = 0.0
'c      IF(BA.LT.10.0) GO TO 211

'C  PARTITIONING TO ROOTS>5MM WHEN NAP SIMULATED FROM LEAF AREA FOR LARGE TREES
'C  THE ARITHMETIC MEAN DBH REQUIRES THE CV OF DBH. ASSUME THIS TO BE CONSTANT
'C  BUT SEE BEETS AND KIMBERLEY 1993: GENOTYPE X STOCKING INTERACTIONS IN PINUS
'C  RADIATA: PRODUCTIVITY AND YIELD IMPLICATIONS.STUDIA FORESTALIA SUECICA 191.
'C  USE BA AT START OF INCREMENT PERIOD, AS END PERIOD DBH NOT KNOWN YET.
      
    CVD = 0.2
    DBHMN1 = dbh
    dbh = Sqr(4 * BA / (3.14159 * SPHA * (1 + CVD ^ 2))) * 100
    ROOTS = (0.00000587 * dbh ^ 2.938) * SPHA
    PREVRTS = (0.00000587 * DBHMN1 ^ 2.938) * SPHA

    ROOTSG2 = (0.00000597 * dbh ^ 2.8068) * SPHA
    PREVRG2 = (0.00000597 * DBHMN1 ^ 2.8068) * SPHA

'C Deactivate

'C BG = ROOTS - PREVRTS
'C F(9, 2) = ROOTSG2 - PREVRG2
      
'c      write(10,212)t, nap, bg
'c  212 format(1h , 6f7.2)
'C  211   CONTINUE

'C   CALCULATE ABOVE GROUND PRODUCTION (BY DIFFERENCE).
  
    TNAPAG = NAP - BG

'C    Above ground partitioning coefficients based on logistic equation.
'C    Data from Webber and Madgwick (1983) NZJFSci vol 13(2); Madgwick (1985)
'C    NZJFSci vol 15(3); Beets and Pollock (1987),
'C    NZJFSci vol 17(2/3).
'C    Note, estimated value applies to age at end of increment period.
'C    eg partitioning from age 0-1 years given by equation at age 1.

'C  PARTITION TNAPAG TO NEW NEEDLES
      
    A = 39.5801
    B = 3.3371
    C = 7.867
    D = 12
    PFOL = ((A - D) / (1 + (PT / C) ^ B) + D) / 100
    F(3, 2) = TNAPAG * PFOL

'C  PARTITION TNAPAG TO STEM (WOOD excl. BARK)
      
    A = 65
    B = -2.9473
    C = 8.3314
    D = 28.5213
    PSTMW = ((A - D) / (1 + (PT / C) ^ B) + D) / 100
    F(8, 2) = TNAPAG * PSTMW
      
'C  PARTITION TNAPAG TO STEM (WOOD Plus BARK)
      
    A = 72
    B = -2.9383
    C = 8.7097
    D = 33.8497
    PSTMWB = ((A - D) / (1 + (PT / C) ^ B) + D) / 100
      
'C PARTITIONING TO STEM BARK

    F(16, 2) = TNAPAG * PSTMWB - F(8, 2)

'C  SCALER APPLIED TO ALLOW FOR STEM BARK MORTALITY FROM STANDPAK VOLUMES
    STMMORT = (TNAPAG * PSTMWB) / F82SP(j + 1)

'C  PARTITION TNAPAG TO CONE PRODUCTION (CURRENTLY MERGED WITH BRCH)

    A = 2.597
    B = -3.8161
    C = 12.5417
    D = 0
    PREP = ((A - D) / (1 + (PT / C) ^ B) + D) / 100

'C  PARTITION TNAPAG TO BRANCHES AND REPRODUCTION, COMBINED (BY DIFFERENCE).

    PBRRP = (1 - PFOL - PSTMWB)
    F(6, 2) = TNAPAG * PBRRP
      
'C  PARTITION BG TO COARSE AND FINE ROOT
      
    F(9, 2) = (F(8, 2) + F(6, 2)) * 0.2
    F(14, 2) = BG - F(9, 2)

'C DONT APPLY THIS CONSTRAINT TO FOLIAGE PRODUCTION NOW
'C     IF(F(3,2).GT.X3MAX)F(3,2)=X3MAX


'C    PARTITIONING TO FINE ROOT(<2MM) PROPORTIONAL TO FASCICLE PRODUCTION
'C    (JACKSON AND CHITTENDEN N.Z J. FOR SCI 1981) IN YOUNG STANDS, AND
'C    COARSE ROOTS BY DIFFERENCE IN OLDER STANDS
 
'C Deactivate

'c      IF(F(9,2).EQ.0.0) F(14,2) = F(3,2)*0.5
'c      IF(INPNAP.NE.3.AND.F(9,2).GT.0.0) F(14,2) = BG - F(9,2)
'c      IF(F(9,2).EQ.0.0) F(9,2) = BG - F(14,2)

      
    GoTo 21

'C    WITH NLIMITED OPTION, APPLY PARTITIONING COEF TO COMPUTE COMPONENT GROWTH
'redundant code20: F(3, 2) = PARTF * NAP
'redundant code    F(6, 2) = PARTB * NAP
'redundant code    F(8, 2) = PARTS * NAP
'redundant code    F(9, 2) = PARTRC * NAP
'redundant code    F(14, 2) = NAP - (F(3, 2) + F(6, 2) + F(8, 2) + F(9, 2))
21:

'C    COMPUTE MAXIMUM STEM VOLUME ACCORDING TO THE SIZE-SPACING RELATION
    If INDEAD = 0 Then GoTo 3
'redundant code    X8MAX = (VMA * SPHA ^ VMB) * SPHA
'redundant code    X8NEW = x(8) + F(8, 2)
'redundant code    X16NEW = x(16) + F(16, 2)
'redundant code    RWB = X8NEW / (X8NEW + X16NEW)
'redundant code    X816NEW = x(8) + x(16) + F(8, 2) + F(16, 2)
'redundant code    If X816NEW <= X8MAX Then GoTo 3
'redundant code    X816MRT = X816NEW - X8MAX
'redundant code    F(12, 8) = X816MRT * RWB
'redundant code    F(12, 16) = X816MRT * (1 - RWB)
'redundant code    RCW = F(12, 8) / X8MAX
'redundant code    TMBIV = 1 / TMB
'redundant code    RCS = RCW ^ TMBIV
'redundant code    SPHA = (1 - RCS) * SPHA
'redundant code    GoTo 4

'C    TRANSFER STANDING COMPONENTS TO LITTER COMPARTMENTS
3:  F(12, 8) = 0
    F(12, 16) = 0
'C    IF GROSS/NET VOLUMES INPUT FROM SP, USE THIS FOR MORTALITY
    If INPNAP = 1 Then F(12, 8) = F128SP(j + 1)
    If INPNAP = 1 Then F(12, 16) = (F128SP(j + 1) * STMMORT) - F128SP(j + 1)
4:

'C Consumption of 1yr fascicles WHAT ABOUT NUTRIENTS???
    Y(3) = CONS * X(3)
'The following statement appears to reduce 1-year foliage prematurely - this is correctly carried out in DMAMAT
'    X(3) = (1 - CONS) * X(3)
'C    FASCICLE RETENTION ASSUMED CONSTANT FOR AGE 1 AND AGE 2 FASCICLES.
'C    THE RETENTION COEFFICIENT SX3 WILL VARY DEPENDING ON SITE,
'C    CLIMATE AND BIOLOGICAL FACTORS.
    X3RET = SX3
    X4RET = SX3 / 2
    X3LOSS = 1 - X3RET
    X4LOSS = 1 - X4RET
'C FRACTIONAL RETENTION AFTER CONSUMPTION BY HERBIVORS
    F(10, 3) = X(3) * X3LOSS
    F(10, 4) = X(4) * X4LOSS
    F(10, 5) = X(5)
    F(11, 7) = 0
'C    RECRUITMENT TO OLDER NEEDLE AGE CLASSES

'C    ONE TO TWO YEAR NEEDLES
    F(4, 3) = X(3) - F(10, 3)
'C    TWO TO THREE YEAR NEEDLES
    F(5, 4) = X(4) - F(10, 4)
'C    RECRUITMENT OF LIVE BRANCHES TO ATTACHED DEAD COMPARTMENT
'C      IF(F(3,2).LT.X3MAX)GO TO 10
'C    BELOW ONE YEAR NEEDLE MASS MAXIMUM NO BRANCH MORTALITY IS ASSUMED
'C    AT OR ABOVE BRANCH MASS MAXIMUM BRANCH MORTALITY EQUAL TO BRANCH
'C    ANNUAL PRODUCTION ASSUMED.
    If X(6) >= X6MAX Then GoTo 9
'C    ALTERNATIVELY, BRANCH MORTALITY IS DEPENDENT ON THE DIFFERENCE
'C    BETWEEN BRANCH MAXIMUM AND SIMULATED  WEIGHT.
    RAT = X(6) / X6MAX
    F(7, 6) = F(6, 2) * RAT ^ 4
    GoTo 11
9:  F(7, 6) = F(6, 2)
    GoTo 11
'redundant code10: F(7, 6) = 0
'C    COARSE ROOT MORTALITY ASSUMED 10% OF STEM MORTALITY
11: F(13, 9) = F(12, 8) * 0.1
    F(15, 14) = X(14) * 1.5
'C    DECOMPOSITION OF LITTER COMPARTMENTS: APPLY TURNOVER RATES
'C    (based on Will et al., NZ J For Sc 13:266-304).
        DECAY_CONSTANT_STEM = 0.0376 * Exp(0.093 * MATEMP)
        DECAY_CONSTANT_BRANCH = 0.0429 * Exp(0.093 * MATEMP)
        DECAY_CONSTANT_COARSEROOT = 0.0684 * Exp(0.093 * MATEMP)
        DECAY_CONSTANT_NEEDLES = 0.081 * Exp(0.093 * MATEMP)
        PROPORTION_STEM_LOSS = 1 - Exp(-DECAY_CONSTANT_STEM)
        PROPORTION_BRANCH_LOSS = 1 - Exp(-DECAY_CONSTANT_BRANCH)
        PROPORTION_ATTACHEDBRANCH_LOSS = PROPORTION_BRANCH_LOSS * 0.0607 / 0.0936
        PROPORTION_COARSEROOT_LOSS = 1 - Exp(-DECAY_CONSTANT_COARSEROOT)
        PROPORTION_NEEDLE_LOSS = 1 - Exp(-DECAY_CONSTANT_NEEDLES)
            
    Y(10) = X(10) * PROPORTION_NEEDLE_LOSS
    Y(11) = X(11) * PROPORTION_BRANCH_LOSS
'C      ! guess 0.18 for x(11)
'C    STEM DECOMPOSITION BASED ON PURUKI-RUA LOG DECAY STUDY (Jefford 1989)
    Y(12) = X(12) * PROPORTION_STEM_LOSS
'C  COARSE ROOT
    Y(13) = X(13) * PROPORTION_COARSEROOT_LOSS
'C      ! guess  0.12 for X(13)
'C BRANCH
    Y(7) = X(7) * PROPORTION_ATTACHEDBRANCH_LOSS
'C    FINE ROOT DECOMPOSITION (based on Santantonio and Grace Can. J For Res 17:900-908
    Y(15) = X(15) * 0.52
'C    ADJUST COMPARTMENT 1 AND 2 OUTFLOWS SO THAT THEY CONFORM TO COMPART
'C    DEFINITION.
    TMORT = F(7, 6) + F(12, 8) + F(13, 9) + F(15, 14) + Y(3) + F(10, 3) + F(10, 4) + F(10, 5)
    Y(1) = X(1)
    Y(2) = X(2) + TMORT

End Sub



Sub DMAMAT(A, F, X, Y)
'C
'C PURPOSE:
'C
'C  CALCULATES RATE COEFFICIENTS BY DIVIDING FLOWS BY DONOR COMPARTMENT STATE
'C  COMPARTMENT 1
    If X(1) <> 0 Then GoTo 1
    A(1, 1) = 0
    A(2, 1) = 0
    GoTo 102
1:  A(1, 1) = -1 * (F(2, 1) + Y(1)) / X(1)
'C  COMPARTMENT 2
    A(2, 1) = F(2, 1) / X(1)
102:    If X(2) <> 0 Then GoTo 2
    A(2, 2) = 0
    A(3, 2) = 0
    GoTo 103
2:  A(2, 2) = -1 * (F(3, 2) + F(6, 2) + F(8, 2) + F(9, 2) + F(14, 2) + Y(2)) / X(2)
'C  COMPARTMENT 3
    A(3, 2) = F(3, 2) / X(2)
103:    If X(3) <> 0 Then GoTo 3
    A(3, 3) = 0
    A(4, 3) = 0
    GoTo 104
3:  A(3, 3) = -1 * (F(4, 3) + F(10, 3) + Y(3)) / X(3)
'C  COMPARTMENT 4
    A(4, 3) = F(4, 3) / X(3)
104:    If X(4) <> 0 Then GoTo 4
    A(4, 4) = 0
    A(5, 4) = 0
    GoTo 105
4:  A(4, 4) = -1 * ((F(5, 4) + F(10, 4)) / X(4))
'C  COMPARTMENT  5
    A(5, 4) = F(5, 4) / X(4)
105:    If X(5) <> 0 Then GoTo 5
    A(5, 5) = 0
    GoTo 205
5:  A(5, 5) = -1 * (F(10, 5) / X(5))
'C  COMPARTMENT  6
205:    If X(2) <> 0 Then GoTo 6
    A(6, 2) = 0
    GoTo 106
6:  A(6, 2) = F(6, 2) / X(2)
106:    If X(6) <> 0 Then GoTo 16
    A(6, 6) = 0
    A(7, 6) = 0
    GoTo 107
16: A(6, 6) = -1 * (F(7, 6) / X(6))
'C  COMPARTMENT  7
    A(7, 6) = F(7, 6) / X(6)
107:    If X(7) <> 0 Then GoTo 7
    A(7, 7) = 0
    GoTo 108
7:  A(7, 7) = -1 * (F(11, 7) + Y(7)) / X(7)
'C  COMPARTMENT  8
108:    If X(2) <> 0 Then GoTo 8
    A(8, 2) = 0
    GoTo 208
8:  A(8, 2) = F(8, 2) / X(2)
208: If X(8) <> 0 Then GoTo 18
    A(8, 8) = 0
    GoTo 109
18: A(8, 8) = -1 * (F(12, 8) / X(8))
'C  COMPARTMENT  9
109:    If X(2) <> 0 Then GoTo 9
    A(9, 2) = 0
    GoTo 209
9:  A(9, 2) = F(9, 2) / X(2)
209:    If X(9) <> 0 Then GoTo 19
    A(9, 9) = 0
    GoTo 1010
19: A(9, 9) = -1 * (F(13, 9) / X(9))
'C  COMPARTMENT 10
1010:   If X(4) <> 0 Then GoTo 10
    A(10, 4) = 0
    GoTo 2010
10: A(10, 4) = F(10, 4) / X(4)
2010: If X(5) <> 0 Then GoTo 110
    A(10, 5) = 0
    GoTo 3010
110:    A(10, 5) = F(10, 5) / X(5)
3010: If X(10) <> 0 Then GoTo 210
    A(10, 10) = 0
    GoTo 4010
210:    A(10, 10) = -1 * (Y(10) / X(10))
4010:   If X(3) <> 0 Then GoTo 310
    A(10, 3) = 0
    GoTo 1011
310:    A(10, 3) = F(10, 3) / X(3)
'C  COMPARTMENT 11
1011: If X(7) <> 0 Then GoTo 11
    A(11, 7) = 0
    GoTo 2011
11: A(11, 7) = F(11, 7) / X(7)
2011:   If X(11) <> 0 Then GoTo 111
    A(11, 11) = 0
    GoTo 1012
111:    A(11, 11) = -1 * (Y(11) / X(11))
'C  COMPARTMENT 12
1012:   If X(8) <> 0 Then GoTo 12
    A(12, 8) = 0
    GoTo 2112
12: A(12, 8) = F(12, 8) / X(8)
2112:   If X(16) <> 0 Then GoTo 1112
    A(12, 16) = 0
    GoTo 2012
1112:   A(12, 16) = F(12, 16) / X(16)
2012:   If X(12) <> 0 Then GoTo 112
    A(12, 12) = 0
    GoTo 1013
112:    A(12, 12) = -1 * (Y(12) / X(12))
'C  COMPARTMENT 13
1013:   If X(9) <> 0 Then GoTo 13
    A(13, 9) = 0
    GoTo 2013
13: A(13, 9) = F(13, 9) / X(9)
2013:   If X(13) <> 0 Then GoTo 113
    A(13, 13) = 0
    GoTo 1014
113:    A(13, 13) = -1 * (Y(13) / X(13))
'C   COMPARTMENT 14
1014:   If X(2) <> 0 Then GoTo 14
    A(14, 2) = 0
    GoTo 214
14: A(14, 2) = F(14, 2) / X(2)
214:    If X(14) <> 0 Then GoTo 114
    A(14, 14) = 0
    GoTo 1015
114:    A(14, 14) = -1 * (F(15, 14) / X(14))
'C   COMPARTMENT 15
1015:   If X(14) <> 0 Then GoTo 15
   A(15, 14) = 0
    GoTo 2015
15: A(15, 14) = F(15, 14) / X(14)
2015:   If X(15) <> 0 Then GoTo 115
    A(15, 15) = 0
    GoTo 2227
115:    A(15, 15) = -1 * (Y(15) / X(15))
'C  COMPARTMENT  16
2227: If X(2) <> 0 Then GoTo 2228
    A(16, 2) = 0
    GoTo 2229
2228:   A(16, 2) = F(16, 2) / X(2)
2229:   If X(16) <> 0 Then GoTo 2230
    A(16, 16) = 0
2230:   A(16, 16) = -1 * (F(12, 16) / X(16))
'C  ALL OTHER ELEMENTS OF A ARE 0.0

End Sub
   



'C******************************************************************************

Sub STRUCT(AVSSA, AVBSA, AVCRA, POTCL, SPHSP, DENSREG, X8GAIN)

'C
'C PURPOSE:
'C
'C    SIMULATES HEIGHT GROWTH,GREEN CROWN LENGTH,HEIGHT TO BASE OF
'C    GREEN CROWN ,STAND BASAL AREA AND STOCKING
'C    ALSO CALCULATES ONE YEAR MAXIMUM NEEDLE MASS AND AREA, AND BRANCH
'C    MAXIMUM WEIGHT.

'C Much of the following is not required when C_Change is linked with an external model eg 300 INDEX
    Dim LAMN1 As Double, XMAXL1 As Double, HOLD As Double, NOTMAX As Long, ISPT As Long, th As Double, _
        BSS As Double, X3RET As Double, X3MAX As Double, X4MAX As Double, X5MAX As Double, GCLN As Double, _
        BIN32 As Double, RAGE As Double, HTT As Double, TESTAG As Double, DBHBAR As Double, C As Double, _
        CNEG As Double, CFR1 As Double, VOLBAR As Double, SUMDI As Double, CT As Double, CMIN1 As Double, _
        HI As Double, SDBARS As Double, i As Long, HII As Double, DI As Double, PRESSA As Double, _
        BSA As Double, PREBSA As Double, PRECRA As Double, SDSQBC As Double

    LAMN1 = LA1Y
    XMAXL1 = MAXL1Y
    HOLD = HT
    If T = ADIST(1) Then NOTMAX = 0

    T = T + 1
    ISPT = Int(T)
    th = T

'C    CALCULATE STAND HEIGHT AND HEIGHT INCREMENT
    BSS = S / (1 - Exp(B31 * S * 20)) ^ B32
    HT = BSS * (1 - Exp(B31 * S * th)) ^ B32

'C  USE INPUT HEIGHT STOCKING AND BA IF RUNNING MODEL FROM VOLUME
'C Note the convention - names ending with SP are from external model (eg 300 Index)
    If HTSP(ISPT) > 0 Then HT = HTSP(ISPT + ID - 1)
    If SPHSP(ISPT) > 0 Then SPHA = SPHSP(ISPT + ID - 1)
    If BASP(ISPT) > 0 Then BA = BASP(ISPT + ID - 1)
    HINC = HT - HOLD
'C    CALCULATE STAND LEAF AREA AND LEAF AREA AT CANOPY CLOSURE
'C    WILL ASSUME THAT NEEDLE MAXIMUM WEIGHT RELATED COMPLETELY TO STAND
'C    AGE,RATHER THAN TO A COMBINATION WITH STOCKING.
    X3RET = SX3
    X3MAX = AM1 * Exp(AM2 * (T + 1))
    X4MAX = X3MAX * X3RET
    X5MAX = X4MAX * X3RET * 0.5
    LA1Y = (AY1A * X(3) ^ 2 + AY1B * X(3) + AY1C) * X(3)
    LA2Y = (AY1A * X(4) ^ 2 + AY1B * X(4) + AY1C) * X(4)
    LA3Y = (AY1A * X(5) ^ 2 + AY1B * X(5) + AY1C) * X(5)
    MAXL1Y = (AY1A * X3MAX ^ 2 + AY1B * X3MAX + AY1C) * X3MAX
'C    GREEN CROWN LENGTH INCREMENTED BY HEIGHT GROWTH
    GCLN = GCL + HINC
'C    COMPUTE MAXIMUM GREEN CROWN LENGTH FOR THE STOCKING AND STAND HEIGHT
'C    BASED ON BEEKHUIS HARMONIZED CURVES.
'C BS = 62.57
'C R = -0.21137
'C      AS = BS * EXP(R*ALOG(SPHA))
'C BM = -0.11
'C      GCLMAX = AS * (1.0 - EXP(BM * HT))

'C     OR BASED ON ANDY DUNNINGHAM
    

'C    COMPARISON OF POTENTIAL CROWN LENGTH USING ANDY DUNNINGHAM FUNCTION
    If SPHA >= 1000 Then GCLMAX = HT ^ 0.45077 * (Sqr(10000 / SPHA)) ^ 0.32111 * 2.43641
    If SPHA < 1000 Then GCLMAX = HT ^ 0.8888 * (Sqr(10000 / SPHA)) ^ 0.20861 * 0.77889
    If GCLN < GCLMAX Then GCL = GCLN
    If GCLN >= GCLMAX Then GCL = GCLMAX
    HTBGC = HT - GCL
'C    COMPUTE CANOPY AGE
    If HTBGC < 0.2 Then GoTo 650
    BIN32 = 1 / B32
    RAGE = Log(1 - ((HTBGC / BSS) ^ BIN32)) / (B31 * S)
650:    If HTBGC < 0.2 Then RAGE = 0
    TBRCH = T - RAGE
'C    IF DISTURBANCE INTENDED AT THIS AGE THEN SIMULATE GCLMAX
    If T < ADIST(ID) Then GoTo 400
    If GCL = GCLMAX Then GoTo 400
'C    PROJECT FORWARD IN TIME TILL GCL ATTAINS A MAXIMUM
'C    CAN ASSUME ZERO MORTALITY OVER THIS TIME INTERVAL AS ARE BELOW FULL
'C    SITE OCCUPANCY
    HTT = th
    TESTAG = HT
300:    HTT = HTT + 1
    HOLD = TESTAG
    TESTAG = BSS * (1 - Exp(B31 * S * HTT)) ^ B32
    HOLD = TESTAG - HOLD
    GCLN = GCLN + HOLD
'************ bug    GCLMAX = AS*(1.0 - EXP(BM * TESTAG))
'    If GCLN < GCLMAX Then GoTo 300
'The above line was modified as below by M Kimberley in August 2009 to remove endless loop problem in one LUCAS plot
    If GCLN < GCLMAX And HTT < 200 Then GoTo 300
400:
'C    SIMULATE STAND BASAL AREA
'C    UTILIZES RELATIONSHIP BETWEEN STEM WEIGHT/HA AND THE PRODUCT OF
'C    BA(M**2/HA) AND HEIGHT(M)
    If BASP(ISPT) = 0 Then BA = ((X(8) + X(16)) - SFINT * SPHA) / (HT * SFSPE)
    If BA < 0 Then BA = 0
'C    THE EMPIRICAL FORM COEFFICIENTS, WHICH UNDERESTIMATE BA FOR
'C    STANDS LESS THAN FOUR YEARS OF AGE, ARE FOR THE PRESENT ASSUMED TO
'C    BE CONSTANTS
'C    TEST IF STEM NATURAL MORTALITY LIKELY
'C    STEM MORTALITY SIMULATED IN DMFMAT;A CASE OF FUNCTION INFLUENCING
'C    STRUCTURE.
'C
'C    CALCULATE MEAN DIMENSIONS OVER BARK OF STEM
'C    STEM MUST EXCEED 1.4M HEIGHT
    If HT <= 1.4 Then SDBAR = 0
    If HT <= 1.4 Then GoTo 703
'C
'C    DIAMETER OVER BARK OF STEM WITH MEAN BA.
    DBHBAR = Sqr(BA * 4 / (SPHA * 3.1416))
'C
'C    VOLUME OVER BARK OF STEM WITH MEAN BA.
    C = 3.06
    CNEG = -1 * C
    CFR1 = 1 - C
    VOLBAR = (DBHBAR ^ 2) * ((HT - 1.4) ^ CFR1) * (HT ^ C) / C
    SUMDI = 0
    CT = 0
'C    MEAN TAPER DIAMETER OVER BARK OF STEM WITH MEAN BA.
    CMIN1 = C - 1
    HI = HT / 20
    SDBARS = 0
    For i = 1 To 20
        HII = HI * CT
        DI = (HT - HII) ^ CMIN1
        CT = CT + 1
        SDBAR = Sqr(VOLBAR * C * DI * (HT ^ CNEG))
        SDBARS = SDBARS + SDBAR
    Next i
'C    STEM DIAMETER OVER BARK IN CMS FOR BENECKES FORMULA
    SDBAR = (SDBARS / 21) * 100
'C    STEM SURFACE AREA OVER BARK IN M**2
'C     (STORE PREVIOUS YEARS SURFACE AREA VALUES PRIOR TO UPDATING
    PRESSA = SSA
    PREBSA = BSA
    PRECRA = CRA
    SSABAR = 2.708 * Sqr(VOLBAR * HT)
    SSA = SSABAR * SPHA / 10000
'C    BASAL AREA OVER BARK AT BASE OF GREEN CROWN
    SDSQBC = VOLBAR * C * (HT - HTBGC) ^ CMIN1 * (HT ^ CNEG)
    BABGC = SDSQBC * SPHA * 3.1416 / 4
    TVOL = VOLBAR * SPHA
703:

'C    STEM VOLUME UNDER BARK FROM STEM TOTAL DRY WEIGHT AND WOOD DENSITY
'C    CALCULATION ASSUMES WOOD DENSITY IS BASED ON WHOLE STEM DENSITY
 
    If DENSREG <> 6 Then vol = X(8) / dens

'C    OR FROM RING DENSITY
    If INPNAP = 2 And DENSREG = 6 Then vol = vol + (X8GAIN) / DENINC(ISPT)
'C  NOTE THAT WHEN INPNAP= 1 DENSITY IS NOT ORDERED BY (UNLIKE WHEN=2)
    If INPNAP = 1 And DENSREG = 6 Then vol = vol + (X8GAIN) / DENINC(ISPT + ID - 1)

    If TVOLSP(ISPT) > 0 Then vol = TVOLSP(ISPT + ID - 1)

'C WHEN DENSREG 6 (RING DENSITY) USED, THEN DERIVE WHOLE STEM DENSITY FROM:
'C STEM VOLUME AND MASS (APPLIES TO BOTH INPNAP 1 AND 2):
    If vol > 0 And DENSREG = 6 Then dens = X(8) / vol


'C    CALCULATE STAND MAXIMUM BRANCH WEIGHT AT CANOPY CLOSURE AS
'C    DETERMINED BY STOCKING.
    X6MAX = BM1 * (1# / SPHA) ^ BM2
'C    TOTAL LEAF AREA
    TLA = LA1Y + LA2Y + LA3Y
'C    BRANCH SA (ASSUMED 10% 0F TLA)
    BSA = TLA * 0.1
'C    COARSE ROOT SA(RUNNER ROOTS EQUAL BSA, AND ROOT STOCK PROPORTIONAL
'C    TO STEM SURFACE AREA
    CRA = BSA + (SSA * 0.22)
'C    COMPUTE AVERAGE SURFACE AREA FOR YEAR FOR RESPIR
    AVSSA = (PRESSA + SSA) / 2
    AVBSA = (PREBSA + BSA) / 2
    AVCRA = (PRECRA + CRA) / 2

End Sub



'C******************************************************************************

Sub NUTRIENT(F, X, XMN1, XN, XP, XNMN1, XPMN1, UPNAG, UPPAG, UPNBG, UPPBG, T, TBRCH, Y)
'C
'C PURPOSE:
'C
'C    SIMULATES NUTRIENT CONTENT AND UPTAKE OF STAND COMPARTMENTS FROM COMPONENT
'C    WEIGHTS AND COMPONENT NUTRIENT CONCENTRATIONS (PERCENT OD WEIGHT).
'      DIMENSION x(16), xn(16), xp(16), XMN1(16), XNMN1(16), F(16, 16)
'C
'C    FASCICLE N CONTENT(KG/HA) FROM COMPONENT WT(T/HA) AND CONCENTRATION(%)
'C
    Dim DBN As Double, DBP As Double, DSN As Double, DSP As Double, UPWDY As Double, UPWDYP As Double, UPBARK As Double, UPBARKP As Double, _
        FLC As Double, FLCP As Double, REMXN3 As Double, REMXP3 As Double, REMXN4 As Double, REMXP4 As Double, REMXN5 As Double, REMXP5 As Double, _
        UPFAS As Double, UPFASP As Double, DRNC As Double, DRPC As Double, DRNF As Double, DRPF As Double, _
        decayN10 As Double, decayN11 As Double, decayN12 As Double, decayN13 As Double, decayN15 As Double, decayN7 As Double, _
        decayP10 As Double, decayP11 As Double, decayP12 As Double, decayP13 As Double, decayP15 As Double, decayP7 As Double
'      FVAR1 = SQRT(2)
'      FVAR2 = SQRT(3)
'      FNC1 = FNA
'      FNC2 = FNB
'      FNC3 = FNB
'C FNC2 = FNA * Exp(FNB)
'C FNC3 = FNA * Exp(FNB * FVAR1)
'      FNCPL1 = FNA * Exp(FNB * FVAR2)
    XN(3) = X(3) * Nconc(3, T) * 10
    XN(4) = X(4) * Nconc(4, T) * 10
    XN(5) = X(5) * Nconc(5, T) * 10
    XP(3) = X(3) * Pconc(3, T) * 10
    XP(4) = X(4) * Pconc(4, T) * 10
    XP(5) = X(5) * Pconc(5, T) * 10
'C
'C    BRANCH N CONTENT
'C
'      BVAR = SQRT(TBRCH)
'      BNC = BNA
'C BNC = BNA * Exp(BNB * BVAR)
    XN(6) = X(6) * Nconc(6, T) * 10
    XP(6) = X(6) * Pconc(6, T) * 10
'C
'C    STEM N CONTENT
'C
'      SVAR = SQRT(t)
'      SNC = SNA
'C SNC = SNA * Exp(SNB * SVAR)
    XN(8) = X(8) * Nconc(8, T) * 10
    XP(8) = X(8) * Pconc(8, T) * 10
'C
'C    STEMBARK N CONTENT
'C
    XN(16) = X(16) * Nconc(16, T) * 10
    XP(16) = X(16) * Pconc(16, T) * 10
'C
'C    DEAD BRANCH N CONTENT
'C
    XN(7) = X(7) * Nconc(7, T) * 10
    XP(7) = X(7) * Pconc(7, T) * 10
'C    ROOT N CONTENT(FORMULATION ASSUMES COARSE ROOT CONCENTRATION SAME AS
'C    BRANCHES AND FINE ROOTS SAME AS 2 YEAR FASCICLES)
'C
    XN(9) = X(9) * Nconc(9, T) * 10
    XN(14) = X(14) * Nconc(14, T) * 10
    XP(9) = X(9) * Pconc(9, T) * 10
    XP(14) = X(14) * Pconc(14, T) * 10
'C
'C    N UPTAKE BY ABOVE GROUND WOODY COMPONENTS, WITH ADJUSTMENT FOR NATURAL
'C    BRCH AND STEM MORTALITY(FORMULATION ASSUMES CURRENT MORTALITY HAS SAME
'C    N CONC. AS LIVING COMPONENT)REMOBILIZATION IS IMPLIED IN COEFFICIENTS
'C SNA, SNB, BNA, BNB
'C
'C
'C    N UPTAKE BY FASCICLES, WITH ADJUSTMENT FOR REMOBILIZATION FROM OLDER
'C    FASCICLES(FORMULATION ASSUMES ABSCISED FASCICLES DROP TO THE NUTRIENT
'C    CONCENTRATION SPECIFIED FOR PURPOSE OF ESTIMATING REMOBILIZATION)
'C
'      FLC = SNB
    FLC = Nconc(10, T)  'N content of litter fall
    REMXN3 = XNMN1(3) - (F(4, 3) * Nconc(4, T) * 10) - (F(10, 3) * FLC * 10)
    REMXN4 = XNMN1(4) - (F(5, 4) * Nconc(5, T) * 10) - (F(10, 4) * FLC * 10)
    REMXN5 = XNMN1(5) - (F(10, 5) * FLC * 10)
    UPFAS = XN(3) - (REMXN3 + REMXN4 + REMXN5)
    
    FLCP = Pconc(10, T)  'P content of litter fall
    REMXP3 = XPMN1(3) - (F(4, 3) * Pconc(4, T) * 10) - (F(10, 3) * FLCP * 10)
    REMXP4 = XPMN1(4) - (F(5, 4) * Pconc(5, T) * 10) - (F(10, 4) * FLCP * 10)
    REMXP5 = XPMN1(5) - (F(10, 5) * FLCP * 10)
    UPFASP = XP(3) - (REMXP3 + REMXP4 + REMXP5)
'C
'C    N INTO DEAD BRANCH AND LITTER COMPARTMENTS
'C
    decayN10 = 0
    decayN11 = 0
    decayN12 = 0
    decayN13 = 0
    decayN15 = 0
    decayN7 = 0
    If X(10) <> 0 Then decayN10 = XN(10) * PROPORTION_NEEDLE_LOSS * 0.5 'Release rate of N during decay is half that of C
    If X(11) <> 0 Then decayN11 = XN(11) * PROPORTION_BRANCH_LOSS * 0.5
    If X(12) <> 0 Then decayN12 = XN(12) * PROPORTION_STEM_LOSS * 0.5
    If X(13) <> 0 Then decayN13 = XN(13) * PROPORTION_COARSEROOT_LOSS * 0.5
'    If X(7) <> 0 Then decayN7 = XN(7) * PROPORTION_ATTACHEDBRANCH_LOSS * 0.5
'    DecayN = decayN10 + decayN11 + decayN12 + decayN13 + decayN7
    DecayN = decayN10 + decayN11 + decayN12 + decayN13
     
    decayP10 = 0
    decayP11 = 0
    decayP12 = 0
    decayP13 = 0
    decayP15 = 0
    decayP7 = 0
    If X(10) <> 0 Then decayP10 = XP(10) * PROPORTION_NEEDLE_LOSS * 0.76 'Release rate of P during decay is 76% that of C
    If X(11) <> 0 Then decayP11 = XP(11) * PROPORTION_BRANCH_LOSS * 0.76
    If X(12) <> 0 Then decayP12 = XP(12) * PROPORTION_STEM_LOSS * 0.76
    If X(13) <> 0 Then decayP13 = XP(13) * PROPORTION_COARSEROOT_LOSS * 0.76
'    If X(7) <> 0 Then decayP7 = XP(7) * PROPORTION_ATTACHEDBRANCH_LOSS * 0.76
'    DecayP = decayP10 + decayP11 + decayP12 + decayP13 + decayP7
    DecayP = decayP10 + decayP11 + decayP12 + decayP13
     
'    XN(7) = XN(7) + DBN - decayN7
    XN(10) = XN(10) + (F(10, 3) * FLC + F(10, 4) * FLC + F(10, 5) * FLC) * 10 - decayN10
    XN(11) = XN(11) - decayN11
    XN(12) = XN(12) + DSN - decayN12
    DRNC = XN(9) * (F(13, 9) / X(9))
    XN(13) = XN(13) + DRNC - decayN13
    DRNF = XN(14) * (F(15, 14) / X(14))

'    XP(7) = XP(7) + DBP - decayP7
    XP(10) = XP(10) + (F(10, 3) * FLCP + F(10, 4) * FLCP + F(10, 5) * FLCP) * 10 - decayP10
    XP(11) = XP(11) - decayP11
    XP(12) = XP(12) + DSP - decayP12
    DRPC = XP(9) * (F(13, 9) / X(9))
    XP(13) = XP(13) + DRPC - decayP13
    DRPF = XP(14) * (F(15, 14) / X(14))
'C    N UPTAKE FOR ROOTS WITH ADJUSTMENTS FOR NATURAL MORTALITY(FORMULATION
'C    ASSUMES CURRENT MORTALITY HAS SAME N CONC. AS LIVING COMPONENT)
'C    REMOBILIZATION FROM COARSE ROOTS IMPLIED IN COEF'S ,BNA,BNB.
    UPNAG = (XN(3) + XN(4) + XN(5) + XN(6) + XN(7) + XN(8) + XN(16)) - (XNMN1(3) + XNMN1(4) + XNMN1(5) + XNMN1(6) + XNMN1(7) + XNMN1(8) + XNMN1(16))
    UPNBG = (XN(9) + XN(14)) - (XNMN1(9) + XNMN1(14))
    UPNFF = (XN(10) + XN(11) + XN(12) + XN(13)) - (XNMN1(10) + XNMN1(11) + XNMN1(12) + XNMN1(13))
    DeficitN = UPNAG + UPNBG + UPNFF
    UPPAG = (XP(3) + XP(4) + XP(5) + XP(6) + XP(7) + XP(8) + XP(16)) - (XPMN1(3) + XPMN1(4) + XPMN1(5) + XPMN1(6) + XPMN1(7) + XPMN1(8) + XPMN1(16))
    UPPBG = (XP(9) + XP(14)) - (XPMN1(9) + XPMN1(14))
    UPPFF = (XP(10) + XP(11) + XP(12) + XP(13)) - (XPMN1(10) + XPMN1(11) + XPMN1(12) + XPMN1(13))
    DeficitP = UPPAG + UPPBG + UPPFF
'      WRITE(10,1)FNC1,FNC2,FNC3,BNC,SNC,T,TBRCH,DBN,DSN,REMXN3,REMXN4,
'1     REMXN5 , UPFAS, UPWDY
'   1  FORMAT(1H ,14F8.3)
'      Return
'      End
End Sub

Function Puruki_Stemwd_Nconc(T)
'   Estimate nitrogen concentration of stem wood at Puruki at age t years
'    Puruki_Stemwd_Nconc = 1.0044 * (T + 2.753) ^ (-0.848)
    Puruki_Stemwd_Nconc = 0.0386 + 17.69 * (T + 8.074) ^ (-1.884)
End Function

Function Puruki_Stembk_Nconc(T)
'   Estimate nitrogen concentration of stem bark at Puruki at age t years
    Puruki_Stembk_Nconc = 9.8398 * (T + 6.199) ^ (-1.0672)
End Function

Function Puruki_Livebr_Nconc(T)
'   Estimate nitrogen concentration of live branches at Puruki at age t years
    Puruki_Livebr_Nconc = 1.4806 - 1.2025 * (1 - Exp(-0.5452 * T))
End Function

Function Puruki_Deadbr_Nconc(T)
'   Estimate nitrogen concentration of dead branches at Puruki at age t years
    Puruki_Deadbr_Nconc = 3.5524 - 3.2556 * (1 - Exp(-0.5595 * T))
End Function

Function Puruki_Fol_1yr_Nconc(T)
'   Estimate nitrogen concentration of 1-yr foliage at Puruki at age t years
    Puruki_Fol_1yr_Nconc = 1.5984
End Function

Function Puruki_Fol_2yr_Nconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_Fol_2yr_Nconc = 1.4859 - 1 / (2.232 + 0.2803 * (T - 5) + 0.8822 * (T - 5) ^ 2)
    If Puruki_Fol_2yr_Nconc < 0.975 Then Puruki_Fol_2yr_Nconc = 0.975 'Ensure that N conc in 2-yr foliage is not lower than in needle fall
End Function

Function Puruki_Fol_3yr_Nconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_Fol_3yr_Nconc = 1.2939 - 1 / (2.277 - 4.0399 * (T - 5) + 5.6249 * (T - 5) ^ 2)
    If Puruki_Fol_3yr_Nconc < 0.975 Then Puruki_Fol_3yr_Nconc = 0.975 'Ensure that N conc in 3-yr foliage is not lower than in needle fall
End Function

Function Puruki_needle_fall_Nconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_needle_fall_Nconc = 0.975
End Function

Function Nconc(i, T)
'   Estimate nitrogen concentration of component i at age t years
    Dim SoilC As Double, SoilN As Double, CN_adj As Double, Age13_Nconc As Double, Nconc_2yr As Double, Nconc_needle_fall As Double
    'Soil variables required for %N and %P component models
    SoilC = Worksheets("C Change").Cells(35, 3)
    SoilN = Worksheets("C Change").Cells(36, 3)
    'Default values
    If SoilC = 0 Then SoilC = 5.57
    If SoilN = 0 Then SoilN = 0.296
    'Adjusted C/N ratio
    If SoilN <= 0.014 Then CN_adj = 50 Else CN_adj = SoilC / (SoilN - 0.014)
    If CN_adj > 50 Then CN_adj = 50
    If i = 3 Then '1-yr foliage N concentration - regression model based on Jackson data with adjusted intercept based on GxE trial data
        Age13_Nconc = 1.9015 - 0.01791 * CN_adj
        Nconc = Puruki_Fol_1yr_Nconc(T) * Age13_Nconc / Puruki_Fol_1yr_Nconc(13)
    End If
    If i = 4 Or i = 14 Then '2-yr foliage N concentration, also use for fine roots
'        Age13_Nconc = 1.6725 - 0.016149 * CN_adj
        Age13_Nconc = 1.732 - 0.0177 * CN_adj
        Nconc = Puruki_Fol_2yr_Nconc(T) * Age13_Nconc / Puruki_Fol_2yr_Nconc(13)
    End If
    If i = 5 Then   '3+yr foliage N concentration
'        Age13_Nconc = 1.6725 - 0.016149 * CN_adj
        Age13_Nconc = 1.732 - 0.0177 * CN_adj
        Nconc_2yr = Puruki_Fol_2yr_Nconc(T) * Age13_Nconc / Puruki_Fol_2yr_Nconc(13)    '2-yr foliage Nconc
        Nconc_needle_fall = (1.5105 - 0.019599 * CN_adj) * 0.828    'Needle fall
        Nconc = Nconc_2yr - (Nconc_2yr - Nconc_needle_fall) * _
            (Puruki_Fol_2yr_Nconc(T) - Puruki_Fol_3yr_Nconc(T)) / (Puruki_Fol_2yr_Nconc(T) - Puruki_needle_fall_Nconc(T))
    End If
    If i = 10 Then  'needle fall N concentration
        Nconc = (1.5105 - 0.019599 * CN_adj) * 0.828     'CN_adj regression adjusted downwards to reflect better sampling techniques used at Puruki
            'This adjustment = 0.975 / Predicted_Litter_N, i.e., the measured N conc in litter fall at Puruki divided by the predicted N in litter
                'at Puruki based on the GxE trial regression using the Puruki adjusted C/N ratio = 16.99
    End If
    If i = 16 Then  'stem bark N concentration
'        Age13_Nconc = 0.4566 - 0.0043832 * CN_adj
        Age13_Nconc = 0.452 - 0.00402 * CN_adj
        Nconc = Puruki_Stembk_Nconc(T) * Age13_Nconc / Puruki_Stembk_Nconc(13)
    End If
    If i = 6 Then   'live branch N concentration
'        Age13_Nconc = 0.0858 - 0.001193 * CN_adj  'firstly obtain Nconc for stem wood
        Age13_Nconc = 0.0844 - 0.001 * CN_adj    'firstly obtain Nconc for stem wood
        Nconc = (Age13_Nconc / Puruki_Stemwd_Nconc(13)) * Puruki_Livebr_Nconc(T)
    End If
    If i = 7 Or i = 11 Or i = 13 Then   'dead branch N concentration, also use for initial concentrations of branch and coarse root litter
'        Age13_Nconc = 0.0858 - 0.001193 * CN_adj  'firstly obtain Nconc for stem wood
        Age13_Nconc = 0.0844 - 0.001 * CN_adj    'firstly obtain Nconc for stem wood
        Nconc = (Age13_Nconc / Puruki_Stemwd_Nconc(13)) * Puruki_Deadbr_Nconc(T)
    End If
    If i = 8 Or i = 9 Then  'stem wood N concentration, also use for coarse roots
'        Age13_Nconc = 0.0858 - 0.001193 * CN_adj
        Age13_Nconc = 0.0844 - 0.001 * CN_adj
        Nconc = Puruki_Stemwd_Nconc(T) * Age13_Nconc / Puruki_Stemwd_Nconc(13)
    End If
End Function


Function Puruki_Stemwd_Pconc(T)
'   Estimate nitrogen concentration of stem wood at Puruki at age t years
'    Puruki_Stemwd_Pconc = 0.358 * (T + 3.1561) ^ (-1.1618)
    Puruki_Stemwd_Pconc = 0.0136 + 48.6 * (T + 10.56) ^ (-2.998)
End Function

Function Puruki_Stembk_Pconc(T)
'   Estimate nitrogen concentration of stem bark at Puruki at age t years
    Puruki_Stembk_Pconc = 1.1094 * (T + 3.7653) ^ (-1.1175)
End Function

Function Puruki_Livebr_Pconc(T)
'   Estimate nitrogen concentration of live branches at Puruki at age t years
    Puruki_Livebr_Pconc = 0.2907 - 0.2435 * (1 - Exp(-0.5442 * T))
End Function

Function Puruki_Deadbr_Pconc(T)
'   Estimate nitrogen concentration of live branches at Puruki at age t years
    Puruki_Deadbr_Pconc = 0.162 - 0.1376 * (1 - Exp(-0.2642 * T))
End Function

Function Puruki_Fol_1yr_Pconc(T)
'   Estimate nitrogen concentration of 1-yr foliage at Puruki at age t years
    Puruki_Fol_1yr_Pconc = 0.1848
End Function

Function Puruki_Fol_2yr_Pconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_Fol_2yr_Pconc = 0.1546 - 1 / (29.8995 - 42.7266 * (T - 5) + 39.5945 * (T - 5) ^ 2)
    If Puruki_Fol_2yr_Pconc < 0.137 Then Puruki_Fol_2yr_Pconc = 0.137 'Ensure that N conc in 2-yr foliage is not lower than in 3-yr foliage
End Function

Function Puruki_Fol_3yr_Pconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_Fol_3yr_Pconc = 0.137
End Function

Function Puruki_needle_fall_Pconc(T)
'   Estimate nitrogen concentration of 2-yr foliage at Puruki at age t years
    Puruki_needle_fall_Pconc = 0.0827
End Function


Function Pconc(i, T)
'   Estimate nitrogen concentration of component i at age t years
    Dim SoilC As Double, SoilOrganicP As Double, ln_CP As Double, Age13_Pconc As Double, Pconc_2yr As Double, Pconc_needle_fall As Double
    'Soil variables required for %N and %P component models
    SoilC = Worksheets("C Change").Cells(35, 3)
    SoilOrganicP = Worksheets("C Change").Cells(37, 3)
    'Default values
    If SoilC = 0 Then SoilC = 5.57
    If SoilOrganicP = 0 Then SoilOrganicP = 333
    ln_CP = Log(SoilC / (SoilOrganicP / 10000))
    If i = 3 Then '1-yr foliage P concentration - regression model based on Jackson data with adjusted intercept based on GxE trial data
'        Age13_Pconc = -0.35538 - 0.037992 * ln_CP
        Age13_Pconc = 0.332 - 0.0362 * ln_CP
        Pconc = Puruki_Fol_1yr_Pconc(T) * Age13_Pconc / Puruki_Fol_1yr_Pconc(13)
    End If
    If i = 4 Or i = 14 Then '2-yr foliage P concentration, also use for fine roots
'        Age13_Pconc = -0.26621 - 0.030253 * ln_CP
        Age13_Pconc = 0.278 - 0.0284 * ln_CP
        Pconc = Puruki_Fol_2yr_Pconc(T) * Age13_Pconc / Puruki_Fol_2yr_Pconc(13)
    End If
    If i = 5 Then   '3+yr foliage P concentration
'        Age13_Pconc = -0.26621 - 0.030253 * ln_CP
        Age13_Pconc = 0.278 - 0.0284 * ln_CP
        Pconc = Puruki_Fol_2yr_Pconc(T) * Age13_Pconc / Puruki_Fol_2yr_Pconc(13)    '2-yr foliage Nconc
        Pconc_needle_fall = 0.0827    'Needle fall %P at Puruki - note that litter %P was not related to soil(C/P) and averaged 0.084 for the GxE trials
            '- so the Puruki value can be used without adjustment
        Pconc = Pconc_2yr - (Pconc_2yr - Pconc_needle_fall) * _
            (Puruki_Fol_2yr_Pconc(T) - Puruki_Fol_3yr_Pconc(T)) / (Puruki_Fol_2yr_Pconc(T) - Puruki_needle_fall_Pconc(T))
    End If
    If i = 10 Then  'needle fall P concentration
        Pconc = 0.0827 'Needle fall %P at Puruki - note that litter %P was not related to soil(C/P) and averaged 0.084 for the GxE trials
            '- so the Puruki value can be used without adjustment
    End If
    If i = 16 Then  'stem bark P concentration
        Pconc = Puruki_Stembk_Pconc(T)
    End If
    If i = 6 Then   'live branch N concentration
        Pconc = Puruki_Livebr_Pconc(T)
    End If
    If i = 7 Or i = 11 Or i = 13 Then   'dead branch N concentration, also use for initial concentrations of branch and coarse root litter
        Pconc = Puruki_Deadbr_Pconc(T)
    End If
    If i = 8 Or i = 9 Then  'stem wood N concentration, also use for coarse roots
        Pconc = Puruki_Stemwd_Pconc(13)
    End If
End Function

