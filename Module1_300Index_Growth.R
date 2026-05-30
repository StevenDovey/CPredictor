# Module1_300Index_Growth.R
# 300 Index Growth Model and Wood Density Model - RADNAT v1.09
# Structural R translation of VBA Module1.bas
# All Cells(row, col) references map to data_300_indexX[row, col]
# Output written to output_300 matrix (rows = outputline, cols = column)

MODULE1_ENV <- .GlobalEnv

# ---------------------------------------------------------------------------
# Module-level variable declarations (Public)
# ---------------------------------------------------------------------------
implementation      <- 0L
SI                  <- 0
I300                <- 0
ha                  <- 0
hb                  <- 0
D200                <- 0
A200                <- 0
N_MaxBA             <- 0
site_effect         <- 0
DBHsqd_add_offset   <- 0
DBHsqd_mult_offset  <- 1
MTH_add_offset      <- 0
MTH_mult_offset     <- 1
DBH_calibration_age <- 0
MTH_calibration_age <- 0

# Module-level variable declarations (Private)
adjage          <- 0
yieldline       <- 0L
predicted       <- FALSE
agezero         <- 0
rotlth          <- 0
startplotrow    <- 0L
plotI300        <- 0
SoilC           <- 0
SoilN           <- 0
Temp            <- 0
GeneticAdj      <- 0
CoreDens        <- 0
CoreAge         <- 0L
InnerRing       <- 0L
OuterRing       <- 0L
drift           <- 0
sum_ba          <- 0
sum_height      <- 0
N               <- 0
Thin            <- 0L
Nshist          <- 0L
outputline      <- 5L
last_outputline <- 5L
mth             <- 0
Meanht          <- numeric(10)
prht            <- numeric(10)
prlag           <- numeric(10)
ncum            <- numeric(10)
maxage          <- 0
Nlifts          <- 0L
Nelements       <- 1L
Nthins          <- 0L
vol             <- 0
ThinLag         <- numeric(8)
sellag          <- numeric(10)
adjageel        <- numeric(10)
thinage         <- FALSE
nelement        <- numeric(10)
age             <- 0
dbh             <- 0
dbhsqd          <- 0
dbhelement      <- numeric(10)
agethin         <- numeric(8)
initiallag      <- numeric(8)
totalthinlag    <- 0
BA              <- 0
crlth           <- 0
shist_T         <- numeric(999)
shist_N1        <- numeric(999)
shist_N2        <- numeric(999)
shist_thincoeff <- numeric(999)
Initialstocking <- 0
steplength      <- 0
steps           <- 0L
lift_T          <- numeric(999)
lift_height     <- numeric(999)
Npruned         <- 0
Mortality       <- numeric(999)
OUTPUT          <- FALSE
shist           <- 1L
j               <- 0L
el              <- 0L
lineprinted     <- FALSE
lift            <- 0L
prevN           <- 0
total_prlag     <- numeric(10)
prevprht        <- 0
voltabarray     <- character(11)
DBH300          <- 0
shist_thinratio <- numeric(999)
age_300Index    <- 0
agediff         <- 0
prevagediff     <- 0
lift_sph        <- numeric(10)
v               <- matrix(0, nrow = 11, ncol = 8)
voltable        <- 0L
attrition       <- 0
pctmortadj      <- 0
X               <- 0
sdi             <- 0
lift_prunecoeff <- numeric(10)
heightmodel     <- 0L
mortmodel       <- 0L
mnheight        <- 0
specage         <- 0
predage         <- 0
altitude        <- 0
latitude        <- 0
indexage        <- 0
bias_young      <- FALSE
bias_old        <- FALSE
bias_SI         <- FALSE

# Output matrix: rows 1..200, cols 1..70 (maps directly to VBA Cells indices)
output_300 <- matrix(0, nrow = 200, ncol = 70)

# ---------------------------------------------------------------------------
# Model coefficients (Private Const in VBA)
# ---------------------------------------------------------------------------
da1       <- 56.523
db1       <- -0.09045
dr        <- 2.6416
dl        <- 28.1224
dc        <- 1.4821
dbSI      <- -0.00212
dn        <- 15.7581
dm        <- -0.00455
dbdia     <- -0.1325
Ds        <- 0.1702
dbsidia   <- -0.0084
drsi      <- 0.0209
dr2       <- 0.8234
pra       <- 0.0934
prb       <- 1.98
prc       <- 0.2119
ha0       <- -2.475
ha1       <- -0.01406
hb0       <- 0.33417
hb1       <- 0.0104
hae0      <- -1.335
hae1      <- -0.03581
hae2      <- -0.0006306
hbe0      <- 0.499
hbe1      <- 0.005059
hNSWa     <- -2.6842
hNSWb     <- 0.7293
hNSWp     <- -0.00176
thb       <- 0.5
thc       <- -0.47
mortd     <- 0.2493
tha       <- 0.5
db2       <- 1
thincoeff <- 0.784
mortc     <- 1.5
morte     <- -0.0555
mortNSW   <- 0.869
morta     <- 0.000688
mortb     <- -14.91
mortp     <- -44.691
mortq     <- -4.611
morts     <- 3.901
mortt     <- 1.3533
mortu     <- 0.00246
mortv     <- -30.565
mortw     <- 2.536
mortx     <- 1.125
morty     <- 0.000438
morta1    <- 0.00206
mortb1    <- -46.3216
mortc1    <- 3.1704
mortd1    <- 1.7477
morte1    <- -0.1631
mortf1    <- 0.1991
mort2007_a <- 0.000459
mort2007_b <- 0.974
mort2007_c <- 3.06
mort2007_d <- 0.786
mort2007_f <- -0.037
mort2007_g <- 0.0371
mort2007_h <- -0.32

# ---------------------------------------------------------------------------
# voltab: Assign values to volume table coefficients
# ---------------------------------------------------------------------------
voltab <- function() {
    v[1, 1] <<- 0.942;    v[1, 2] <<- -1.161;    v[1, 3] <<- 0.317
    v[2, 1] <<- 0.989;    v[2, 2] <<- -1.2752;   v[2, 3] <<- 0.3191
    v[3, 1] <<- 1.492912924;  v[3, 2] <<- -0.999113309;  v[3, 3] <<- 1.250753941
    v[3, 4] <<- -0.397037159; v[3, 5] <<- 0.027218164;   v[3, 6] <<- -0.063166205
    v[3, 7] <<- 0.064609459;  v[3, 8] <<- -0.030665365
    v[4, 1] <<- 1.633105986;  v[4, 2] <<- -1.039327204;  v[4, 3] <<- 1.212696953
    v[4, 4] <<- -0.359131176; v[4, 5] <<- 0.026454943;   v[4, 6] <<- -0.067457458
    v[4, 7] <<- 0.066992488;  v[4, 8] <<- -0.030528278
    v[5, 1] <<- 0.730448717;  v[5, 2] <<- -0.617440226;  v[5, 3] <<- 1.095616037
    v[5, 4] <<- -0.222220223; v[5, 5] <<- 0.013858949;   v[5, 6] <<- -0.11022445
    v[5, 7] <<- 0.059157535;  v[5, 8] <<- -0.016942593
    v[6, 1] <<- 1.09857999;   v[6, 2] <<- -0.883862258;  v[6, 3] <<- 1.165375013
    v[6, 4] <<- -0.28047221;  v[6, 5] <<- 0.022081234;   v[6, 6] <<- -0.059261776
    v[6, 7] <<- 0.053187392;  v[6, 8] <<- -0.025226521
    v[7, 1] <<- 1.403009551;  v[7, 2] <<- -0.96392392;   v[7, 3] <<- 1.221046594
    v[7, 4] <<- -0.358337009; v[7, 5] <<- 0.024975712;   v[7, 6] <<- -0.061374804
    v[7, 7] <<- 0.061895757;  v[7, 8] <<- -0.028672533
    v[8, 1] <<- 2.834246614;  v[8, 2] <<- -1.856804825;  v[8, 3] <<- 1.152097786
    v[8, 4] <<- -0.201346156; v[8, 5] <<- -0.000721117;  v[8, 6] <<- 0.081503044
    v[8, 7] <<- 0.024428222;  v[8, 8] <<- 0.001938887
    v[9, 1] <<- 2.7023;   v[9, 2] <<- -2.1301;  v[9, 3] <<- 1.3901
    v[9, 4] <<- -0.5056;  v[9, 5] <<- 0.0548;   v[9, 6] <<- 0.0991
    v[9, 7] <<- 0.1478;   v[9, 8] <<- -0.088
    v[10, 1] <<- 6.2733;  v[10, 2] <<- 0.1284;  v[10, 3] <<- -0.00097
    v[11, 1] <<- 2.1819;  v[11, 2] <<- 0.2504;  v[11, 3] <<- -0.00081
}

# ---------------------------------------------------------------------------
# Inputparms: Read stand parameters from data_300_indexX
# ---------------------------------------------------------------------------
Inputparms <- function() {
    implementation <<- as.integer(as.numeric(data_300_indexX[8, 6]))
    I300 <<- as.numeric(data_300_indexX[3, 3])
    SI   <<- as.numeric(data_300_indexX[4, 3])
    Initialstocking <<- as.numeric(data_300_indexX[19, 3])
    drift <<- as.numeric(data_300_indexX[64, 6])

    bias_old   <<- FALSE
    bias_young <<- FALSE
    bias_SI    <<- FALSE
    if (tolower(data_300_indexX[51, 6]) == "x") bias_old   <<- TRUE
    if (tolower(data_300_indexX[52, 6]) == "x") bias_young <<- TRUE
    if (tolower(data_300_indexX[53, 6]) == "x") bias_SI    <<- TRUE

    maxage     <<- as.numeric(data_300_indexX[47, 3])
    steplength <<- as.numeric(data_300_indexX[48, 3])
    if (steplength < 0.01) steplength <<- 0.01

    heightmodel <<- heightmod()

    if (tolower(data_300_indexX[68, 4]) == "x") {
        mortmodel <<- 1L
    } else if (tolower(data_300_indexX[69, 4]) == "x") {
        mortmodel <<- 2L
    } else if (tolower(data_300_indexX[70, 4]) == "x") {
        mortmodel <<- 3L
    } else if (tolower(data_300_indexX[71, 4]) == "x") {
        mortmodel <<- 5L
    } else {
        mortmodel <<- 6L
    }

    if (mortmodel >= 4L) {
        if (!is.na(data_300_indexX[68, 6]) && as.numeric(data_300_indexX[68, 6]) != 0) {
            attrition <<- as.numeric(data_300_indexX[68, 6]) / 100
        } else if (mortmodel == 4L) {
            attrition <<- mortu
        } else if (mortmodel == 5L) {
            attrition <<- morta1
        } else {
            attrition <<- 0
        }
        if (!is.na(data_300_indexX[69, 6]) && as.numeric(data_300_indexX[69, 6]) != 0) {
            pctmortadj <<- as.numeric(data_300_indexX[69, 6])
        } else {
            pctmortadj <<- 0
        }
    }

    for (i in 1:11) {
        voltabarray[i] <<- data_300_indexX[50 + i, 4]
    }

    if      (tolower(voltabarray[1])  == "x") { voltable <<- 1L
    } else if (tolower(voltabarray[2]) == "x") { voltable <<- 2L
    } else if (tolower(voltabarray[3]) == "x") { voltable <<- 3L
    } else if (tolower(voltabarray[4]) == "x") { voltable <<- 4L
    } else if (tolower(voltabarray[5]) == "x") { voltable <<- 5L
    } else if (tolower(voltabarray[6]) == "x") { voltable <<- 6L
    } else if (tolower(voltabarray[7]) == "x") { voltable <<- 7L
    } else if (tolower(voltabarray[8]) == "x") { voltable <<- 8L
    } else if (tolower(voltabarray[9]) == "x") { voltable <<- 9L
    } else if (tolower(voltabarray[10]) == "x") { voltable <<- 10L
    } else if (tolower(voltabarray[11]) == "x") { voltable <<- 11L
    }

    steps <<- as.integer(maxage / steplength)
    calcheightcoeff(SI)

    for (s in 1:17) {
        shist_T[s]         <<- 0
        shist_N1[s]        <<- 0
        shist_N2[s]        <<- 0
        shist_thinratio[s] <<- 0
        shist_thincoeff[s] <<- 0
        Mortality[s]       <<- 0
    }
    for (lft in 1:5) {
        lift_T[lft]         <<- 0
        lift_height[lft]    <<- 0
        lift_sph[lft]       <<- 0
        lift_prunecoeff[lft] <<- 0
    }

    startline <- 19L
    nlines    <- 17L
    Nshist <<- 0L
    for (s in 1:nlines) {
        val <- data_300_indexX[startline + s, 2]
        if (!is.na(val) && val != "") {
            shist_T[s]         <<- as.numeric(data_300_indexX[startline + s, 2])
            shist_N1[s]        <<- as.numeric(data_300_indexX[startline + s, 3])
            shist_N2[s]        <<- as.numeric(data_300_indexX[startline + s, 4])
            shist_thincoeff[s] <<- as.numeric(data_300_indexX[startline + s, 5])
            shist_thinratio[s] <<- as.numeric(data_300_indexX[startline + s, 6])
            Nshist <<- as.integer(s)
        }
    }

    mort()

    startline <- 39L
    Nlifts <<- 0L
    for (lft in 1:5) {
        val <- data_300_indexX[startline + lft, 2]
        if (!is.na(val) && val != "") {
            lift_T[lft]          <<- as.numeric(data_300_indexX[startline + lft, 2])
            lift_height[lft]     <<- as.numeric(data_300_indexX[startline + lft, 3])
            lift_sph[lft]        <<- as.numeric(data_300_indexX[startline + lft, 4])
            if (lift_sph[lft] == 0) lift_sph[lft] <<- 10000
            lift_prunecoeff[lft] <<- as.numeric(data_300_indexX[startline + lft, 5])
            Nlifts <<- as.integer(lft)
        }
    }
}

# ---------------------------------------------------------------------------
# OutputGrowth: Input I300, SI & stand history and predict growth
# ---------------------------------------------------------------------------
OutputGrowth <- function() {
    if (!checkinput_site())        return(invisible(NULL))
    if (!checkinput_SI())          return(invisible(NULL))
    if (!checkinput_htfn())        return(invisible(NULL))
    if (!checkinput_initialstock()) return(invisible(NULL))
    if (!checkinput_stocking())    return(invisible(NULL))
    if (!checkinput_prune())       return(invisible(NULL))
    if (!checkinput_fellage())     return(invisible(NULL))
    if (!checkinput_steplth())     return(invisible(NULL))
    if (!checkinput_volfn())       return(invisible(NULL))
    if (!checkinput_mortfn())      return(invisible(NULL))
    OUTPUT <<- TRUE
    output_300 <<- matrix(0, nrow = 200, ncol = 70)
    Inputparms()
    voltab()
    if (implementation == 2L) {
        CalcOffsets()
        Inputparms()
        OUTPUT <<- TRUE
    }
    Growth(OUTPUT)
    earlyield()
    mortvol()
    density()
}

# ---------------------------------------------------------------------------
# Calc300Index: Calculate 300 Index from plot measurement
# ---------------------------------------------------------------------------
Calc300Index <- function() {
    if (!checkinput_I300())        return(invisible(NULL))
    if (!checkinput_SI())          return(invisible(NULL))
    if (!checkinput_htfn())        return(invisible(NULL))
    if (!checkinput_initialstock()) return(invisible(NULL))
    if (!checkinput_stocking())    return(invisible(NULL))
    if (!checkinput_prune())       return(invisible(NULL))
    if (!checkinput_fellage())     return(invisible(NULL))
    if (!checkinput_steplth())     return(invisible(NULL))
    if (!checkinput_volfn())       return(invisible(NULL))
    if (!checkinput_mortfn())      return(invisible(NULL))
    Inputparms()
    voltab()
    age300   <- as.numeric(data_300_indexX[7, 3])
    Stock300 <- as.numeric(data_300_indexX[8, 3])
    maxage   <<- age300
    steps    <<- as.integer(maxage / steplength)
    MTH300   <- CalcMTH(SI, age300)
    DBH300   <<- as.numeric(data_300_indexX[9, 3])
    if (DBH300 == 0) {
        BA300 <- as.numeric(data_300_indexX[10, 3])
        if (BA300 != 0) {
            DBH300 <<- CalcDBHfromBA(BA300, Stock300)
        } else {
            Vol300 <- as.numeric(data_300_indexX[11, 3])
            DBH300 <<- CalcDBHfromBA(calcBAfromVol(MTH300, Vol300, Stock300), Stock300)
        }
    }
    Index300()
    data_300_indexX[3, 3] <<- as.character(I300)
    OutputGrowth()
}

# ---------------------------------------------------------------------------
# Index300: Calculate 300 Index using bisection method
# ---------------------------------------------------------------------------
Index300 <- function() {
    I300 <<- Bisection(1.328, 60, 14, 1, 0, 0, 0, 0)
}

# ---------------------------------------------------------------------------
# Growth: Predict growth from given stand parameters
# ---------------------------------------------------------------------------
Growth <- function(output_flag) {
    N          <<- Initialstocking
    shist      <<- 1L
    Thin       <<- 0L
    lift       <<- 0L
    age        <<- 0
    Nelements  <<- 1L
    A200       <<- CalcA200start(age, I300, SI)
    dbh        <<- 0
    for (el_i in 1:10) {
        dbhelement[el_i] <<- 0
    }
    vol       <<- 0.0000064 * N
    BA        <<- 0
    mth       <<- 0.25
    mnheight  <<- 0.25
    nelement[1]    <<- N
    ncum[1]        <<- N
    prht[1]        <<- 0
    prlag[1]       <<- 0
    totalthinlag   <<- 0
    sellag[1]      <<- 0
    total_prlag[1] <<- 0
    outputline  <<- 5L
    lineprinted <<- FALSE
    if (implementation != 2L) {
        DBHsqd_add_offset   <<- 0
        DBHsqd_mult_offset  <<- 1
        MTH_add_offset      <<- 0
        MTH_mult_offset     <<- 1
        DBH_calibration_age <<- 0
        MTH_calibration_age <<- 0
    }
    if (output_flag) OutStep()

    for (jj in 1:steps) {
        tl_prev_standDBH      <- dbh
        tl_prev_standN        <- N
        tl_prev_standBA       <- BA
        tl_prev_standmnheight <- mnheight
        tl_prev_standage      <- age
        age <<- age + steplength
        A200 <<- CalcA200start(age, I300, SI)
        stock()
        Height()
        Ageshifts()
        Diameter()
        VolBA()
        if ((output_flag && (age - floor(age) < 0.001 || age - floor(age) > 0.999)) ||
            abs(age - maxage) < 0.001) OutStep()
        if (shist_N2[shist] != 0 && age >= shist_T[shist] - 0.001) {
            if (output_flag && !lineprinted) OutStep()
            Thin <<- Thin + 1L
            thinning()
            if (output_flag) {
                OutThin()
                OutElements()
            }
        }
        if (lift < Nlifts && age >= lift_T[lift + 1] - 0.001) {
            if (output_flag && !lineprinted) OutStep()
            lift <<- lift + 1L
            Newlift()
            if (output_flag) {
                OutPrune()
                OutElements()
            }
        }
        if (shist < Nshist && age >= shist_T[shist] - 0.001) shist <<- shist + 1L
        if (lineprinted) {
            outputline  <<- outputline + 1L
            lineprinted <<- FALSE
        }
    }
    last_outputline <<- outputline - 1L
}

# ---------------------------------------------------------------------------
# OutStep: Output yield for a single prediction iteration
# ---------------------------------------------------------------------------
OutStep <- function() {
    offset_corrected_dbh <- 0
    offset_corrected_MTH <- 0
    if (age < MTH_calibration_age) {
        offset_corrected_MTH <- mth * MTH_mult_offset
    } else {
        offset_corrected_MTH <- mth + MTH_add_offset
    }
    if (age < DBH_calibration_age) {
        offset_corrected_dbh <- sqrt(dbh^2 * DBHsqd_mult_offset)
    } else {
        offset_corrected_dbh <- sqrt(dbh^2 + DBHsqd_add_offset)
    }
    output_300[outputline,  7] <<- round(age, 2)
    output_300[outputline,  8] <<- N
    output_300[outputline, 10] <<- offset_corrected_MTH
    output_300[outputline, 16] <<- offset_corrected_dbh
    output_300[outputline, 14] <<- CalcBAfromDBH(offset_corrected_dbh, N)
    output_300[outputline, 12] <<- CalcVol(offset_corrected_MTH, CalcBAfromDBH(offset_corrected_dbh, N), N)
    output_300[outputline, 18] <<- calcMeanht(offset_corrected_MTH, N)
    OutElements()
    output_300[outputline, 43] <<- DBHmodel(A200, SI, 20, N)
    lineprinted <<- TRUE
}

# ---------------------------------------------------------------------------
# OutElements: Output stocking, mean DBH and volume of each pruned element
# ---------------------------------------------------------------------------
OutElements <- function() {
    for (col_i in 19:36) output_300[outputline, col_i] <<- 0
    for (el_i in 1:Nelements) {
        output_300[outputline, 16 + el_i * 3]     <<- nelement[el_i]
        output_300[outputline, 16 + el_i * 3 + 1] <<- dbhelement[el_i]
        output_300[outputline, 16 + el_i * 3 + 2] <<-
            CalcVol(mth, CalcBAfromDBH(dbhelement[el_i], nelement[el_i]), nelement[el_i])
    }
}

# ---------------------------------------------------------------------------
# OutThin: Output stand parameters following thinning
# ---------------------------------------------------------------------------
OutThin <- function() {
    offset_corrected_dbh <- 0
    offset_corrected_MTH <- 0
    if (age < MTH_calibration_age) {
        offset_corrected_MTH <- mth * MTH_mult_offset
    } else {
        offset_corrected_MTH <- mth + MTH_add_offset
    }
    if (age < DBH_calibration_age) {
        offset_corrected_dbh <- sqrt(dbh^2 * DBHsqd_mult_offset)
    } else {
        offset_corrected_dbh <- sqrt(dbh^2 + DBHsqd_add_offset)
    }
    output_300[outputline,  9] <<- N
    output_300[outputline, 17] <<- offset_corrected_dbh
    output_300[outputline, 15] <<- CalcBAfromDBH(offset_corrected_dbh, N)
    output_300[outputline, 13] <<- CalcVol(offset_corrected_MTH, CalcBAfromDBH(offset_corrected_dbh, N), N)
    output_300[outputline, 18] <<- calcMeanht(offset_corrected_MTH, N)
    for (el_i in 1:Nelements) {
        output_300[outputline, 16 + el_i * 3]     <<- nelement[el_i]
        output_300[outputline, 16 + el_i * 3 + 1] <<- dbhelement[el_i]
        output_300[outputline, 16 + el_i * 3 + 2] <<-
            CalcVol(mth, CalcBAfromDBH(dbhelement[el_i], nelement[el_i]), nelement[el_i])
    }
    output_300[outputline, 43] <<- DBHmodel(A200, SI, 20, N)
}

# ---------------------------------------------------------------------------
# OutPrune: Output crown length for a pruning lift
# ---------------------------------------------------------------------------
OutPrune <- function() {
    output_300[outputline, 11] <<- crlth
    if (lift_sph[lift] == 10000) {
        output_300[outputline, 37] <<- N
    } else {
        output_300[outputline, 37] <<- lift_sph[lift]
    }
    output_300[outputline, 38] <<- lift_height[lift]
    output_300[outputline, 43] <<- DBHmodel(A200, SI, 20, N)
}

# ---------------------------------------------------------------------------
# mort: Calculate mortalities for each stocking history interval
# ---------------------------------------------------------------------------
mort <- function() {
    if (shist_T[Nshist] < maxage) {
        Nshist <<- Nshist + 1L
        shist_T[Nshist] <<- maxage
    }
    prevage <- 0
    prevN_m <- Initialstocking
    for (s in 1:Nshist) {
        if (is.na(shist_N1[s]) || shist_N1[s] == 0) {
            Mortality[s] <<- -1
        } else {
            Mortality[s] <<- 100 * log(prevN_m / shist_N1[s]) / (shist_T[s] - prevage)
        }
        prevage <- shist_T[s]
        if (is.na(shist_N2[s]) || shist_N2[s] == 0) {
            prevN_m <- shist_N1[s]
        } else {
            prevN_m <- shist_N2[s]
        }
    }
}

# ---------------------------------------------------------------------------
# stock: Generate stocking using mortality function
# ---------------------------------------------------------------------------
stock <- function() {
    prevN <<- N
    if (Mortality[shist] >= 0) {
        N <<- prevN / exp(Mortality[shist] * steplength / 100)
    } else if (mortmodel == 1L) {
        mortrate <- mortNSW
        N <<- prevN / exp(mortrate * steplength / 100)
    } else if (mortmodel == 2L) {
        if (dbh == 0) {
            mortrate <- 0
        } else {
            X_loc <- exp(mortb + morte * SI + mortc * (log(N) + mortd * log(dbh^2)))
            mortrate <- (morta + (1 - morta) * X_loc / (1 + X_loc)) * 100
            N <<- prevN / exp(mortrate * steplength / 100)
        }
    } else if (mortmodel == 3L) {
        if (dbh == 0) {
            mortrate <- 0
        } else {
            X_loc <- exp(mortv + mortw * (log(N) + mortx * log(dbh)))
            mortrate <- (morty + (1 - morty) * X_loc / (1 + X_loc)) * 100
            N <<- prevN / exp(mortrate * steplength / 100)
        }
    } else if (mortmodel == 4L) {
        if (dbh == 0) {
            mortrate <- 0
        } else {
            X_loc <- exp(mortp + mortq * I300 / SI + morts * (log(N) + mortt * log(dbh)))
            mortrate <- (attrition + (1 - attrition) * X_loc / (1 + X_loc)) * 100
            N <<- prevN / exp(mortrate * steplength / 100)
        }
    } else if (mortmodel == 5L) {
        if (dbh == 0) {
            mortrate <- 0
        } else {
            X_loc <- exp(mortb1 + morte1 * I300 + mortf1 * SI + mortc1 * (log(N) + mortd1 * log(dbh)))
            mortrate <- (attrition + (1 - attrition) * X_loc / (1 + X_loc)) * 100
            N <<- prevN / exp(mortrate * steplength / 100)
        }
    } else if (mortmodel == 6L) {
        if (dbh == 0) {
            mortrate <- 0
        } else {
            sdi <<- exp(mort2007_f * I300 + mort2007_g * SI + log(N) +
                        mort2007_d * log(dbh / 100) +
                        mort2007_h * (log(dbh / 100))^2) / 1000
            mortrate <- attrition * 100 +
                        100 * (1 + pctmortadj / 100) * (mort2007_a + mort2007_b * sdi^mort2007_c)
            if (mortrate > 95) mortrate <- 95
            if (mortrate < 0)  mortrate <- 0
            N <<- prevN * (1 - mortrate / 100)^steplength
        }
    }
    for (el_i in 1:Nelements) {
        nelement[el_i] <<- nelement[el_i] * N / prevN
        ncum[el_i]     <<- ncum[el_i]     * N / prevN
    }
}

# ---------------------------------------------------------------------------
# Height: Predict MTH and mean height for a given age, SI and stocking
# ---------------------------------------------------------------------------
Height <- function() {
    mth      <<- CalcMTH(SI, age)
    mnheight <<- calcMeanht(mth, N)
    Meanht[1] <<- calcMeanht(mth, nelement[1])
    for (el_i in 2:Nelements) {
        Meanht[el_i] <<- (ncum[el_i] * calcMeanht(mth, ncum[el_i]) -
                          ncum[el_i - 1] * calcMeanht(mth, ncum[el_i - 1])) /
                         (ncum[el_i] - ncum[el_i - 1])
    }
}

# ---------------------------------------------------------------------------
# Ageshifts: Calculate pruning and thinning time shifts for each element
# ---------------------------------------------------------------------------
Ageshifts <- function() {
    for (el_i in 1:Nelements) {
        if (prht[el_i] > 0) prlag[el_i] <<- prlag[el_i] + 0.3 * steplength
        if (prlag[el_i] > total_prlag[el_i]) prlag[el_i] <<- total_prlag[el_i]
    }
    totalthinlag <<- 0
    for (th in 1:Thin) {
        timesincethin <- age - agethin[th]
        ThinLag[th] <<- initiallag[th] +
            min(initiallag[th], tha) * thb * (1 - exp(thc * timesincethin))
        totalthinlag <<- totalthinlag + ThinLag[th]
    }
    adjage <<- 0
    for (el_i in 1:Nelements) {
        adjageel[el_i] <<- age - prlag[el_i] - sellag[el_i] - totalthinlag
        adjage <<- adjage + adjageel[el_i] * nelement[el_i]
    }
    adjage <<- adjage / N
}

# ---------------------------------------------------------------------------
# Newlift: Calculate element means following a pruning lift
# ---------------------------------------------------------------------------
Newlift <- function() {
    if (lift_sph[lift] + 0.0001 < nelement[1]) {
        Nelements <<- Nelements + 1L
        for (el_i in Nelements:2) {
            prht[el_i]        <<- prht[el_i - 1]
            nelement[el_i]    <<- nelement[el_i - 1]
            ncum[el_i]        <<- ncum[el_i - 1]
            total_prlag[el_i] <<- total_prlag[el_i - 1]
            prlag[el_i]       <<- prlag[el_i - 1]
            dbhelement[el_i]  <<- dbhelement[el_i - 1]
            Meanht[el_i]      <<- Meanht[el_i - 1]
            adjageel[el_i]    <<- adjageel[el_i - 1]
            sellag[el_i]      <<- sellag[el_i - 1]
        }
        prht[1]     <<- lift_height[lift]
        nelement[1] <<- lift_sph[lift]
        nelement[2] <<- nelement[2] - nelement[1]
        ncum[1]     <<- nelement[1]
        Meanht[1]   <<- calcMeanht(mth, nelement[1])
        crlth       <<- Meanht[1] - prht[1]
        dbhb4pr     <- dbhelement[1]
        if (lift_prunecoeff[lift] != 0) {
            prunecoeff <- lift_prunecoeff[lift]
        } else {
            prunecoeff <- thincoeff
        }
        dbhelement[1] <<- dbhelement[1] * (nelement[1] / ncum[2])^((prunecoeff - 1) / 2)
        dbhelement[2] <<- sqrt((ncum[2] * dbhelement[2]^2 - nelement[1] * dbhelement[1]^2) /
                               nelement[2])
        sellag[1] <<- sellag[1] + adjageel[1] - CalcAge(dbhelement[1], A200, N, SI)
        sellag[2] <<- sellag[2] + adjageel[2] - CalcAge(dbhelement[2], A200, N, SI)
        adjageel[1] <<- age - prlag[1] - totalthinlag - sellag[1]
        adjageel[2] <<- age - prlag[2] - totalthinlag - sellag[2]
        total_prlag[1] <<- total_prlag[2] +
            pra * (prht[1]^prb - prht[2]^prb) * exp(-prc * crlth)
    } else {
        prevprht <<- prht[1]
        prht[1]  <<- lift_height[lift]
        crlth    <<- Meanht[1] - prht[1]
        total_prlag[1] <<- total_prlag[1] +
            pra * (prht[1]^prb - prevprht^prb) * exp(-prc * crlth)
    }
}

# ---------------------------------------------------------------------------
# Diameter: Calculate mean DBH of each element
# ---------------------------------------------------------------------------
Diameter <- function() {
    N <<- ncum[Nelements]
    dbhsqd <<- 0
    for (el_i in 1:Nelements) {
        prevdbh_el <- dbhelement[el_i]
        dbhelement[el_i] <<- CalcDBH(I300, SI, adjageel[el_i], N)
        if (dbhelement[el_i] < prevdbh_el) dbhelement[el_i] <<- prevdbh_el
        dbhsqd <<- dbhsqd + nelement[el_i] * dbhelement[el_i]^2
    }
    dbh <<- sqrt(dbhsqd / N)
}

# ---------------------------------------------------------------------------
# thinning: Predict stand after thinning
# ---------------------------------------------------------------------------
thinning <- function() {
    prevN_t  <- N
    prevdbh_t <- dbh

    if (shist_thincoeff[shist] != 0) {
        kcoeff <- shist_thincoeff[shist]
    } else {
        kcoeff <- thincoeff
    }

    thinN <- prevN_t - shist_N2[shist]
    for (el_i in Nelements:1) {
        if (thinN + 0.0001 >= nelement[el_i]) {
            thinN  <- thinN - nelement[el_i]
            Nelements <<- Nelements - 1L
        } else {
            prevNel  <- nelement[el_i]
            prevNcum <- ncum[el_i]
            nelement[el_i] <<- nelement[el_i] - thinN
            ncum[el_i]     <<- ncum[el_i] - thinN
            if (el_i != 1L) {
                dbhelement[el_i] <<- dbhelement[el_i] *
                    (ncum[el_i]^((kcoeff + 1) / 2) - ncum[el_i - 1]^((kcoeff + 1) / 2)) *
                    (prevNcum - ncum[el_i - 1]) /
                    ((prevNcum^((kcoeff + 1) / 2) - ncum[el_i - 1]^((kcoeff + 1) / 2)) *
                     (ncum[el_i] - ncum[el_i - 1]))
            } else {
                dbhelement[el_i] <<- dbhelement[el_i] * (ncum[el_i] / prevNcum)^((kcoeff - 1) / 2)
            }
            Nelements <<- as.integer(el_i)
            break
        }
    }

    N <<- ncum[Nelements]
    dbhsqd_t <- 0
    for (el_i in 1:Nelements) {
        dbhsqd_t <- dbhsqd_t + nelement[el_i] * dbhelement[el_i]^2
    }
    dbh <<- sqrt(dbhsqd_t / N)

    if (shist_thinratio[shist] != 0 && prevdbh_t != 0) {
        current_thinratio <- dbh / prevdbh_t
        for (el_i in 1:Nelements) {
            dbhelement[el_i] <<- dbhelement[el_i] * shist_thinratio[shist] / current_thinratio
        }
        dbh <<- dbh * shist_thinratio[shist] / current_thinratio
    }

    VolBA()

    if (prevdbh_t == 0) {
        initiallag[Thin] <<- 0
    } else {
        A200 <<- CalcA200start(adjage, I300, SI)
        initiallag[Thin] <<- adjage - CalcAge(prevdbh_t, A200, N, SI)
        A200 <<- CalcA200start(adjage - initiallag[Thin], I300, SI)
        initiallag[Thin] <<- adjage - CalcAge(prevdbh_t, A200, N, SI)
        A200 <<- CalcA200start(adjage - initiallag[Thin], I300, SI)
        initiallag[Thin] <<- adjage - CalcAge(prevdbh_t, A200, N, SI)
    }
    ThinLag[Thin]  <<- initiallag[Thin]
    agethin[Thin]  <<- age
    totalthinlag   <<- totalthinlag + initiallag[Thin]
    for (el_i in 1:Nelements) {
        if (dbhelement[el_i] == 0) {
            sellag[el_i] <<- 0
        } else {
            sellag[el_i] <<- age - prlag[el_i] - totalthinlag - CalcAge(dbhelement[el_i], A200, N, SI)
        }
        adjageel[el_i] <<- age - prlag[el_i] - totalthinlag - sellag[el_i]
    }
    adjage <<- 0
    for (el_i in 1:Nelements) {
        adjage <<- adjage + adjageel[el_i] * nelement[el_i]
    }
    adjage <<- adjage / N
    mnheight <<- calcMeanht(mth, N)
}

# ---------------------------------------------------------------------------
# VolBA: Calculate Volume and BA from DBH, MTH and stocking
# ---------------------------------------------------------------------------
VolBA <- function() {
    BA  <<- CalcBAfromDBH(dbh, N)
    vol <<- CalcVol(mth, BA, N)
}

# ---------------------------------------------------------------------------
# siteIndex: Calculate SI from height measurement using bisection
# ---------------------------------------------------------------------------
siteIndex <- function() {
    if (!checkinput_htage()) return(invisible(NULL))
    if (!checkinput_htfn())  return(invisible(NULL))
    HAge <- as.numeric(data_300_indexX[14, 3])
    HMTH <- as.numeric(data_300_indexX[15, 3])
    if (HAge == 20) {
        SI <<- HMTH
    } else {
        heightmodel <<- heightmod()
        SI <<- Bisection(5, 60, 15, 2, HMTH, HAge, 0, 0)
    }
    data_300_indexX[4, 3] <<- as.character(SI)
}

# ---------------------------------------------------------------------------
# calcheightcoeff: Calculate coefficients for height model
# ---------------------------------------------------------------------------
calcheightcoeff <- function(si_val) {
    if (heightmodel == 1L) {
        ha <<- exp(hNSWa)
        hb <<- 1 / (hNSWb + hNSWp * si_val)
    } else if (heightmodel == 2L) {
        ha <<- exp(ha0 + ha1 * si_val)
        hb <<- 1 / (hb0 + hb1 * si_val)
    } else {
        latitude <<- abs(as.numeric(data_300_indexX[3, 6]))
        altitude <<- as.numeric(data_300_indexX[4, 6])
        ha <<- exp(hae0 + hae1 * latitude + hae2 * altitude)
        latitude <<- abs(as.numeric(data_300_indexX[3, 6]))
        altitude <<- as.numeric(data_300_indexX[4, 6])
        ha <<- exp(hae0 + hae1 * latitude + hae2 * altitude)
        hb <<- 1 / (hbe0 + hbe1 * si_val)
    }
}

# ---------------------------------------------------------------------------
# Utility: BA / DBH / Volume conversions
# ---------------------------------------------------------------------------
CalcDBHfromBA <- function(BA_val, N_val) {
    sqrt(1.273 * BA_val / N_val) * 100
}

CalcBAfromDBH <- function(dbh_val, N_val) {
    N_val / 1.273 * (dbh_val / 100)^2
}

calcBAfromVol <- function(mth_val, vol_val, stock_val) {
    if (vol_val <= 0 || mth_val <= 1.6 || stock_val <= 0) {
        return(0)
    } else if (voltable == 1L || voltable == 2L) {
        return(vol_val / (mth_val * (v[voltable, 1] * (mth_val - 1.4)^v[voltable, 2] + v[voltable, 3])))
    } else if (voltable == 10L || voltable == 11L) {
        return(vol_val / (v[voltable, 1] + v[voltable, 2] * mth_val + v[voltable, 3] * stock_val))
    } else {
        return(exp(v[voltable, 1] +
                   v[voltable, 2] * log(mth_val) +
                   v[voltable, 3] * log(vol_val) +
                   v[voltable, 4] * log(stock_val) +
                   v[voltable, 5] * log(stock_val) * log(stock_val) +
                   v[voltable, 6] * log(mth_val) * log(mth_val) +
                   v[voltable, 7] * log(mth_val) * log(stock_val) +
                   v[voltable, 8] * log(vol_val) * log(stock_val)))
    }
}

CalcVol <- function(mth_val, BA_val, stock_val) {
    if (BA_val <= 0 || mth_val <= 1.6 || stock_val <= 0) {
        return(0)
    } else if (voltable == 1L || voltable == 2L) {
        return(mth_val * BA_val * (v[voltable, 1] * (mth_val - 1.4)^v[voltable, 2] + v[voltable, 3]))
    } else if (voltable == 10L || voltable == 11L) {
        return(BA_val * (v[voltable, 1] + v[voltable, 2] * mth_val + v[voltable, 3] * stock_val))
    } else {
        return(exp(-(v[voltable, 1] + v[voltable, 2] * log(mth_val) +
                     v[voltable, 4] * log(stock_val) +
                     v[voltable, 5] * log(stock_val)^2 +
                     v[voltable, 6] * log(mth_val)^2 +
                     v[voltable, 7] * log(mth_val) * log(stock_val) -
                     log(BA_val)) /
                   (v[voltable, 3] + v[voltable, 8] * log(stock_val))))
    }
}

CalcMTH <- function(si_val, HAge) {
    calcheightcoeff(si_val)
    0.25 + (si_val - 0.25) * ((1 - exp(-ha * HAge)) / (1 - exp(-ha * 20)))^hb
}

Calcagezero <- function() {
    -log(-(1 - exp(-ha * 20)) * ((1.4 - 0.25) / (SI - 0.25))^(1 / hb) + 1) / ha
}

calcMeanht <- function(mth_val, stock_val) {
    A_mh <- 0.07
    B_mh <- -0.00399
    if (is.numeric(mth_val) && is.numeric(stock_val)) {
        return(mth_val * (1 - A_mh * (1 - exp(B_mh * (stock_val - 100)))))
    }
}

MH2MTH <- function(MH_val, stock_val) {
    A_mh <- 0.07
    B_mh <- -0.00399
    if (is.numeric(mth) && is.numeric(stock_val)) {
        return((1 / MH_val * (1 - A_mh * (1 - exp(B_mh * (stock_val - 100)))))^-1)
    }
}

# ---------------------------------------------------------------------------
# CalcDBH: DBH with cubic interpolation between ages 20 and 40 (v1.08 fix)
# ---------------------------------------------------------------------------
CalcDBH <- function(I300_val, SI_val, age_val, stock_val) {
    if (age_val <= 20 || age_val >= 40) {
        A200 <<- CalcA200start(age_val, I300_val, SI_val)
        return(DBHmodel(A200, SI_val, age_val, stock_val))
    } else {
        A200 <<- CalcA200start(19.5, I300_val, SI_val)
        DBH1 <- DBHmodel(A200, SI_val, 19.5, stock_val)
        A200 <<- CalcA200start(20.5, I300_val, SI_val)
        DBH2 <- DBHmodel(A200, SI_val, 20.5, stock_val)
        A200 <<- CalcA200start(39.5, I300_val, SI_val)
        DBH3 <- DBHmodel(A200, SI_val, 39.5, stock_val)
        A200 <<- CalcA200start(40.5, I300_val, SI_val)
        DBH4 <- DBHmodel(A200, SI_val, 40.5, stock_val)
        Y0  <- (DBH1 + DBH2) / 2
        Y1  <- (DBH3 + DBH4) / 2
        Y0p <- (DBH2 - DBH1)
        Y1p <- (DBH4 - DBH3)
        A_c <- Y0
        B_c <- Y0p
        D_c <- (2 * (Y0 + Y0p * 20 - Y1) + 20 * (Y1p - Y0p)) / (20^3)
        C_c <- (Y1p - Y0p - 3 * D_c * 20^2) / (2 * 20)
        return(A_c + B_c * (age_val - 20) + C_c * (age_val - 20)^2 + D_c * (age_val - 20)^3)
    }
}

# ---------------------------------------------------------------------------
# DBHmodel: Predict DBH at given age and stocking
# ---------------------------------------------------------------------------
DBHmodel <- function(A200_val, SI_val, age_val, stock_val) {
    agezero_loc <- Calcagezero()
    site_effect_loc <- A200_val / da1 - 1
    stk <- stock_val
    A_d <- da1 * (1 + site_effect_loc)
    B_d <- db2 * (db1 + dbSI * (SI_val - 28) + dbdia * site_effect_loc +
                  dbsidia * (SI_val - 28) * site_effect_loc)
    if (B_d > -0.05) B_d <- -0.05
    if (age_val < agezero_loc) {
        result <- 0
    } else {
        D200_loc <- OldAgeCorrection(age_val, agezero_loc, B_d) *
            A_d * ((1 - exp(B_d * (age_val - agezero_loc))) /
                   (1 - exp(B_d * (30 - agezero_loc))))^dc
        if (stk > 220) {
            qq <- (log(stk) - log(200))^dr2
        } else {
            qq <- 2 * (log(220) - log(200))^dr2 - (log(242) - log(stk))^dr2
        }
        q_d <- dr * (1 + drsi * (SI_val - 28)) * qq
        P_d <- dl + dm * stk + dn * site_effect_loc
        result <- D200_loc - q_d * log(1 + exp(Ds * (D200_loc - P_d)))
        if (stk > 250) {
            if (dBA_dN(D200_loc, P_d, q_d, stk) <= 0) {
                N_MaxBA_loc <- MaxBAStocking(D200_loc, site_effect_loc, SI_val, stk)
                q_d <- dr * (1 + drsi * (SI_val - 28)) *
                    sign(N_MaxBA_loc - 200) * (abs(log(N_MaxBA_loc) - log(200)))^dr2
                P_d <- dl + dm * N_MaxBA_loc + dn * site_effect_loc
                result <- (D200_loc - q_d * log(1 + exp(Ds * (D200_loc - P_d)))) *
                    sqrt(N_MaxBA_loc / stk)
            }
        }
    }
    if (result < 0) result <- 0
    result
}

# ---------------------------------------------------------------------------
# dBA_dN: Derivative of predicted BA w.r.t. stocking
# ---------------------------------------------------------------------------
dBA_dN <- function(D200_val, P_val, q_val, N_val) {
    dp_dN <- dm
    dq_dN <- q_val * dr2 / N_val / (log(N_val) - log(200))
    dD_dN <- -Ds * D200_val * dq_dN + Ds * P_val * dq_dN + Ds * q_val * dp_dN
    D_loc <- approxDBH(D200_val, P_val, q_val)
    if (D_loc < 0) {
        return(0)
    } else {
        return(D_loc * (D_loc + 2 * N_val * dD_dN))
    }
}

# ---------------------------------------------------------------------------
# MaxBAStocking: Stocking that produces maximum predicted BA
# ---------------------------------------------------------------------------
MaxBAStocking <- function(D200_val, site_effect_val, SI_val, N_val) {
    NA_s <- 250
    q_s  <- dr * (1 + drsi * (SI_val - 28)) * sign(NA_s - 200) *
            (abs(log(NA_s) - log(200)))^dr2
    P_s  <- dl + dm * NA_s + dn * site_effect_val
    FA   <- dBA_dN(D200_val, P_s, q_s, NA_s)
    NB_s <- N_val
    q_s  <- dr * (1 + drsi * (SI_val - 28)) * sign(NB_s - 200) *
            (abs(log(NB_s) - log(200)))^dr2
    P_s  <- dl + dm * NB_s + dn * site_effect_val
    FB   <- dBA_dN(D200_val, P_s, q_s, NB_s)
    NC_s <- NA_s
    for (jj in 1:13) {
        NC_s <- (NA_s + NB_s) / 2
        q_s  <- dr * (1 + drsi * (SI_val - 28)) * sign(NC_s - 200) *
                (abs(log(NC_s) - log(200)))^dr2
        P_s  <- dl + dm * NC_s + dn * site_effect_val
        FC   <- dBA_dN(D200_val, P_s, q_s, NC_s)
        if (FA * FC < 0) {
            NB_s <- NC_s
            FB   <- FC
        } else {
            NA_s <- NC_s
            FA   <- FC
        }
    }
    NC_s
}

# ---------------------------------------------------------------------------
# approxDBH: Approximate DBH (linear term only)
# ---------------------------------------------------------------------------
approxDBH <- function(D200_val, P_val, q_val) {
    D200_val - q_val * Ds * (D200_val - P_val)
}

# ---------------------------------------------------------------------------
# OldAgeCorrection: Old-age growth correction factor
# ---------------------------------------------------------------------------
OldAgeCorrection <- function(age_val, agez_val, B_val) {
    a1_oa <- 1;    a2_oa <- 0.001473784; a3_oa <- 0.973636099
    a4_oa <- 4.350585474; a5_oa <- 25
    T_oa <- (age_val - agez_val) - a5_oa
    if (T_oa < 0) T_oa <- 0
    1 + a4_oa * (1 - exp(-a2_oa * T_oa))^a3_oa
}

# ---------------------------------------------------------------------------
# CalcAge: Calculate age from DBH using bisection
# ---------------------------------------------------------------------------
CalcAge <- function(dbh_val, A200_val, stock_val, SI_val) {
    Bisection(0.001, 150, 15, 3, A200_val, SI_val, stock_val, dbh_val)
}

# ---------------------------------------------------------------------------
# CalcA200: Calculate A200 from DBH, Age, Stocking & SI using bisection
# ---------------------------------------------------------------------------
CalcA200 <- function(dbh_val, age_val, stock_val, SI_val) {
    agezero <<- Calcagezero()
    Bisection(10, 150, 20, 4, SI_val, age_val, stock_val, dbh_val)
}

# ---------------------------------------------------------------------------
# CalcA200start: Calculate A200 from the 300 Index and SI
# ---------------------------------------------------------------------------
CalcA200start <- function(age_val, I300_val, SI_val) {
    adjI300 <- I300_val
    b_adj  <- 0.0206 / 19.488
    c_adj  <- -0.0182 / 100 / 19.488
    k1_adj <- 25
    k2_adj <- 55
    k3_adj <- 215.97
    k4_adj <- -0.05532
    k_adj  <- k3_adj * exp(k4_adj * I300_val)

    if (bias_young && age_val < 6.77) {
        i300adjustment <- 180.5 * adjI300^(-3.256) * (age_val - 6.77)^2
        if (i300adjustment > 5) i300adjustment <- 5
        adjI300 <- adjI300 + i300adjustment
    }

    if (bias_SI) {
        if (SI_val < 25 && SI_val >= 15)
            adjI300 <- adjI300 * (30 - 0.02 * (25 - SI_val) * (age_val - 28.6)) / 30
        if (SI_val < 15)
            adjI300 <- adjI300 * (30 - 0.2 * (age_val - 28.6)) / 30
        if (SI_val > 35 && SI_val <= 45)
            adjI300 <- adjI300 * (30 - 0.02 * (SI_val - 35) * (age_val - 28.6)) / 30
        if (SI_val > 45)
            adjI300 <- adjI300 * (30 - 0.2 * (age_val - 28.6)) / 30
    }

    if (age_val < 30) {
        adjI300 <- adjI300 * (30 + drift * (age_val - 28.6)) / 30
    } else {
        adjI300 <- adjI300
    }

    BA300_30  <- calcBAfromVol(CalcMTH(SI_val, 30), adjI300 * 30, 300)
    DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
    CalcA200(DBH300_30, 28.7, 300, SI_val)
}

# ---------------------------------------------------------------------------
# heightmod: Determine height model (1=NSW, 2=Simple NZ, 3=Environmental NZ)
# ---------------------------------------------------------------------------
heightmod <- function() {
    if (tolower(data_300_indexX[64, 4]) == "x") {
        return(1L)
    } else if (is.na(data_300_indexX[3, 6]) || is.na(data_300_indexX[4, 6]) ||
               abs(as.numeric(data_300_indexX[3, 6])) < 30 ||
               abs(as.numeric(data_300_indexX[3, 6])) > 48) {
        return(2L)
    } else {
        return(3L)
    }
}

# ---------------------------------------------------------------------------
# Bisection: Find root of function fnno using bisection method
# ---------------------------------------------------------------------------
Bisection <- function(xlower, xupper, niterations, fnno, p1, p2, p3, p4) {
    xA <- xlower
    FA <- fn(xA, fnno, p1, p2, p3, p4)
    xB <- xupper
    FB <- fn(xB, fnno, p1, p2, p3, p4)
    xC <- xA
    for (jj in 1:niterations) {
        xC <- (xA + xB) / 2
        FC <- fn(xC, fnno, p1, p2, p3, p4)
        if (FA * FC < 0) {
            xB <- xC
            FB <- FC
        } else {
            xA <- xC
            FA <- FC
        }
    }
    xC
}

# ---------------------------------------------------------------------------
# fn: Functions to be zeroed by the bisection method
# ---------------------------------------------------------------------------
fn <- function(X_val, fnno, p1, p2, p3, p4) {
    if (fnno == 1L) {
        I300 <<- X_val
        Growth(FALSE)
        return(DBH300 - dbh)
    } else if (fnno == 2L) {
        return(p1 - CalcMTH(X_val, p2))
    } else if (fnno == 3L) {
        return(p4 - DBHmodel(p1, p2, X_val, p3))
    } else if (fnno == 4L) {
        return(p4 - DBHmodel(X_val, p1, p2, p3))
    }
}

# ---------------------------------------------------------------------------
# earlyield: Correct early volume predictions when DBH < 2 cm
# ---------------------------------------------------------------------------
earlyield <- function() {
    initialvol <- 0.0000064
    dbh_ey <- 0
    for (i in 5:20) {
        age_ey   <- output_300[i, 7]
        prevdbh  <- dbh_ey
        dbh_ey   <- output_300[i, 16]
        if (dbh_ey >= 2 && prevdbh < 2) {
            treevolinc <- output_300[i, 12] / output_300[i, 8] - initialvol
            k_ey <- treevolinc / (age_ey^2.7)
            for (jj in (i - 1):5) {
                T_ey <- output_300[jj, 7]
                if (T_ey < 0) T_ey <- 0
                output_300[jj, 12] <<- (initialvol + k_ey * T_ey^2.7) * output_300[jj, 8]
                if (output_300[jj, 9] != 0) {
                    output_300[jj, 13] <<- (initialvol + k_ey * T_ey^2.7) * output_300[jj, 9]
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
# mortvol: Estimate volume lost to mortality within each growth increment
# ---------------------------------------------------------------------------
mortvol <- function() {
    numrows <- last_outputline - 4L
    output_300[5, 39] <<- 0
    for (i in 6:(numrows + 4)) {
        sph1 <- output_300[i - 1, 9]
        vol1 <- output_300[i - 1, 13]
        if (sph1 == 0) {
            sph1 <- output_300[i - 1, 8]
            vol1 <- output_300[i - 1, 12]
        }
        sph2 <- output_300[i, 8]
        vol2 <- output_300[i, 12]
        output_300[i, 39] <<- mort_vol(sph1, vol1, sph2, vol2)
    }
}

# ---------------------------------------------------------------------------
# mort_vol: Predict volume lost to mortality within a growth increment
# ---------------------------------------------------------------------------
mort_vol <- function(sph1, vol1, sph2, vol2) {
    (vol1 + vol2) * 0.5 * ((sph1 / sph2)^0.541 - 1)
}

# ---------------------------------------------------------------------------
# density: Predict wood density of each year growth sheath and of the stem
# ---------------------------------------------------------------------------
density <- function() {
    SoilC      <<- as.numeric(data_300_indexX[75, 4])
    SoilN      <<- as.numeric(data_300_indexX[76, 4])
    Temp       <<- as.numeric(data_300_indexX[77, 4])
    CoreDens   <<- as.numeric(data_300_indexX[78, 4])
    CoreAge    <<- as.integer(as.numeric(data_300_indexX[79, 4]))
    InnerRing  <<- as.integer(as.numeric(data_300_indexX[80, 4]))
    OuterRing  <<- as.integer(as.numeric(data_300_indexX[81, 4]))
    GeneticAdj <<- as.numeric(data_300_indexX[82, 4])
    densitymodel <- as.integer(as.numeric(data_300_indexX[83, 4]))

    if (!is.na(SoilC) && SoilC != 0 && !is.na(SoilN) && SoilN != 0 &&
        !is.na(Temp) && Temp != 0) {
        densityinfo <- 1L
    } else if (!is.na(CoreDens) && CoreDens != 0 && !is.na(CoreAge) && CoreAge != 0) {
        densityinfo <- 2L
    } else if (!is.na(CoreDens) && CoreDens != 0 &&
               !is.na(InnerRing) && InnerRing != 0 && !is.na(OuterRing) && OuterRing != 0) {
        densityinfo <- 3L
    } else if (!is.na(CoreDens) && CoreDens != 0) {
        densityinfo <- 4L
    } else {
        densityinfo <- 5L
    }

    numrows <- last_outputline - 4L
    agezero_d <- Calcagezero()

    if (densityinfo == 1L) {
        CoreAge  <<- 26L
        stocking_d <- 250
        CoreDens   <<- outdens26(SoilC, SoilN, Temp, stocking_d, GeneticAdj)
        outdensring <- 18.95 - 0.024 * SI
        Wcal        <- 10.19 + 0.0893 * I300 - 0.255 * SI + 0.00373 * SI * SI - 0.00339 * I300 * SI
    }

    first_ringwidth <- 0
    for (i in 6:(numrows + 4)) {
        if (output_300[i - 1, 17] != 0) {
            prevdbh_d <- output_300[i - 1, 17]
        } else {
            prevdbh_d <- output_300[i - 1, 16]
        }
        if (prevdbh_d != 0) {
            currdbh_d <- output_300[i, 16]
            prevage_d <- output_300[i - 1, 7]
            currage_d <- output_300[i, 7]
            first_ringwidth <- 10 * (currdbh_d - prevdbh_d) / (currage_d - prevage_d) / 2
            break
        }
    }

    for (i in 5:(numrows + 4)) {
        age_d <- output_300[i, 7]
        vol_d <- output_300[i, 12]

        if (i == 5) {
            ringwidth <- first_ringwidth
        } else {
            if (output_300[i - 1, 17] != 0) {
                prevdbh_d <- output_300[i - 1, 17]
            } else {
                prevdbh_d <- output_300[i - 1, 16]
            }
            if (prevdbh_d == 0) {
                ringwidth <- first_ringwidth
            } else {
                currdbh_d <- output_300[i, 16]
                prevage_d <- output_300[i - 1, 7]
                currage_d <- output_300[i, 7]
                ringwidth <- 10 * (currdbh_d - prevdbh_d) / (currage_d - prevage_d) / 2
            }
        }

        if (densityinfo == 4L) {
            output_300[i, 44] <<- CoreDens / 1000
            output_300[i, 40] <<- CoreDens / 1000
        } else {
            if (densitymodel == 2L) {
                ring_d <- age_d - agezero_d
                if (ring_d < 1) ring_d <- 1
                od <- outdens(ring_d, ringwidth, CoreDens, outdensring, Wcal)
            } else {
                od <- old_outdens(age_d, CoreDens, outdensring)
            }
            output_300[i, 44] <<- od / 1000
            output_300[i, 40] <<- sheathdens(od, age_d) / 1000
        }
    }
}

# ---------------------------------------------------------------------------
# outdens: Predict density for a breast height ring from reference density
# ---------------------------------------------------------------------------
outdens <- function(ring_val, ringwidth_val, refdens, refring, refwidth) {
    S_od <- (refdens - 477.8 + 46.2 * log(refwidth) + 84.8 * exp(-0.258 * refring)) /
            (1 - 46.2 * 0.0042 * log(refwidth))
    ringwidth_adj <- ringwidth_val
    if (ringwidth_adj < 1.5) ringwidth_adj <- 1.5
    477.8 + S_od - 46.2 * (1 + 0.0042 * S_od) * log(ringwidth_adj) - 84.8 * exp(-0.258 * ring_val)
}

# ---------------------------------------------------------------------------
# old_outdens: 2007 wood density function
# ---------------------------------------------------------------------------
old_outdens <- function(ring_val, refdens, refring) {
    A_od  <- 332.2
    C_od  <- 0.0193
    g_od  <- 0.0809
    k_od  <- 23.8
    D_od  <- 10.94
    va    <- 0.968
    if (refring < k_od) {
        B_od <- (refdens - A_od * va) / (refring - C_od * refring^2)
    } else {
        B_od <- (refdens - A_od * va) / (D_od + g_od * refring)
    }
    if (ring_val < k_od) {
        A_od * va + B_od * (ring_val - C_od * ring_val^2)
    } else {
        A_od * va + B_od * (D_od + g_od * ring_val)
    }
}

# ---------------------------------------------------------------------------
# sheathdens: Predict density of growth sheath from BH density
# ---------------------------------------------------------------------------
sheathdens <- function(outdens_val, age_val) {
    S_sh <- 1.33415953
    T_sh <- -0.0108173186465
    u_sh <- -0.000963837
    v_sh <- 0.000061770373226
    w_sh <- 0.00002435373794
    age1 <- age_val
    if (age_val > 30) age1 <- 30 + (age_val - 30) * 0.33
    if (age_val > 40) age1 <- 34
    (S_sh + T_sh * age1 + u_sh * outdens_val + v_sh * age1^2 + w_sh * age1 * outdens_val) * outdens_val
}

# ---------------------------------------------------------------------------
# outdens26: Predict outerwood BH density at age 26 from environmental vars
# ---------------------------------------------------------------------------
outdens26 <- function(SoilC_val, SoilN_val, Temp_val, stocking_val, GeneticAdj_val) {
    P_o  <- 143;  q_o  <- 15.9;  R_o  <- 4.1
    A_o  <- 332.2; z_o  <- 18.64; va   <- 0.968
    if (is.na(SoilN_val) || SoilN_val <= 0.014) {
        CN_adj <- 50
    } else {
        CN_adj <- SoilC_val / (SoilN_val - 0.014)
    }
    if (CN_adj > 50) CN_adj <- 50
    outdens_26_250 <- P_o + q_o * Temp_val + R_o * CN_adj
    (A_o * va + (outdens_26_250 * va - A_o * va) *
     (z_o + sqrt(stocking_val)) / (z_o + sqrt(250))) * (1 + GeneticAdj_val / 100)
}

# ---------------------------------------------------------------------------
# CalcOffsets: Derive DBH and MTH offsets for implementation mode 2
# ---------------------------------------------------------------------------
CalcOffsets <- function() {
    Inputparms()
    voltab()
    DBH_calibration_age <<- as.numeric(data_300_indexX[7, 3])
    Stock300_co         <- as.numeric(data_300_indexX[8, 3])
    DBH300_co           <- as.numeric(data_300_indexX[9, 3])
    if (DBH300_co == 0) {
        BA300_co <- as.numeric(data_300_indexX[10, 3])
        if (BA300_co != 0) {
            DBH300_co <- CalcDBHfromBA(BA300_co, Stock300_co)
        } else {
            Vol300_co <- as.numeric(data_300_indexX[11, 3])
            MTH300_co <- CalcMTH(SI, DBH_calibration_age)
            DBH300_co <- CalcDBHfromBA(calcBAfromVol(MTH300_co, Vol300_co, Stock300_co), Stock300_co)
        }
    }
    MTH_calibration_age <<- as.numeric(data_300_indexX[14, 3])
    MTH300_co           <- as.numeric(data_300_indexX[15, 3])

    OUTPUT <<- FALSE
    maxage <<- DBH_calibration_age
    steps  <<- as.integer(maxage / steplength)
    Growth(OUTPUT)
    DBHsqd_add_offset  <<- DBH300_co^2 - dbh^2
    DBHsqd_mult_offset <<- DBH300_co^2 / dbh^2

    maxage <<- MTH_calibration_age
    steps  <<- as.integer(maxage / steplength)
    Growth(OUTPUT)
    MTH_add_offset  <<- MTH300_co - mth
    MTH_mult_offset <<- MTH300_co / mth
}

# ---------------------------------------------------------------------------
# Checkinput stubs (Module3 not yet translated)
# ---------------------------------------------------------------------------
checkinput_site        <- function() TRUE
checkinput_SI          <- function() TRUE
checkinput_htfn        <- function() TRUE
checkinput_initialstock <- function() TRUE
checkinput_stocking    <- function() TRUE
checkinput_prune       <- function() TRUE
checkinput_fellage     <- function() TRUE
checkinput_steplth     <- function() TRUE
checkinput_volfn       <- function() TRUE
checkinput_mortfn      <- function() TRUE
checkinput_I300        <- function() TRUE
checkinput_htage       <- function() TRUE
