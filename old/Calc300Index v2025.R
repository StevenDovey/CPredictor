#==================================================================================================
# 300 Index Growth Model (Radiata pine, NZ) — Pure R refactor
# - No hidden globals; everything flows via explicit params
# - Pure function for DBH(t | I300, SI, history) to use inside bisection for I300/SI
# - Mirrors VBA v1.09 logic (key guards, cubic splice, drift/bias, mortality v6, vol tables)
#==================================================================================================

# ----------------------------------------------------------------------------
# Model coefficients (from VBA Module 1)
# ----------------------------------------------------------------------------
coefs <- list(
  da1 = 56.523, db1 = -0.09045, dr = 2.6416, dl = 28.1224, dc = 1.4821,
  dbSI = -0.00212, dn = 15.7581, dm = -0.00455, dbdia = -0.1325,
  Ds = 0.1702, dbsidia = -0.0084, drsi = 0.0209, dr2 = 0.8234,
  pra = 0.0934, prb = 1.98, prc = 0.2119,
  ha0 = -2.475, ha1 = -0.01406, hb0 = 0.33417, hb1 = 0.0104,
  hae0 = -1.335, hae1 = -0.03581, hae2 = -0.0006306, hbe0 = 0.499, hbe1 = 0.005059,
  hNSWa = -2.6842, hNSWb = 0.7293, hNSWp = -0.00176,
  thb = 0.5, thc = -0.47, tha = 0.5,
  db2 = 1,
  thincoeff = 0.784,
  # Mortality (incl. 2007 model)
  mortNSW = 0.869,
  morta = 0.000688, mortb = -14.91, mortc = 1.5, morte = -0.0555, mortd = 0.2493,
  mortp = -44.691, mortq = -4.611, morts = 3.901, mortt = 1.3533, mortu = 0.00246,
  mortv = -30.565, mortw = 2.536, mortx = 1.125, morty = 0.000438,
  morta1 = 0.00206, mortb1 = -46.3216, mortc1 = 3.1704, mortd1 = 1.7477,
  morte1 = -0.1631, mortf1 = 0.1991,
  mort2007_a = 0.000459, mort2007_b = 0.974, mort2007_c = 3.06,
  mort2007_d = 0.786, mort2007_f = -0.037, mort2007_g = 0.0371, mort2007_h = -0.32
)

# ----------------------------------------------------------------------------
# Volume table coefficients v[table, k] as per VBA voltab()
# ----------------------------------------------------------------------------
voltab_v <- function() {
  v <- matrix(0, nrow = 11, ncol = 8)
  v[1,1:3]  <- c(0.942, -1.161, 0.317)           # Kimberley & Beets 2007
  v[2,1:3]  <- c(0.989, -1.2752, 0.3191)         # Kimberley 2006
  v[3,]     <- c(1.492912924, -0.999113309, 1.250753941, -0.397037159, 0.027218164, -0.063166205, 0.064609459, -0.030665365)
  v[4,]     <- c(1.633105986, -1.039327204, 1.212696953, -0.359131176, 0.026454943, -0.067457458, 0.066992488, -0.030528278)
  v[5,]     <- c(0.730448717, -0.617440226, 1.095616037, -0.222220223, 0.013858949, -0.11022445, 0.059157535, -0.016942593)
  v[6,]     <- c(1.09857999,  -0.883862258, 1.165375013, -0.28047221,  0.022081234, -0.059261776, 0.053187392, -0.025226521)
  v[7,]     <- c(1.403009551, -0.96392392,  1.221046594, -0.358337009, 0.024975712, -0.061374804, 0.061895757, -0.028672533)
  v[8,]     <- c(2.834246614, -1.856804825, 1.152097786, -0.201346156,-0.000721117,  0.081503044, 0.024428222,  0.001938887)
  v[9,]     <- c(2.7023,      -2.1301,      1.3901,      -0.5056,      0.0548,       0.0991,      0.1478,     -0.088)
  v[10,1:3] <- c(6.2733,       0.1284,     -0.00097)     # NSW1
  v[11,1:3] <- c(2.1819,       0.2504,     -0.00081)     # NSW2
  v
}

# ----------------------------------------------------------------------------
# Small helpers (height model, MTH, agezero, mean height)
# ----------------------------------------------------------------------------
height_coeffs <- function(SI, heightmodel, pars = coefs) {
  if (heightmodel == 1) {
    ha <- exp(pars$hNSWa); hb <- 1 / (pars$hNSWb + pars$hNSWp * SI)
  } else if (heightmodel == 2) {
    ha <- exp(pars$ha0 + pars$ha1 * SI); hb <- 1 / (pars$hb0 + pars$hb1 * SI)
  } else {
    stop("Environmental NZ height model requires latitude & altitude; not included in pure core.")
  }
  list(ha = ha, hb = hb)
}


Calcagezero <- function(SI, heightmodel = 2, latitude = NA, elevation = NA) {
  hc <- calcheightcoeff(SI, heightmodel, latitude = latitude, elevation = elevation)
  -log(-(1 - exp(-hc$ha * 20)) * ((1.4 - 0.25) / (SI - 0.25))^(1 / hc$hb) + 1) / hc$ha
}


calcMeanht <- function(MTH, N) {
  # From VBA: Mean height = MTH * (1 - A*(1 - exp(B*(N - 100))))
  A <- 0.07; B <- -0.00399
  MTH * (1 - A * (1 - exp(B * (N - 100))))
}

# ----------------------------------------------------------------------------
# BA, DBH, Volume functions (stand-level), with same guards as VBA
# ----------------------------------------------------------------------------
CalcDBHfromBA <- function(BA, N) sqrt(1.273 * BA / N) * 100
CalcBAfromDBH <- function(DBH, N) N / 1.273 * (DBH / 100)^2

CalcVol <- function(MTH, BA, N, voltable, vmat) {
  if (BA <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2)) {
    return(MTH * BA * (vmat[voltable,1] * (MTH - 1.4)^vmat[voltable,2] + vmat[voltable,3]))
  }
  if (voltable %in% c(10, 11)) {
    return(BA * (vmat[voltable,1] + vmat[voltable,2] * MTH + vmat[voltable,3] * N))
  }
  # General form
  exp(-(vmat[voltable,1] + vmat[voltable,2]*log(MTH) + vmat[voltable,4]*log(N) +
          vmat[voltable,5]*log(N)^2 + vmat[voltable,6]*log(MTH)^2 +
          vmat[voltable,7]*log(MTH)*log(N) - log(BA)) /
        (vmat[voltable,3] + vmat[voltable,8]*log(N)))
}

calcBAfromVol <- function(MTH, Vol, N, voltable, vmat) {
  if (Vol <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2)) {
    return(Vol / (MTH * (vmat[voltable,1] * (MTH - 1.4)^vmat[voltable,2] + vmat[voltable,3])))
  }
  if (voltable %in% c(10, 11)) {
    return(Vol / (vmat[voltable,1] + vmat[voltable,2] * MTH + vmat[voltable,3] * N))
  }
  exp( vmat[voltable,1] + vmat[voltable,2]*log(MTH) + vmat[voltable,3]*log(Vol) +
         vmat[voltable,4]*log(N) + vmat[voltable,5]*log(N)^2 + vmat[voltable,6]*log(MTH)^2 +
         vmat[voltable,7]*log(MTH)*log(N) + vmat[voltable,8]*log(Vol)*log(N) )
}

# ----------------------------------------------------------------------------
# Old-age correction & BA sensitivity wrt N (needed for high-stocking guard)
# ----------------------------------------------------------------------------
OldAgeCorrection <- function(Age, agez, B) {
  # VBA’s chosen set: a2=0.001473784, a3=0.973636099, a4=4.350585474, a5=25
  T <- (Age - agez) - 25
  if (T < 0) T <- 0
  1 + 4.350585474 * (1 - exp(-0.001473784 * T))^0.973636099
}

approxDBH <- function(D200, P, q, pars = coefs) D200 - q * pars$Ds * (D200 - P)

dBA_dN <- function(D200, P, q, N, pars = coefs) {
  dp_dN <- pars$dm
  dq_dN <- q * pars$dr2 / N / (log(N) - log(200))
  dD_dN <- -pars$Ds*D200*dq_dN + pars$Ds*P*dq_dN + pars$Ds*q*dp_dN
  D <- approxDBH(D200, P, q, pars)
  if (D < 0) return(0)
  D * (D + 2 * N * dD_dN)
}

MaxBAStocking <- function(D200, site_effect, SI, Nstart, pars = coefs) {
  # Bisection in N for root of dBA/dN = 0, between 250 and current N
  NA <- 250
  NB <- max(Nstart, 260)
  f <- function(N) {
    q <- pars$dr * (1 + pars$drsi * (SI - 28)) * sign(N - 200) * (abs(log(N) - log(200))) ^ pars$dr2
    P <- pars$dl + pars$dm * N + pars$dn * site_effect
    dBA_dN(D200, P, q, N, pars)
  }
  FA <- f(NA); FB <- f(NB)
  for (j in 1:13) {
    NC <- (NA + NB)/2
    FC <- f(NC)
    if (FA * FC < 0) { NB <- NC; FB <- FC } else { NA <- NC; FA <- FC }
  }
  (NA + NB)/2
}

# ----------------------------------------------------------------------------
# DBH model (DBH at given Age & N), with cubic splice 20–40 years
# ----------------------------------------------------------------------------
DBHmodel_raw <- function(A200, SI, Age, N,
                         SI_heightmodel = 2, latitude = NA, elevation = NA,
                         pars = coefs) {
  agezero <- Calcagezero(SI, SI_heightmodel, latitude = latitude, elevation = elevation)
  site_effect <- A200 / pars$da1 - 1
  A <- pars$da1 * (1 + site_effect)
  B <- pars$db2 * (pars$db1 + pars$dbSI * (SI - 28) + pars$dbdia * site_effect + pars$dbsidia * (SI - 28) * site_effect)
  B <- min(B, -0.05)
  if (Age < agezero) return(0)
  
  D200 <- OldAgeCorrection(Age, agezero, B) *
    A * ((1 - exp(B * (Age - agezero))) / (1 - exp(B * (30 - agezero)))) ^ pars$dc
  
  if (N > 220) {
    qq <- (log(N) - log(200)) ^ pars$dr2
  } else {
    qq <- 2 * (log(220) - log(200)) ^ pars$dr2 - (log(242) - log(N)) ^ pars$dr2
  }
  q <- pars$dr * (1 + pars$drsi * (SI - 28)) * qq
  P <- pars$dl + pars$dm * N + pars$dn * site_effect
  D <- D200 - q * log(1 + exp(pars$Ds * (D200 - P)))
  
  if (N > 250 && dBA_dN(D200, P, q, N, pars) <= 0) {
    Nmax <- MaxBAStocking(D200, site_effect, SI, N, pars)
    q2 <- pars$dr * (1 + pars$drsi * (SI - 28)) * sign(Nmax - 200) * (abs(log(Nmax) - log(200))) ^ pars$dr2
    P2 <- pars$dl + pars$dm * Nmax + pars$dn * site_effect
    D2 <- D200 - q2 * log(1 + exp(pars$Ds * (D200 - P2)))
    D  <- D2 * sqrt(Nmax / N)
  }
  max(D, 0)
}

CalcA200start <- function(Age, I300, SI,
                          bias_young = FALSE, bias_SI = FALSE, drift = 0,
                          heightmodel = 2, latitude = NA, elevation = NA,
                          vmat = voltab_v(), pars = coefs) {
  adjI300 <- I300
  if (bias_young && Age < 6.77) {
    i300adj <- 180.5 * adjI300^(-3.256) * (Age - 6.77)^2
    if (i300adj > 5) i300adj <- 5
    adjI300 <- adjI300 + i300adj
  }
  if (bias_SI) {
    if (SI < 25 && SI >= 15) adjI300 <- adjI300 * (30 - 0.02 * (25 - SI) * (Age - 28.6)) / 30
    if (SI < 15)             adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
    if (SI > 35 && SI <= 45) adjI300 <- adjI300 * (30 - 0.02 * (SI - 35) * (Age - 28.6)) / 30
    if (SI > 45)             adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
  }
  if (Age < 30) adjI300 <- adjI300 * (30 + drift * (Age - 28.6)) / 30
  
  BA300_30 <- calcBAfromVol(
    CalcMTH(SI, 30,
            latitude  = latitude,
            elevation = elevation,
            force_nsw = (heightmodel == 1)),
    adjI300 * 30, 300, voltable = 1, vmat)
  
  DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
  
  f <- function(A200) DBHmodel_raw(A200, SI, 28.7, 300,
                                   SI_heightmodel = heightmodel,
                                   latitude = latitude, elevation = elevation,
                                   pars = pars) - DBH300_30
  a <- 10; b <- 150; fa <- f(a); fb <- f(b)
  for (i in 1:22) { m <- 0.5*(a+b); fm <- f(m); if (fa*fm <= 0) { b <- m; fb <- fm } else { a <- m; fa <- fm } }
  0.5*(a+b)
}


CalcDBH_cubic <- function(I300, SI, Age, N,
                          heightmodel = 2, latitude = NA, elevation = NA,
                          pars = coefs, bias_young = FALSE, bias_SI = FALSE, drift = 0) {
  if (Age <= 20 || Age >= 40) {
    A200 <- CalcA200start(Age, I300, SI, bias_young, bias_SI, drift,
                          heightmodel, latitude, elevation, voltab_v(), pars)
    return(DBHmodel_raw(A200, SI, Age, N,
                        SI_heightmodel = heightmodel,
                        latitude = latitude, elevation = elevation,
                        pars = pars))
  }
  DBH1 <- DBHmodel_raw(CalcA200start(19.5, I300, SI, bias_young, bias_SI, drift,
                                     heightmodel, latitude, elevation, voltab_v(), pars),
                       SI, 19.5, N, heightmodel, latitude, elevation, pars)
  DBH2 <- DBHmodel_raw(CalcA200start(20.5, I300, SI, bias_young, bias_SI, drift,
                                     heightmodel, latitude, elevation, voltab_v(), pars),
                       SI, 20.5, N, heightmodel, latitude, elevation, pars)
  DBH3 <- DBHmodel_raw(CalcA200start(39.5, I300, SI, bias_young, bias_SI, drift,
                                     heightmodel, latitude, elevation, voltab_v(), pars),
                       SI, 39.5, N, heightmodel, latitude, elevation, pars)
  DBH4 <- DBHmodel_raw(CalcA200start(40.5, I300, SI, bias_young, bias_SI, drift,
                                     heightmodel, latitude, elevation, voltab_v(), pars),
                       SI, 40.5, N, heightmodel, latitude, elevation, pars)
  Y0 <- (DBH1 + DBH2)/2; Y1 <- (DBH3 + DBH4)/2
  Y0p <- (DBH2 - DBH1);   Y1p <- (DBH4 - DBH3)
  A <- Y0; B <- Y0p
  D <- (2 * (Y0 + Y0p*20 - Y1) + 20 * (Y1p - Y0p)) / (20^3)
  C <- (Y1p - Y0p - 3 * D * 20^2) / (2 * 20)
  A + B * (Age - 20) + C * (Age - 20)^2 + D * (Age - 20)^3
}


# ----------------------------------------------------------------------------
# Mortality step (supports NSW=1, NZ=2/3, 2006=4/5, 2007=6). Default: 6
# ----------------------------------------------------------------------------
step_mortality <- function(Nprev, DBHprev, SI, I300, mortmodel = 6,
                           attrition = 0, pctmortadj = 0, dt = 1, pars = coefs) {
  if (mortmodel == 1) {
    # NSW constant rate
    return(Nprev / exp(pars$mortNSW * dt / 100))
  }
  if (DBHprev == 0) return(Nprev)  # no mortality when DBH=0 (as per VBA guards)
  
  if (mortmodel == 2) {
    X <- exp(pars$mortb + pars$morte*SI + pars$mortc*(log(Nprev) + pars$mortd * log(DBHprev^2)))
    mortrate <- (pars$morta + (1 - pars$morta) * X / (1 + X)) * 100
    return(Nprev / exp(mortrate * dt / 100))
  }
  if (mortmodel == 3) {
    X <- exp(pars$mortv + pars$mortw * (log(Nprev) + pars$mortx * log(DBHprev)))
    mortrate <- (pars$morty + (1 - pars$morty) * X / (1 + X)) * 100
    return(Nprev / exp(mortrate * dt / 100))
  }
  if (mortmodel == 4) {
    X <- exp(pars$mortp + pars$mortq * I300 / SI + pars$morts * (log(Nprev) + pars$mortt * log(DBHprev)))
    mortrate <- (attrition + (1 - attrition) * X / (1 + X)) * 100
    return(Nprev / exp(mortrate * dt / 100))
  }
  if (mortmodel == 5) {
    X <- exp(pars$mortb1 + pars$morte1 * I300 + pars$mortf1 * SI + pars$mortc1 * (log(Nprev) + pars$mortd1 * log(DBHprev)))
    mortrate <- (attrition + (1 - attrition) * X / (1 + X)) * 100
    return(Nprev / exp(mortrate * dt / 100))
  }
  # mortmodel == 6 (2007 SDI-based)
  sdi <- exp(pars$mort2007_f * I300 + pars$mort2007_g * SI + log(Nprev) +
               pars$mort2007_d * log(DBHprev / 100) +
               pars$mort2007_h * (log(DBHprev / 100))^2) / 1000
  mortrate <- attrition * 100 + 100 * (1 + pctmortadj / 100) *
    (pars$mort2007_a + pars$mort2007_b * sdi ^ pars$mort2007_c)
  mortrate <- min(max(mortrate, 0), 95)
  Nprev * (1 - mortrate / 100)^dt
}

# ----------------------------------------------------------------------------
# Pure simulator: DBH at target age T, given I300/SI and stand history
# ----------------------------------------------------------------------------
dbh_at_age <- function(
    I300, SI,
    T,                       # target age (years)
    N0,                      # initial stocking at age 0 (sph)
    steplength = 0.1,        # time step (years) — match VBA discretisation
    heightmodel = 2,         # 1=NSW, 2=Simple NZ
    mortmodel = 6,           # default 2007 model
    attrition = 0, pctmortadj = 0,
    bias_young = FALSE, bias_SI = FALSE, drift = 0,
    # optional thinning/pruning schedule (simple: list of events with new N)
    thinnings = data.frame(age = numeric(0), N_after = numeric(0)),
    pars = coefs
) {
  steps <- ceiling(T / steplength)
  N <- N0
  DBH <- 0
  t <- 0
  
  # simple scheduler pointer
  next_thin <- 1
  has_thin <- nrow(thinnings) > 0
  
  for (j in 1:steps) {
    t <- min(T, t + steplength)
    
    # age-shift logic (prune/thin lags) in VBA is element-based; for a pure stand mean DBH
    # we take the stand-level path (keeps maths identical for stand DBH in common use).
    # Compute DBH at current adjusted age = t (no element lags)
    DBH <- CalcDBH_cubic(I300 = I300, SI = SI, Age = t, N = N,
                         heightmodel = heightmodel, pars = pars,
                         bias_young = bias_young, bias_SI = bias_SI, drift = drift)
    
    # apply scheduled thinning when we cross its age (from below)
    if (has_thin && next_thin <= nrow(thinnings)) {
      if (t >= thinnings$age[next_thin] - 1e-6) {
        # emulate waste thin to stated post-thin N (selection/lag effects on DBH are ignored here;
        # if you want the element-level selection & lag, plug in full element engine)
        N <- thinnings$N_after[next_thin]
        next_thin <- next_thin + 1
      }
    }
    
    # mortality to next step based on *current* DBH & N (VBA: uses prev DBH; here identical within small dt)
    if (t < T) {
      N <- step_mortality(N, DBH, SI, I300, mortmodel, attrition, pctmortadj, steplength, pars)
    }
  }
  DBH
}

# Pure, stateless version for the solver
DBH_at <- function(I300, SI, T, N, latitude = NA, elevation = NA, heightmodel = 2) {
  CalcDBH(I300 = I300, SI = SI, Age = T, stockn = N,
          latitude = latitude, elevation = elevation, heightmodel = heightmodel)
}

CalcDBH <- function(I300, SI, Age, stockn, latitude = NA, elevation = NA, heightmodel = 2) {
  CalcDBH_cubic(I300 = I300, SI = SI, Age = Age, N = stockn,
                heightmodel = heightmodel, latitude = latitude, elevation = elevation)
}


# ----------------------------------------------------------------------------
# Bisection solvers that call the pure simulator or MTH height model
# ----------------------------------------------------------------------------
bisection <- function(f, lo, hi, iters = 20) {
  flo <- f(lo); fhi <- f(hi)
  for (i in 1:iters) {
    mid <- 0.5*(lo+hi); fm <- f(mid)
    if (flo * fm <= 0) { hi <- mid; fhi <- fm } else { lo <- mid; flo <- fm }
  }
  0.5*(lo+hi)
}


# Solve SI from a measured MTH at Age (uses existing CalcMTH(SI, HAge))
solve_SI_from_MTH <- function(MTH_obs, Age,
                              latitude = NA, elevation = NA,
                              lo = 5, hi = 60, iters = 20, ...) {
  f <- function(SI) CalcMTH(SI, Age, latitude = latitude, elevation = elevation) - MTH_obs
  bisection(f, lo, hi, iters)
}

# Solve I300 from a measured DBH at Age with known SI and observed stocking N_at_T
solve_I300_from_DBH <- function(DBH_obs_cm, SI, T, N_at_T,
                                latitude = NA, elevation = NA, heightmodel = 2,
                                lo = 1.328, hi = 60, iters = 20, ...) {
  g <- function(I300) DBH_at(I300 = I300, SI = SI, T = T, N = N_at_T,
                             latitude = latitude, elevation = elevation,
                             heightmodel = heightmodel) - DBH_obs_cm
  bisection(g, lo, hi, iters)
}

# Height model selector
# 1 = NSW, 2 = Simple NZ, 3 = Environmental NZ (uses latitude & elevation)
heightmod <- function(latitude = NA, elevation = NA, force_nsw = FALSE) {
  if (isTRUE(force_nsw)) return(1)                                   # explicit NSW override
  if (is.na(latitude) || is.na(elevation) || latitude < 30 || latitude > 48) return(2)  # Simple NZ fallback
  3                                                                  # Environmental NZ
}

# Height model coefficients for the selected model
# Uses your existing coefficient symbols (ha0, ha1, hb0, hb1, hae0, hae1, hae2, hbe0, hbe1, hNSWa, hNSWb, hNSWp)
# Paste over your existing calcheightcoeff()
calcheightcoeff <- function(SI, heightmodel, latitude = NA, elevation = NA, pars = coefs) {
  if (heightmodel == 1) {                       # NSW
    ha <- exp(pars$hNSWa)
    hb <- 1 / (pars$hNSWb + pars$hNSWp * SI)
  } else if (heightmodel == 2) {                # Simple NZ
    ha <- exp(pars$ha0 + pars$ha1 * SI)
    hb <- 1 / (pars$hb0 + pars$hb1 * SI)
  } else {                                      # Environmental NZ
    ha <- exp(pars$hae0 + pars$hae1 * latitude + pars$hae2 * elevation)
    hb <- 1 / (pars$hbe0 + pars$hbe1 * SI)
  }
  list(ha = ha, hb = hb)
}

# Site–height curve (Mean Top Height at age)
# Call as: CalcMTH(SI, HAge, latitude = 46.2, elevation = 128, force_nsw = FALSE)
CalcMTH <- function(SI, HAge, latitude = NA, elevation = NA, force_nsw = FALSE) {
  hm <- heightmod(latitude = latitude, elevation = elevation, force_nsw = force_nsw)
  hc <- calcheightcoeff(SI, hm, latitude = latitude, elevation = elevation)
  0.25 + (SI - 0.25) * ((1 - exp(-hc$ha * HAge)) / (1 - exp(-hc$ha * 20)))^hc$hb
}



derive_DBH_cm <- function(BA_m2_ha, N_sph) {
  if (is.na(BA_m2_ha) || is.na(N_sph) || BA_m2_ha <= 0 || N_sph <= 0) return(NA_real_)
  100 * sqrt(1.273 * BA_m2_ha / N_sph)
}


# ----------------------------------------------------------------------------
# Usage (commented)
# ----------------------------------------------------------------------------.

#VBA model SI= 29.5 I300=27.2

# inputs
Age       <- 7
Stocking  <- 972
BA        <- 20
MTH       <- 12
latitude  <- 46.2
elevation <- 128

DBH_obs_cm <- 100 * sqrt(1.273 * BA / Stocking)

# Solve SI with Environmental NZ (heightmodel=3 via latitude/elevation)
SI_env <- bisection(function(SI) MTH - CalcMTH(SI, Age, latitude = latitude, elevation = elevation), 5, 60, 20)

# Solve SI with NSW override (matches spreadsheet when NSW box is ticked)
SI_nsw <- bisection(function(SI) MTH - CalcMTH(SI, Age, force_nsw = TRUE), 5, 60, 20)


I300_env <- solve_I300_from_DBH(DBH_obs_cm = DBH_obs_cm, SI = SI_env,
                                T = Age, N_at_T = Stocking,
                                latitude = latitude, elevation = elevation,
                                heightmodel = 3)

I300_nsw <- solve_I300_from_DBH(DBH_obs_cm = DBH_obs_cm, SI = SI_nsw,
                                T = Age, N_at_T = Stocking,
                                heightmodel = 1)



# prints
cat(sprintf("VBA model : SI = %4.1f  I300 = %4.1f\n", 29.5, 27.2))
cat(sprintf("R model  (Env NZ) : SI = %4.1f  I300 = %4.1f\n", round(SI_env,1), round(I300_env,1)))
cat(sprintf("R model   (NSW)   : SI = %4.1f  I300 = %4.1f\n", round(SI_nsw,1), round(I300_nsw,1)))
