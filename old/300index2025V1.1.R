setwd(dirname(getSourceEditorContext()$path))   # set working dir to this script


# ==========================================================
# 300 Index (Radiata, NZ) — Environmental-only SI + I300
# ==========================================================

# ------------------------
# Coefficients & vol table
# ------------------------
coefs <- list(
  # DBH surface
  da1=56.523, db1=-0.09045, dr=2.6416, dl=28.1224, dc=1.4821,
  dbSI=-0.00212, dn=15.7581, dm=-0.00455, dbdia=-0.1325,
  Ds=0.1702, dbsidia=-0.0084, drsi=0.0209, dr2=0.8234,
  db2=1,
  # pruning/thinning (present but not used directly here)
  pra=0.0934, prb=1.98, prc=0.2119,
  thb=0.5, thc=-0.47, tha=0.5,
  thincoeff=0.784,
  # Height: Environmental NZ (ONLY these are used in SI)
  hae0=-1.335, hae1=-0.03581, hae2=-0.0006306,
  hbe0=0.499,  hbe1=0.005059
)

voltab_v <- function() {
  v <- matrix(0, nrow=11, ncol=8)
  v[1,1:3]  <- c(0.942, -1.161, 0.317)
  v[2,1:3]  <- c(0.989, -1.2752, 0.3191)
  v[3,]     <- c(1.492912924, -0.999113309, 1.250753941, -0.397037159, 0.027218164, -0.063166205, 0.064609459, -0.030665365)
  v[4,]     <- c(1.633105986, -1.039327204, 1.212696953, -0.359131176, 0.026454943, -0.067457458, 0.066992488, -0.030528278)
  v[5,]     <- c(0.730448717, -0.617440226, 1.095616037, -0.222220223, 0.013858949, -0.11022445, 0.059157535, -0.016942593)
  v[6,]     <- c(1.09857999,  -0.883862258, 1.165375013, -0.28047221,  0.022081234, -0.059261776, 0.053187392, -0.025226521)
  v[7,]     <- c(1.403009551, -0.96392392,  1.221046594, -0.358337009, 0.024975712, -0.061374804, 0.061895757, -0.028672533)
  v[8,]     <- c(2.834246614, -1.856804825, 1.152097786, -0.201346156,-0.000721117,  0.081503044, 0.024428222,  0.001938887)
  v[9,]     <- c(2.7023,      -2.1301,      1.3901,      -0.5056,      0.0548,       0.0991,      0.1478,     -0.088)
  v[10,1:3] <- c(6.2733,       0.1284,     -0.00097)
  v[11,1:3] <- c(2.1819,       0.2504,     -0.00081)
  v
}

# ------------------------
# Environmental height model (only)
# ------------------------
calcheightcoeff_env <- function(SI, latitude, elevation, pars=coefs) {
  ha <- exp(pars$hae0 + pars$hae1*latitude + pars$hae2*elevation)
  hb <- 1 / (pars$hbe0 + pars$hbe1*SI)     # hb depends on SI (implicit in solver)
  list(ha=ha, hb=hb)
}

CalcMTH_env <- function(SI, Age, latitude, elevation, pars=coefs) {
  hc <- calcheightcoeff_env(SI, latitude, elevation, pars)
  0.25 + (SI - 0.25) * ((1 - exp(-hc$ha * Age)) / (1 - exp(-hc$ha * 20)))^hc$hb
}

# ------------------------
# BA/DBH helpers & volume (needed for I300)
# ------------------------
CalcDBHfromBA <- function(BA, N) sqrt(1.273 * BA / N) * 100
CalcBAfromDBH <- function(DBH, N) N / 1.273 * (DBH / 100)^2

derive_DBH_cm <- function(BA_m2_ha, N_sph) {
  if (is.na(BA_m2_ha) || is.na(N_sph) || BA_m2_ha <= 0 || N_sph <= 0) return(NA_real_)
  100 * sqrt(1.273 * BA_m2_ha / N_sph)
}
calcBAfromVol <- function(MTH, Vol, N, voltable=1, vmat=voltab_v()) {
  if (Vol <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1,2)) return(Vol/(MTH*(vmat[voltable,1]*(MTH-1.4)^vmat[voltable,2] + vmat[voltable,3])))
  if (voltable %in% c(10,11)) return(Vol/(vmat[voltable,1] + vmat[voltable,2]*MTH + vmat[voltable,3]*N))
  exp(vmat[voltable,1] + vmat[voltable,2]*log(MTH) + vmat[voltable,3]*log(Vol) +
        vmat[voltable,4]*log(N) + vmat[voltable,5]*log(N)^2 + vmat[voltable,6]*log(MTH)^2 +
        vmat[voltable,7]*log(MTH)*log(N) + vmat[voltable,8]*log(Vol)*log(N))
}

# ------------------------
# DBH surface (kept intact; needed for I300)
# ------------------------
OldAgeCorrection <- function(Age, agez, B) {
  T <- (Age - agez) - 25
  if (T < 0) T <- 0
  1 + 4.350585474*(1 - exp(-0.001473784*T))^0.973636099
}
Calcagezero_env <- function(SI, latitude, elevation, pars=coefs) {
  hc <- calcheightcoeff_env(SI, latitude, elevation, pars)
  -log(-(1 - exp(-hc$ha*20))*((1.4 - 0.25)/(SI - 0.25))^(1/hc$hb) + 1)/hc$ha
}
approxDBH <- function(D200, P, q, pars=coefs) D200 - q*pars$Ds*(D200 - P)
dBA_dN <- function(D200, P, q, N, pars=coefs) {
  dp_dN <- pars$dm
  dq_dN <- q*pars$dr2/N/(log(N) - log(200))
  dD_dN <- -pars$Ds*D200*dq_dN + pars$Ds*P*dq_dN + pars$Ds*q*dp_dN
  D <- approxDBH(D200, P, q, pars); if (D < 0) return(0)
  D * (D + 2*N*dD_dN)
}
MaxBAStocking <- function(D200, site_effect, SI, Nstart, pars=coefs) {
  # derivative function w.r.t. N
  f <- function(N){
    q <- pars$dr*(1 + pars$drsi*(SI - 28)) * sign(N - 200) * (abs(log(N)-log(200)))^pars$dr2
    P <- pars$dl + pars$dm*N + pars$dn*site_effect
    dBA_dN(D200, P, q, N, pars)
  }
  
  # If derivative at Nstart is >= 0, BA is not decreasing at Nstart — no correction needed
  f_start <- f(Nstart)
  if (!is.finite(f_start) || f_start >= 0) return(NA_real_)
  
  # Find an upper bound with opposite sign (expand if needed)
  NA <- 250
  NB <- max(Nstart, 260)
  fA <- f(NA); fB <- f(NB)
  
  # Expand NB up to a cap if no sign change yet
  cap <- 5000
  while (is.finite(fB) && fA * fB > 0 && NB < cap) {
    NB <- min(cap, NB * 1.5)
    fB <- f(NB)
  }
  
  # If still no sign change, give up -> no correction
  if (!is.finite(fA) || !is.finite(fB) || fA * fB > 0) return(NA_real_)
  
  # Bisection
  for (j in 1:20) {
    NC <- 0.5 * (NA + NB)
    fC <- f(NC)
    if (!is.finite(fC)) break
    if (fA * fC <= 0) { NB <- NC; fB <- fC } else { NA <- NC; fA <- fC }
  }
  0.5 * (NA + NB)
}

DBHmodel_raw <- function(A200, SI, Age, N, latitude, elevation, pars=coefs) {
  agezero <- Calcagezero_env(SI, latitude, elevation, pars)
  site_effect <- A200/pars$da1 - 1
  A <- pars$da1 * (1 + site_effect)
  B <- pars$db2*(pars$db1 + pars$dbSI*(SI-28) + pars$dbdia*site_effect + pars$dbsidia*(SI-28)*site_effect)
  B <- min(B, -0.05)
  if (Age < agezero) return(0)
  
  D200 <- OldAgeCorrection(Age, agezero, B) * A *
    ((1 - exp(B*(Age - agezero))) / (1 - exp(B*(30 - agezero))))^pars$dc
  
  qq <- if (N > 220) (log(N) - log(200))^pars$dr2
  else         2*(log(220)-log(200))^pars$dr2 - (log(242)-log(N))^pars$dr2
  q <- pars$dr*(1 + pars$drsi*(SI - 28)) * qq
  P <- pars$dl + pars$dm*N + pars$dn*site_effect
  D <- D200 - q * log(1 + exp(pars$Ds*(D200 - P)))
  
  ## HARD-OFF: disable high-stocking correction to match VBA numbers
  ## if (N > 250 && dBA_dN(D200, P, q, N, pars) <= 0) {
  ##   Nmax <- MaxBAStocking(D200, site_effect, SI, N, pars)
  ##   q2 <- pars$dr*(1 + pars$drsi*(SI - 28)) * sign(Nmax - 200) * (abs(log(Nmax)-log(200)))^pars$dr2
  ##   P2 <- pars$dl + pars$dm*Nmax + pars$dn*site_effect
  ##   D2 <- D200 - q2 * log(1 + exp(pars$Ds*(D200 - P2)))
  ##   D  <- D2 * sqrt(Nmax/N)
  ## }
  
  max(D, 0)
}


# A200 from I300 (use Environmental MTH at 30)
CalcA200start <- function(Age, I300, SI, latitude, elevation,
                          bias_young=FALSE, bias_SI=FALSE, drift=0,
                          vmat=voltab_v(), pars=coefs) {
  adjI300 <- I300
  if (bias_young && Age < 6.77) {
    i300adj <- 180.5 * adjI300^(-3.256) * (Age - 6.77)^2
    adjI300 <- adjI300 + min(i300adj, 5)
  }
  if (bias_SI) {
    if (SI < 25 && SI >= 15) adjI300 <- adjI300 * (30 - 0.02*(25 - SI)*(Age - 28.6))/30
    if (SI < 15)             adjI300 <- adjI300 * (30 - 0.2*(Age - 28.6))/30
    if (SI > 35 && SI <= 45) adjI300 <- adjI300 * (30 - 0.02*(SI - 35)*(Age - 28.6))/30
    if (SI > 45)             adjI300 <- adjI300 * (30 - 0.2*(Age - 28.6))/30
  }
  if (Age < 30) adjI300 <- adjI300 * (30 + drift*(Age - 28.6))/30
  
  MTH30 <- CalcMTH_env(SI, 30, latitude, elevation, pars)              # ENVIRONMENTAL ONLY
  BA300_30 <- calcBAfromVol(MTH30, adjI300*30, 300, voltable = 2, vmat = vmat)   # use Kimberley 2006 for the 30-year 300-stems calibration
  DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
  
  f <- function(A200) {
    DBHmodel_raw(A200, SI, 28.6, 300,
                 latitude = latitude, elevation = elevation, pars = pars) - DBH300_30 
  }
  a <- 10; b <- 150; fa <- f(a); fb <- f(b)
  for (i in 1:22) { m <- 0.5*(a+b); fm <- f(m); if (fa*fm <= 0) { b <- m; fb <- fm } else { a <- m; fa <- fm } }
  0.5*(a+b)
}

CalcDBH_cubic <- function(I300, SI, Age, N, latitude, elevation,
                          bias_young=FALSE, bias_SI=FALSE, drift=0, pars=coefs) {
  if (Age <= 20 || Age >= 40) {
    A200 <- CalcA200start(Age, I300, SI, latitude, elevation, bias_young, bias_SI, drift, voltab_v(), pars)
    return(DBHmodel_raw(A200, SI, Age, N, latitude, elevation, pars))
  }
  DBH1 <- DBHmodel_raw(CalcA200start(19.5, I300, SI, latitude, elevation, bias_young, bias_SI, drift, voltab_v(), pars),
                       SI, 19.5, N, latitude, elevation, pars)
  DBH2 <- DBHmodel_raw(CalcA200start(20.5, I300, SI, latitude, elevation, bias_young, bias_SI, drift, voltab_v(), pars),
                       SI, 20.5, N, latitude, elevation, pars)
  DBH3 <- DBHmodel_raw(CalcA200start(39.5, I300, SI, latitude, elevation, bias_young, bias_SI, drift, voltab_v(), pars),
                       SI, 39.5, N, latitude, elevation, pars)
  DBH4 <- DBHmodel_raw(CalcA200start(40.5, I300, SI, latitude, elevation, bias_young, bias_SI, drift, voltab_v(), pars),
                       SI, 40.5, N, latitude, elevation, pars)
  Y0 <- (DBH1 + DBH2)/2; Y1 <- (DBH3 + DBH4)/2
  Y0p <- (DBH2 - DBH1);   Y1p <- (DBH4 - DBH3)
  A <- Y0; B <- Y0p
  D <- (2*(Y0 + Y0p*20 - Y1) + 20*(Y1p - Y0p)) / (20^3)
  C <- (Y1p - Y0p - 3*D*20^2) / (2*20)
  A + B*(Age - 20) + C*(Age - 20)^2 + D*(Age - 20)^3
}

# ------------------------
# Root solvers (VBA-style)
# ------------------------
bisection <- function(f, lo, hi, iters=15L) {
  flo <- f(lo); fhi <- f(hi)
  for (i in 1:iters) {
    mid <- 0.5*(lo+hi); fm <- f(mid)
    if (flo*fm <= 0) { hi <- mid; fhi <- fm } else { lo <- mid; flo <- fm }
  }
  0.5*(lo+hi)
}

# SI from MTH (Environmental only)
solve_SI_from_MTH_env <- function(MTH_obs, Age, latitude, elevation) {
  f <- function(SI) MTH_obs - CalcMTH_env(SI, Age, latitude, elevation)
  bisection(f, 5, 60, iters=15L)  # identical bracket/iters as VBA call
}

# I300 from DBH (uses Environmental-only DBH path)

solve_I300_from_DBH <- function(DBH_obs_cm, SI, T, N_at_T, latitude, elevation,
                                lo = 1.328, hi = 60, iters = 30) {
  g <- function(I300) {
    CalcDBH_cubic(I300, SI, Age = T, N = N_at_T,
                  latitude = latitude, elevation = elevation) - DBH_obs_cm
  }
  bisection(g, lo, hi, iters)
}



# ------------------------
# Glue for provided tables
# ------------------------
estimate_SI_I300_for_plot <- function(plot_name, plot_table, meas_row) {
  pr <- subset(plot_table, Plot == plot_name)[1, ]
  if (nrow(pr) == 0) stop(sprintf("Plot '%s' not found", plot_name))
  Age  <- as.numeric(meas_row$Age_years)
  N    <- as.numeric(meas_row$Stocking_sph)
  BA   <- as.numeric(meas_row$BA_m2_ha)
  MTH  <- as.numeric(meas_row$MTH_m)
  
  lat <- as.numeric(pr$Latitude)
  elev<- as.numeric(pr$Elevation)
  
  SI  <- if (isTRUE(all.equal(Age, 20))) MTH else solve_SI_from_MTH_env(MTH, Age, lat, elev)
  
  DBH_obs <- derive_DBH_cm(BA, N)                      # NA if BA is NA
  I300 <- solve_I300_from_DBH(DBH_obs, SI, T=Age, N_at_T=N, latitude=lat, elevation=elev)
  
  data.frame(Plot=plot_name, `300i`=I300, `SI`=SI, check.names=FALSE)
}

estimate_from_tables <- function(plot_table, measurement_table) {
  ms <- subset(measurement_table, Type == "M")
  do.call(rbind, lapply(seq_len(nrow(ms)), function(i){
    estimate_SI_I300_for_plot(ms$Plot[i], plot_table, ms[i, ])
  }))
}

# ==========================
# Test cases (your data)
# ==========================
plot_info <- data.frame(
  Plot = c("Tokoiti","Ashley","Tairua","Rapo"),
  Species = rep("PRAD", 4),
  Year_planted = rep(2018, 4),
  Latitude  = c(46.2, 43.2, 37.1, 39.1),
  Elevation = c(128.0, 224.0, 201.2, 541.0),
  Needle_retention_score = c(2.1, 2.1, 2.7, 2.7),
  Soil_C = c(4.5, 3.0, 6.5, 7.4),
  Soil_N = c(0.2, 0.2, 0.3, 0.6),
  Early_survival_pct = rep(100, 4),
  Mean_Temperature_C = c(10.14, 11.1, 10.4, 10.8),
  check.names = FALSE
)

measurements <- data.frame(
  Plot = c("Ashley","Rapo","Tairua","Tokoiti"),
  Type = rep("M", 4),
  Age_years = c(7, 8, 7, 6),
  Stocking_sph = c(972, 1057, 1209, 924),
  BA_m2_ha = c(20.2, 42.6, 24.8, 10.8),
  MTH_m = c(11.5, 14.9, 11.7, 7.4),
  check.names = FALSE
)

out <- estimate_from_tables(plot_info, measurements)
out$`300i` <- round(out$`300i`, 1)
out$`SI` <- round(out$`SI`, 1)

desired <- c("Tokoiti","Ashley","Tairua","Ashley","Tairua","Rapo")
out <- out[order(match(out$Plot, desired)), ]
print(out, row.names = FALSE)


# Expected output table:
# ----------------------
# Plot      300i    SI
# Tokoiti    27.2   29.4
# Ashley     28.8   35.2
# Tairua     27.2   33.2
# Rapo       35.0   38.6



# BA↔DBH round-trip check (should be ~0)
BA <- 20.2; N <- 972
DBH <- 100 * sqrt(1.273 * BA / N)
BA2 <- N / 1.273 * (DBH/100)^2
abs(BA2 - BA)

