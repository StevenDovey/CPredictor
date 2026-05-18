#Bisection stop: R stops when bracket ≤ 0.05; VBA uses fixed iterations only (no width stop).
#Height model selection: R forces Environmental NZ; VBA can switch (NSW / Env NZ / Simple NZ) based on sheet flags & bounds.
#Bias toggles: R currently hard-codes bias_young / bias_SI (TRUE in places); VBA reads these from the sheet.
#Voltable choice: R hard-codes voltable = 2 in the A200 calibration; VBA reads the selected model from the sheet (e.g., Kimberley 2006).

# LOAD LIBRARIES ----------------------------------------------------
library(readxl)
library(dplyr)
library(stringr)
library(tidyr)
library(ggplot2)

t_start <- proc.time()[["elapsed"]]

setwd(dirname(rstudioapi::getSourceEditorContext()$path))

if (!exists("read_data", mode = "function")) source("io_utils.R")
InputData <- "inputRp.xlsx"

# ---- Parameters ---------------------------------------------------
parameters <- read_data(InputData, sheet = "parameters")[,1:4]
coefs <- as.list(setNames(as.numeric(parameters$value), parameters$name))

# ==========================================================
age300 <- 30.0                 # C7 context “Age” for I300/SI
DBHcalage <- 28.7    # A200↔DBH anchor age in calibration
HAge <- 20.0                   # C14 alias for MTH calibration age
steplth <- 1.0                 # C48/D48 simulation step (years per step)
fellage <- 50L                 # C47 fell age (years)
voltable <- 1L                 # D51–D61 model selection; 1..11
heightmod_flag <- 2L           # D64/D65 height model; 1=NSW, 2=NZ (env)
mortmodel_flag <- 6L           # D68–D72 mortality; 1=NSW, 2=NZ2004, 3=NZsimple, 4=NZ2006, 5=NZ2007
bias_old <- FALSE              # F51 bias toggle: old-age
bias_young <- TRUE             # F52 bias toggle: young-age
bias_SI <- TRUE                # F53 bias toggle: SI-dependent
driftO <- 0.0                   # F64 drift term for I300 (<30y scaling)
Add <- 0.0                     # F68 additive mortality adjustment
Mult <- 1.0                    # F69 multiplicative mortality adjustment
Thin <- FALSE                  # F70 apply thin-related mortality
implementation <- 1L           # F8 runtime/mode flag 1: Standard mode (default)m 2. Offset modem 3. Index mode"
GeneticAdj <- 0.0              # D82 genetic adjustment factor
densitymodel <- 2L             # D83 density model selector



# ---- Voltab -------------------------------------------------------
voltab <- read_data(InputData, sheet = "VolTab", col_names = TRUE)
#v <- t(as.matrix(voltab))   # numeric 8 x 11 or 11 x 8 depending on how you laid it out
v <- t(data.matrix(voltab))

# ---- Sites --------------------------------------------------------
sites <- read_data(InputData, sheet = "Sites")[1:35]
colnames(sites) <- c(
  "Plot","Species","Year","latitude","elevation","Needl","SoilC","SoilN","SoilP","Surv","MAT","mDens",
  "otrPB","inrng","otrng","300i","SI","300iD","mortA","multm",
  "iDmNL","iDmBL","iDmSL","iDmCRL","iDmFRL",
  "iNNL","iNBL","iNSL","iNCRL","iNFRL",
  "iPNL","iPBL","iPSL","iPCRL","iPFRL"
)

# ---- Plots --------------------------------------------------------
plots <- read_data(InputData, sheet = "Plots")[1:8]
colnames(plots) <- c("Plot","Type","Age","SPH","BA","MTH","SPHp","HTp")

# ---- Filter + Join TEMPORARY------------------------------------------------
plots_M <- subset(plots, Type == "M")
plotsM_with_site <- merge(plots_M, sites, by = "Plot", all.x = TRUE)






# ==========================================================
# 300 Index (Radiata, NZ) — Environmental-only SI + I300
# Uses: coefs (named numeric list) and v (volume matrix) loaded from Excel
# ==========================================================

# ------------------------
# Environmental height model (only)
# ------------------------

# --- Compute environmental height model coefficients ha (rate) and hb (shape) as functions of SI, latitude, and elevation.
calcheightcoeff_env <- function(SI, latitude, elevation, pars = coefs) {
  ha <- exp(pars$hae0 + pars$hae1 * latitude + pars$hae2 * elevation)
  hb <- 1 / (pars$hbe0 + pars$hbe1 * SI)   # hb depends on SI
  list(ha = ha, hb = hb)
}

# --- Mean Top Height (MTH) at a given Age from environmental height model (scaled to SI at age 20).
CalcMTH_env <- function(SI, Age, latitude, elevation, pars = coefs) {
  hc <- calcheightcoeff_env(SI, latitude, elevation, pars)
  0.25 + (SI - 0.25) * ((1 - exp(-hc$ha * Age)) / (1 - exp(-hc$ha * 20)))^hc$hb  # NOTE: ratio term normalizes growth to match SI at Age=20
}

# ------------------------
# BA/DBH  & volume (needed for I300)
# ------------------------

# --- Convert basal area (m2/ha) and stems per hectare to quadratic mean DBH (cm).
CalcDBHfromBA <- function(BA, N) sqrt(1.273 * BA / N) * 100

# --- Convert quadratic mean DBH (cm) and stems per hectare to basal area (m2/ha).
CalcBAfromDBH <- function(DBH, N) N / 1.273 * (DBH / 100)^2

# --- Safe helper to derive DBH (cm) from BA and N with NA/zero guards.
derive_DBH_cm <- function(BA_m2_ha, N_sph) {
  if (is.na(BA_m2_ha) || is.na(N_sph) || BA_m2_ha <= 0 || N_sph <= 0) return(NA_real_)
  100 * sqrt(1.273 * BA_m2_ha / N_sph)
}

# --- Invert volume functions to estimate BA from total volume, MTH, N using selected volume table (v); handles special cases (1/2, 10/11).
calcBAfromVol <- function(MTH, Vol, N, voltable, vmat = v) {
  if (Vol <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2)) {
    # NOTE: Special simple form where BA is proportional to Vol/(MTH * f(MTH)); matches Kimberley/Beets 2007 & Kimberley 2006 parameterizations.
    return(Vol / (MTH * (vmat[voltable, 1] * (MTH - 1.4)^vmat[voltable, 2] + vmat[voltable, 3])))
  }
  if (voltable %in% c(10, 11)) {
    # NOTE: Linear-in-(MTH,N) denominator; used by NSW tables 1 & 2.
    return(Vol / (vmat[voltable, 1] + vmat[voltable, 2] * MTH + vmat[voltable, 3] * N))
  }
  # NOTE: General log-linear form in (MTH, Vol, N) with interaction and quadratic terms; exponent returns BA directly.
  exp(vmat[voltable, 1] + vmat[voltable, 2] * log(MTH) + vmat[voltable, 3] * log(Vol) +
        vmat[voltable, 4] * log(N) + vmat[voltable, 5] * log(N)^2 + vmat[voltable, 6] * log(MTH)^2 +
        vmat[voltable, 7] * log(MTH) * log(N) + vmat[voltable, 8] * log(Vol) * log(N))
}

# ------------------------
# DBH surface (needed for I300)
# ------------------------

# --- Apply old-age correction factor to D200 for ages beyond (agezero + 25).
OldAgeCorrection <- function(Age, agez, B) {
  T <- (Age - agez) - 25
  if (T < 0) T <- 0
  1 + 4.350585474 * (1 - exp(-0.001473784 * T))^0.973636099  # NOTE: asymptotic multiplier >1 that increases slowly with T; exponents tune onset/slope
}

# --- Compute "age zero" (age-equivalent origin) from SI and environment for the environmental height model.
Calcagezero_env <- function(SI, latitude, elevation, pars = coefs) {
  hc <- calcheightcoeff_env(SI, latitude, elevation, pars)
  -log(-(1 - exp(-hc$ha * 20)) * ((1.4 - 0.25) / (SI - 0.25))^(1 / hc$hb) + 1) / hc$ha  # NOTE: inverts MTH curve to find age where MTH≈1.4 m
}

# --- Smooth approximation of DBH transformation used in derivative computations (avoids kink at max()).
approxDBH <- function(D200, P, q, pars = coefs) D200 - q * pars$Ds * (D200 - P)

# --- Derivative d(BA)/dN of basal area w.r.t. stocking, for identifying maximum BA stocking.
dBA_dN <- function(D200, P, q, N, pars = coefs) {
  dp_dN <- pars$dm                                # NOTE: P is linear in N, so derivative is dm
  dq_dN <- q * pars$dr2 / N / (log(N) - log(200)) # NOTE: q depends on (log N - log 200)^{dr2}; undefined at N=200 but code avoids that case
  dD_dN <- -pars$Ds * D200 * dq_dN + pars$Ds * P * dq_dN + pars$Ds * q * dp_dN  # NOTE: chain rule on D = D200 - q*log1p(exp(Ds*(D200-P)))
  D <- approxDBH(D200, P, q, pars); if (D < 0) return(0)
  D * (D + 2 * N * dD_dN)
}

# --- Find N where d(BA)/dN = 0 (maximum BA stocking) via bracket expansion and bisection.
MaxBAStocking <- function(D200, site_effect, SI, Nstart, pars = coefs) {
  # derivative function w.r.t. N
  f <- function(N) {
    q <- pars$dr * (1 + pars$drsi * (SI - 28)) *
      sign(N - 200) * (abs(log(N) - log(200)))^pars$dr2
    P <- pars$dl + pars$dm * N + pars$dn * site_effect
    dBA_dN(D200, P, q, N, pars)
  }
  
  f_start <- f(Nstart)
  if (!is.finite(f_start) || f_start >= 0) return(NA_real_)  # NOTE: need to be on descending limb (negative slope) to search for zero
  
  A <- 250
  B <- max(Nstart, 260)
  fA <- f(A); fB <- f(B)
  
  cap <- 5000
  while (is.finite(fB) && fA * fB > 0 && B < cap) {
    B  <- min(cap, B * 1.5)     # NOTE: geometric bracket expansion until sign change (or cap)
    fB <- f(B)
  }
  if (!is.finite(fA) || !is.finite(fB) || fA * fB > 0) return(NA_real_)
  
  for (j in 1:20) {              # NOTE: fixed-iteration bisection (matches VBA style)
    C  <- 0.5 * (A + B)
    fC <- f(C)
    if (!is.finite(fC)) break
    if (fA * fC <= 0) { B <- C; fB <- fC } else { A <- C; fA <- fC }
  }
  0.5 * (A + B)
}

# --- Core DBH surface — predicts DBH given A200 (site effect), SI, Age, N, environment; includes high-stocking correction.
DBHmodel_raw <- function(A200, SI, Age, N, latitude, elevation, pars = coefs) {
  agezero <- Calcagezero_env(SI, latitude, elevation, pars)
  site_effect <- A200 / pars$da1 - 1
  A <- pars$da1 * (1 + site_effect)
  B <- pars$db2 * (pars$db1 + pars$dbSI * (SI - 28) + pars$dbdia * site_effect + pars$dbsidia * (SI - 28) * site_effect)
  B <- min(B, -0.05)                         # NOTE: guard to keep B negative (ensures monotone saturation)
  if (Age < agezero) return(0)               # NOTE: DBH undefined before model origin; return 0 to mimic VBA behavior
  
  D200 <- OldAgeCorrection(Age, agezero, B) * A *
    ((1 - exp(B * (Age - agezero))) / (1 - exp(B * (30 - agezero))))^pars$dc  # NOTE: D200 at age with old-age multiplier
  
  qq <- if (N > 220) (log(N) - log(200))^pars$dr2 else 2 * (log(220) - log(200))^pars$dr2 - (log(242) - log(N))^pars$dr2
  # NOTE: Above mirrors the qq curve below 220 to avoid discontinuity; matches VBA's piecewise definition
  
  q <- pars$dr * (1 + pars$drsi * (SI - 28)) * qq   # NOTE: crowding effect scales with SI via drsi
  P <- pars$dl + pars$dm * N + pars$dn * site_effect
  D <- D200 - q * log(1 + exp(pars$Ds * (D200 - P)))  # NOTE: smooth hinge via log1p(exp(.)) prevents non-diff kink at D200=P
  
  ## High-stocking correction (VBA Module1)
  if (N > 250 && dBA_dN(D200, P, q, N, pars) <= 0) {   # NOTE: only trigger if BA is at/after maximum (non-increasing)
    Nmax <- MaxBAStocking(D200, site_effect, SI, N, pars)
    q2 <- pars$dr * (1 + pars$drsi * (SI - 28)) *
      sign(Nmax - 200) * (abs(log(Nmax) - log(200)))^pars$dr2
    P2 <- pars$dl + pars$dm * Nmax + pars$dn * site_effect
    D2 <- D200 - q2 * log(1 + exp(pars$Ds * (D200 - P2)))
    D  <- D2 * sqrt(Nmax / N)  # NOTE: scale DBH down by √(Nmax/N) to reflect extra crowding beyond maximum BA
  }
  max(D, 0)
}

# ------------------------
# A200/I300 calibration & cubic age interpolation
# ------------------------

# --- Calibrate A200 from I300 by matching DBH at ~30 years and 300 sph via chosen volume table and bias options.
CalcA200start <- function(Age, I300, SI, latitude, elevation,
                          bias_young = TRUE, bias_SI = TRUE, drift = driftO,
                          vmat = v, pars = coefs) {
  adjI300 <- I300
  if (bias_young && Age < 6.77) {
    i300adj <- 180.5 * adjI300^(-3.256) * (Age - 6.77)^2  # NOTE: quadratic uplift at young ages; capped at +5
    adjI300 <- adjI300 + min(i300adj, 5)
  }
  if (bias_SI) {  # NOTE: age<30 linear bias with SI bands; matches VBA's piecewise scaling relative to 30
    if (SI < 25 && SI >= 15) adjI300 <- adjI300 * (30 - 0.02 * (25 - SI) * (Age - 28.6)) / 30
    if (SI < 15)             adjI300 <- adjI300 * (30 - 0.2  * (Age - 28.6)) / 30
    if (SI > 35 && SI <= 45) adjI300 <- adjI300 * (30 - 0.02 * (SI - 35) * (Age - 28.6)) / 30
    if (SI > 45)             adjI300 <- adjI300 * (30 - 0.2  * (Age - 28.6)) / 30
  }
  if (Age < age300) adjI300 <- adjI300 * (age300 + drift * (Age - 28.6)) / age300  # NOTE: optional drift term around 30-year pivot
  MTH30 <- CalcMTH_env(SI, age300, latitude, elevation, pars)                      # NOTE: normalize at Age=30 for I300 definition
  BA300_30 <- calcBAfromVol(MTH30, adjI300 * age300, 300, voltable, vmat = vmat)  # Kimberley 2006 (table 1) fixed in R
  DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
  
  f <- function(A200) {
    DBHmodel_raw(A200, SI, DBHcalage, 300, latitude = latitude, elevation = elevation, pars = pars) - DBH300_30
  }
  a <- 10; b <- 150; fa <- f(a); fb <- f(b)
  for (i in 1:22) {                      # NOTE: fixed-iteration bisection to mirror VBA behavior
    m <- 0.5 * (a + b); fm <- f(m)
    if (fa * fm <= 0) { b <- m; fb <- fm } else { a <- m; fa <- fm }
  }
  0.5 * (a + b)
}

# --- Smooth DBH across age by cubic interpolation around 20–40 years, falling back to raw DBH surface outside this range.
CalcDBH_cubic <- function(I300, SI, Age, N, latitude, elevation,
                          bias_young = FALSE, bias_SI = FALSE, drift = driftO, pars = coefs) {
  if (Age <= 20 || Age >= 40) {
    A200 <- CalcA200start(Age, I300, SI, latitude, elevation, bias_young = TRUE, bias_SI = TRUE, drift, v, pars)
    return(DBHmodel_raw(A200, SI, Age, N, latitude, elevation, pars))
  }
  # NOTE: Four-point stencil (19.5,20.5,39.5,40.5) to estimate values and slopes at 20 and 40 for cubic interpolation.
  DBH1 <- DBHmodel_raw(CalcA200start(19.5, I300, SI, latitude, elevation, bias_young = TRUE, bias_SI = TRUE, drift, v, pars),
                       SI, 19.5, N, latitude, elevation, pars)
  DBH2 <- DBHmodel_raw(CalcA200start(20.5, I300, SI, latitude, elevation, bias_young = TRUE, bias_SI = TRUE, drift, v, pars),
                       SI, 20.5, N, latitude, elevation, pars)
  DBH3 <- DBHmodel_raw(CalcA200start(39.5, I300, SI, latitude, elevation, bias_young = TRUE, bias_SI = TRUE, drift, v, pars),
                       SI, 39.5, N, latitude, elevation, pars)
  DBH4 <- DBHmodel_raw(CalcA200start(40.5, I300, SI, latitude, elevation, bias_young = TRUE, bias_SI = TRUE, drift, v, pars),
                       SI, 40.5, N, latitude, elevation, pars)
  Y0 <- (DBH1 + DBH2) / 2; Y1 <- (DBH3 + DBH4) / 2
  Y0p <- (DBH2 - DBH1);     Y1p <- (DBH4 - DBH3)
  A <- Y0; B <- Y0p
  D <- (2 * (Y0 + Y0p * 20 - Y1) + 20 * (Y1p - Y0p)) / (20^3)  # NOTE: solve cubic coefficients to match end values & slopes
  C <- (Y1p - Y0p - 3 * D * 20^2) / (2 * 20)
  A + B * (Age - 20) + C * (Age - 20)^2 + D * (Age - 20)^3
}

# ------------------------
# Root solvers (VBA-style)
# ------------------------

# --- Generic bisection root-finder; iterates up to iters, stops early if bracket ≤ 0.05 and returns to nearest 0.05.
bisectionOLD <- function(f, lo, hi, iters = 30L) {
  flo <- f(lo); fhi <- f(hi)
  step <- 0.05                          # NOTE: discretizes solution to 0.05 grid; also acts as early-stop threshold on bracket width
  for (i in 1:iters) {
    mid <- 0.5 * (lo + hi); fm <- f(mid)
    if (flo * fm <= 0) {
      hi <- mid; fhi <- fm
    } else {
      lo <- mid; flo <- fm
    }
    if ((hi - lo) <= step) break        # NOTE: early exit when interval is sufficiently tight
  }
  cand <- if (abs(flo) <= abs(fhi)) lo else hi
  round(cand / step) * step             # NOTE: snap to nearest 0.05 to mimic spreadsheet/VBA reporting
}

# Brent/golden-section replacement with same signature/name.
# Minimizes squared residual over [lo, hi] and snaps to 0.05 grid.
bisection <- function(f, lo, hi, iters = 30L) {
  obj <- function(x) {
    fx <- f(x)
    if (is.na(fx) || !is.finite(fx)) return(Inf)
    fx * fx
  }
  res <- optimize(obj, interval = c(lo, hi))
  step <- 0.05
  round(res$minimum / step) * step
}



# ------------------------
# SI and I300 solvers (environmental-only path)
# ------------------------

# --- Solve Site Index (SI) that reproduces observed MTH at Age for given latitude/elevation via bisection.
solve_SI_from_MTH_env <- function(MTH_obs, Age, latitude, elevation) {
  f <- function(SI) MTH_obs - CalcMTH_env(SI, Age, latitude, elevation)
  bisection(f, 5, 60, iters = 15L)      # NOTE: SI bracket [5,60] reflects VBA input validation bounds
}

# --- Solve 300 Index (I300) that reproduces observed DBH at age T and stocking N via cubic DBH and bisection.
solve_I300_from_DBH <- function(DBH_obs_cm, SI, T, N_at_T, latitude, elevation,
                                lo = 1.328, hi = 60, iters = 30) {
  g <- function(I300) {
    CalcDBH_cubic(I300, SI, Age = T, N = N_at_T, latitude = latitude, elevation = elevation) - DBH_obs_cm
  }
  bisection(g, lo, hi, iters)           # NOTE: I300 lower bound 1.328 mirrors VBA’s valid range
}

# ------------------------
# Wrapper over joined data
# ------------------------

# --- Row-wise driver over merged plot/site data to estimate SI and I300 (environmental-only path).
estimate_from_joined <- function(plotsM_with_site_df) {
  do.call(rbind, lapply(seq_len(nrow(plotsM_with_site_df)), function(i) {
    r   <- plotsM_with_site_df[i, ]
    Age <- as.numeric(r$Age)
    N   <- as.numeric(r$SPH)
    BA  <- as.numeric(r$BA)
    MTH <- as.numeric(r$MTH)
    latitude <- as.numeric(r$latitude)
    elevation <- as.numeric(r$elevation)
    
    SI   <- if (isTRUE(all.equal(Age, 20))) MTH else solve_SI_from_MTH_env(MTH, Age, latitude, elevation)  # NOTE: at Age=20, SI ≈ MTH by definition
    DBH  <- derive_DBH_cm(BA, N)
    I300 <- solve_I300_from_DBH(DBH, SI, T = Age, N_at_T = N, latitude = latitude, elevation = elevation)
    
    data.frame(Plot = r$Plot, Age=Age, `300i` = I300, `SI` = SI, check.names = FALSE)
  }))
}

# run
out <- estimate_from_joined(plotsM_with_site)
write.csv(out, "out300SI.csv",na = "")




outM <- merge(out,plotsM_with_site, by = "Plot", all.x = TRUE)
outM[,25:26]<- round(outM[,25:26],1)          # NOTE: rounds the original '300i' and 'SI' columns by position (defensive: names could shift)
outM$C<-outM$'300i.y'- outM$'300i.x'          # NOTE: delta = VBA-derived minus R-derived (or vice versa depending on merge order)
outM <-arrange(outM, Plot) 
 print(outM[, c(1:3,5:8,25:26,45)], row.names = FALSE)

write.csv(outM[, c(1:3,5:8,25:26,45)], "out.csv",na = "")  # NOTE: write selected columns only; empty strings for NA to mirror spreadsheet behavior

sites$`300i`[match(out$Plot, sites$Plot)] <- out$`300i`
sites$SI     [match(out$Plot, sites$Plot)] <- out$SI


 
 



.morta  <- 0.000688; .mortb <- -14.91; .mortc <- 1.5; .mortd <- 0.2493; .morte <- -0.0555
.m2007_a <- 0.000557; .m2007_b <- 0.009251; .m2007_c <- 1.3409
.m2007_d <- 0.78227;  .m2007_f <- -30.565;  .m2007_g <- 2.536; .m2007_h <- 1.125


# ---- Growth Model Output  --------------------------------------------

# Volume (forward)
CalcVol <- function(MTH, BA, N, voltable = voltable, vmat = v) {
  if (BA <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2))  return(MTH * BA * (vmat[voltable,1] * (MTH - 1.4)^vmat[voltable,2] + vmat[voltable,3]))
  if (voltable %in% c(10, 11)) return(BA * (vmat[voltable,1] + vmat[voltable,2]*MTH + vmat[voltable,3]*N))
  exp(vmat[voltable,1] + vmat[voltable,2]*log(MTH) + vmat[voltable,3]*log(BA) +
        vmat[voltable,4]*log(N) + vmat[voltable,5]*log(N)^2 + vmat[voltable,6]*log(MTH)^2 +
        vmat[voltable,7]*log(MTH)*log(N) + vmat[voltable,8]*log(BA)*log(N))
}

# Mean height from MTH & stocking
calcMeanht <- function(MTH, stock) {
  A <- 0.07; B <- -0.00399
  if (is.na(MTH) || is.na(stock)) return(NA_real_)
  MTH * (1 - A * (1 - exp(B * (stock - 100))))
}

# ---- Density (VBA-aligned) --------------------------------------------------
outdens <- function(ring, ringwidth, refdens, refring, refwidth) {
  S <- (refdens - 477.8 + 46.2 * log(refwidth) + 84.8 * exp(-0.258 * refring)) /
    (1 - 46.2 * 0.0042 * log(refwidth))
  rw <- ringwidth; if (!is.finite(rw) || rw < 1.5) rw <- 1.5
  477.8 + S - 46.2 * (1 + 0.0042*S) * log(rw) - 84.8 * exp(-0.258 * ring)
}
old_outdens <- function(ring, refdens, refring) {
  A <- 332.2; C <- 0.0193; g <- 0.0809; k <- 23.8; D <- 10.94; adj <- 0.968
  B <- if (refring < k) (refdens - A*adj)/(refring - C*refring^2) else (refdens - A*adj)/(D + g*refring)
  if (ring < k) A*adj + B*(ring - C*ring^2) else A*adj + B*(D + g*ring)
}
sheathdens <- function(outdens_val, Age) {
  S <- 1.33415953; T <- -0.0108173186465; u <- -0.000963837
  V <- 0.000061770373226; w <- 0.00002435373794
  age1 <- Age; if (Age > 30) age1 <- 30 + (Age - 30)*0.33; if (Age > 40) age1 <- 34
  (S + T*age1 + u*outdens_val + V*age1^2 + w*age1*outdens_val) * outdens_val
}
outdens26 <- function(SoilC, SoilN, Temp, stocking, GeneticAdj) {
  p <- 143; q <- 15.9; R <- 4.1; A <- 332.2; z <- 18.64; adj <- 0.968
  CN <- if (SoilN <= 0.014) 50 else SoilC/(SoilN - 0.014); if (CN > 50) CN <- 50
  ow_26_250 <- p + q*Temp + R*CN
  (A*adj + (ow_26_250*adj - A*adj) * (z + sqrt(stocking))/(z + sqrt(250))) * (1 + GeneticAdj/100)
}
calcVol_from_BA <- function(MTH, BA, N, voltable, vmat = v) {
  if (!is.finite(BA) || BA <= 0 || !is.finite(MTH) || MTH <= 1.6 || !is.finite(N) || N <= 0) return(NA_real_)
  if (voltable %in% c(1, 2))  return(BA * MTH * (vmat[voltable,1] * (MTH - 1.4)^vmat[voltable,2] + vmat[voltable,3]))
  if (voltable %in% c(10, 11)) return(BA * (vmat[voltable,1] + vmat[voltable,2]*MTH + vmat[voltable,3]*N))
  NA_real_
}
compute_density_profile <- function(r, ages, SI, I300,
                                    densitymodel = densitymodel,
                                    voltable = voltable) {
  
  SoilC <- suppressWarnings(as.numeric(r$SoilC))
  SoilN <- suppressWarnings(as.numeric(r$SoilN))
  Temp  <- suppressWarnings(as.numeric(r$MAT))
  CoreDens   <- suppressWarnings(as.numeric(get0("CoreDens",  ifnotfound = r$mDens)))
  CoreAge    <- suppressWarnings(as.numeric(get0("CoreAge",   ifnotfound = NA_real_)))
  InnerRing  <- suppressWarnings(as.numeric(r$inrng))
  OuterRing  <- suppressWarnings(as.numeric(r$otrng))
  GeneticAdj <- suppressWarnings(as.numeric(get0("GeneticAdj", ifnotfound = 0)))
  lat  <- suppressWarnings(as.numeric(r$latitude))
  elev <- suppressWarnings(as.numeric(r$elevation))
  
  # Route selection: mimic VBA (non-zero rather than non-NA)
  nz <- function(x) is.finite(x) && x != 0
  densityinfo <-
    if (nz(SoilC) && nz(SoilN) && nz(Temp)) 1
  else if (nz(CoreDens) && nz(CoreAge)) 2
  else if (nz(CoreDens) && nz(InnerRing) && nz(OuterRing)) 3
  else if (nz(CoreDens)) 4 else 5
  
  agezero <- Calcagezero_env(SI, lat, elev)
  
  if (densityinfo == 1) {
    CoreAge    <- 26
    CoreDens   <- outdens26(SoilC, SoilN, Temp, stocking = 250, GeneticAdj)
    outdensring <- 18.95 - 0.024 * SI
    Wcal        <- 10.19 + 0.0893 * I300 - 0.255 * SI + 0.00373 * SI^2 - 0.00339 * I300 * SI
  } else {
    outdensring <- max(1, mean(range(ages)) - agezero)
    Wcal        <- 3
  }
  
  # Stocking series (VBA-style) and curves using age-varying N
  N_series <- stock_series(
    ages, N0 = as.numeric(r$SPH), SI = SI, I300 = I300,
    latitude = lat, elevation = elev,
    mortmodel_flag = mortmodel_flag, attrition = Add, pctmortadj = Mult
  )
  
  dbh <- vapply(seq_along(ages), function(i)
    CalcDBH_cubic(I300, SI, Age = ages[i], N = N_series[i],
                  latitude = lat, elevation = elev),
    numeric(1)
  )
  mth <- vapply(ages, function(a) CalcMTH_env(SI, a, lat, elev), numeric(1))
  ba  <- CalcBAfromDBH(dbh, N_series)
  vol <- mapply(function(M, HBA, n) calcVol_from_BA(M, HBA, n, voltable), mth, ba, N_series)
  
  # First ring width (mm), fallback to 1.5 if we can’t infer
  first_ringwidth <- NA_real_
  for (i in seq_along(ages)) {
    if (i > 1 && is.finite(dbh[i-1]) && dbh[i-1] > 0) {
      da <- ages[i] - ages[i-1]
      if (da > 0) first_ringwidth <- 10 * (dbh[i] - dbh[i-1]) / da / 2
      break
    }
  }
  if (!is.finite(first_ringwidth)) first_ringwidth <- 1.5
  
  # Per-year ring width
  ringwidth <- numeric(length(ages))
  for (i in seq_along(ages)) {
    if (i == 1 || !is.finite(dbh[i-1]) || dbh[i-1] == 0) {
      ringwidth[i] <- first_ringwidth
    } else {
      ringwidth[i] <- 10 * (dbh[i] - dbh[i-1]) / (ages[i] - ages[i-1]) / 2
    }
  }
  
  # BH ring density (kg/m3) and growth-sheath density (kg/m3)
  od_bh <- gs <- rep(NA_real_, length(ages))
  for (i in seq_along(ages)) {
    ring <- max(1, ages[i] - agezero)
    od <- if (densityinfo == 4) {
      CoreDens
    } else if (isTRUE(as.integer(densitymodel) == 2L)) {
      outdens(ring, ringwidth[i], CoreDens, outdensring, Wcal)
    } else {
      old_outdens(ring, CoreDens, outdensring)
    }
    od_bh[i] <- od
    gs[i]    <- sheathdens(od, ages[i])
  }
  
  # Whole-stem density (g/cm3), volume-weighted
  stem <- rep(NA_real_, length(ages))
  for (i in seq_along(ages)) {
    if (!is.finite(vol[i])) next
    if (i == 1 || !is.finite(vol[i-1])) {
      stem[i] <- gs[i] / 1000
    } else {
      stem[i] <- ((gs[i] / 1000) * (vol[i] - vol[i-1]) + stem[i-1] * vol[i-1]) / vol[i]
    }
  }
  
  data.frame(
    Age = ages,
    DBH_cm = dbh,
    BA_m2_ha = ba,
    Vol_m3_ha = vol,
    BH_outdens_gcc = od_bh / 1000,
    GS_dens_gcc = gs / 1000,
    Stem_dens_gcc = stem,
    MTH_m = mth,
    check.names = FALSE
  )
}

# ---- Mortality  ---------------------------------------------------
################### 1=NSW, 2=NZ2004, 3=NZsimple, 4=NZ2006, 5=NZ2007
## --- Mortality models 1..6  --------------------
mort_rate_nsw <- function() {
  if (!is.null(coefs$mortNSW)) as.numeric(coefs$mortNSW) else stop("coefs$mortNSW missing")
}
mort_rate_nz2004 <- function(SI, N, DBH_cm) {
  if (!is.finite(DBH_cm) || DBH_cm <= 0) return(0)
  with(coefs, {
    X  <- exp(mortb + morte*SI + mortc*(log(N) + mortd*log(DBH_cm^2)))
    pmin(pmax((morta + (1 - morta) * X/(1 + X)) * 100, 0), 95)
  })
}
mort_rate_nzsimple <- function(N, DBH_cm) {  # VBA model 3
  if (!is.finite(DBH_cm) || DBH_cm <= 0) return(0)
  with(coefs, {
    X  <- exp(mortv + mortw*(log(N) + mortx*log(DBH_cm)))
    pmin(pmax((morty + (1 - morty) * X/(1 + X)) * 100, 0), 95)
  })
}
mort_rate_nz2006_a <- function(I300, SI, N, DBH_cm, attrition) {  # VBA model 4
  if (!is.finite(DBH_cm) || DBH_cm <= 0) return(0)
  with(coefs, {
    X  <- exp(mortp + mortq*(I300/SI) + morts*(log(N) + mortt*log(DBH_cm)))
    pmin(pmax((attrition + (1 - attrition) * X/(1 + X)) * 100, 0), 95)
  })
}
mort_rate_nz2006_b <- function(I300, SI, N, DBH_cm, attrition) {  # VBA model 5
  if (!is.finite(DBH_cm) || DBH_cm <= 0) return(0)
  with(coefs, {
    X  <- exp(mortb1 + morte1*I300 + mortf1*SI + mortc1*(log(N) + mortd1*log(DBH_cm)))
    pmin(pmax((attrition + (1 - attrition) * X/(1 + X)) * 100, 0), 95)
  })
}
mort_rate_nz2007 <- function(I300, SI, N, DBH_cm, attrition = 0, pctmortadj = 0) {  # VBA model 6
  if (!is.finite(DBH_cm) || DBH_cm <= 0) return(0)
  with(coefs, {
    dlog <- log(pmax(DBH_cm, 1e-6)/100)  # DBH in m
    sdi  <- exp(mort2007_f*I300 + mort2007_g*SI + log(N) + mort2007_d*dlog + mort2007_h*(dlog^2)) / 1000
    mr   <- attrition*100 + 100*(1 + pctmortadj/100)*(mort2007_a + mort2007_b * sdi^mort2007_c)
    pmin(pmax(mr, 0), 95)
  })
}

## --- Annual stocking series (uses VBA numbering 1..6) ---------------
stock_series <- function(ages, N0, SI, I300, latitude, elevation,
                         mortmodel_flag = mortmodel_flag,
                         attrition = Add, pctmortadj = Mult) {
  if (!length(ages)) return(numeric(0))
  N <- numeric(length(ages)); N[1] <- N0
  for (i in 2:length(ages)) {
    a_prev <- ages[i-1]; dt <- ages[i] - a_prev
    DBH_prev <- CalcDBH_cubic(I300, SI, Age = a_prev, N = N[i-1],
                              latitude = latitude, elevation = elevation)
    
    mr <- switch(as.integer(mortmodel_flag),
                 mort_rate_nsw(),                                                # 1
                 mort_rate_nz2004(SI, N[i-1], DBH_prev),                         # 2
                 mort_rate_nzsimple(N[i-1], DBH_prev),                           # 3
                 mort_rate_nz2006_a(I300, SI, N[i-1], DBH_prev, attrition),      # 4
                 mort_rate_nz2006_b(I300, SI, N[i-1], DBH_prev, attrition),      # 5
                 mort_rate_nz2007(I300, SI, N[i-1], DBH_prev, attrition, pctmortadj)  # 6
    )
    
    # VBA uses: N_next = N_prev / exp(mortrate * step / 100) except model 6 uses (1 - mr/100)^step.
    if (mortmodel_flag == 6L) {
      N[i] <- N[i-1] * (1 - mr/100)^dt
    } else {
      N[i] <- N[i-1] / exp(mr * dt / 100)
    }
  }
  N
}

#######################


# ---- Annual per-plot series (final) ------------------------------------------

build_annual_series_for_plot <- function(r, si_i300_row) {
  plot_id  <- as.character(r$Plot)
  SI_est   <- as.numeric(si_i300_row$SI)
  I300_est <- as.numeric(si_i300_row$`300i`)
  lat      <- as.numeric(r$latitude)
  elev     <- as.numeric(r$elevation)
  N_init   <- as.numeric(r$SPH)
  
  ages <- seq(0, fellage, by = 1)
  N_at_age <- stock_series(ages, N0 = N_init, SI = SI_est, I300 = I300_est,
                           latitude = lat, elevation = elev,
                           mortmodel_flag = mortmodel_flag, attrition = Add, pctmortadj = Mult)
  
  MTH <- vapply(ages, function(a) CalcMTH_env(SI_est, a, lat, elev), numeric(1))
  DBH <- mapply(function(a, n) CalcDBH_cubic(I300_est, SI_est, Age = a, N = n, latitude = lat, elevation = elev),
                ages, N_at_age)
  BA  <- N_at_age / 1.273 * (pmax(DBH, 0) / 100)^2
  StemVol <- mapply(function(h, b, n) CalcVol(h, b, n, voltable = voltable, vmat = v), MTH, BA, N_at_age)
  MeanHt  <- mapply(calcMeanht, MTH, N_at_age)
  
  dens_tbl <- compute_density_profile(r = r, ages = ages, SI = SI_est, I300 = I300_est,
                                      densitymodel = densitymodel, voltable = voltable)
  
  data.frame(
    Plot = plot_id,
    `Age (years)` = ages,
    `Stocking (stem/ha)` = N_at_age,
    `BA (m2/ha)` = BA,
    `DBH (cm)` = DBH,
    `Stem Vol. (m3/ha)` = StemVol,
    `Harvest vol. (m3/ha)` = replace(numeric(length(ages)), length(ages), tail(StemVol, 1)),
    `MTH (m)` = MTH,
    `Mean height (m)` = MeanHt,
    `Wood dens. (kg/m3)` = dens_tbl$Stem_dens_gcc * 1000,
    `Crown height (m)` = suppressWarnings(as.numeric(r$HTp)),
    check.names = FALSE
  )
}






si_i300_lookup <- out[, c("Plot", "SI", "300i")]
annual_list <- lapply(seq_len(nrow(plotsM_with_site)), function(i) {
  r <- plotsM_with_site[i, ]
  row <- si_i300_lookup[si_i300_lookup$Plot == as.character(r$Plot), ]
  if (!nrow(row)) return(NULL)
  build_annual_series_for_plot(r, row[1, ])
})
annual_all <- bind_rows(Filter(Negate(is.null), annual_list))
write.csv(annual_all, "annual_series1.csv", row.names = FALSE)

# ---- Build + write for all plots --------------------------------------------
sirow <- si_i300_lookup[si_i300_lookup$Plot == plotsM_with_site$Plot[1], ]
annual_first <- build_annual_series_for_plot(plotsM_with_site[1, ], sirow)
write.csv(annual_first, "annual_series1.csv", row.names = FALSE)







################## GRAPHICS
{  # identical logic to  bisection, but records each midpoint guess
 bisection_with_traceOLD <- function(f, lo, hi, iters = 30L, step = 0.05) {
   flo <- f(lo); fhi <- f(hi)
   rec <- data.frame(iter = integer(0), lo = numeric(0), hi = numeric(0),
                     mid = numeric(0), flo = numeric(0), fhi = numeric(0), fmid = numeric(0))
   for (i in 1:iters) {
     mid <- 0.5 * (lo + hi); fm <- f(mid)
     rec <- rbind(rec, data.frame(iter = i, lo = lo, hi = hi, mid = mid,
                                  flo = flo, fhi = fhi, fmid = fm))
     if (flo * fm <= 0) { hi <- mid; fhi <- fm } else { lo <- mid; flo <- fm }
     if ((hi - lo) <= step) break
   }
   cand <- if (abs(flo) <= abs(fhi)) lo else hi
   root <- round(cand / step) * step
   list(root = root, trace = rec)
 }
 
 
# Brent/golden-section version with the SAME return shape so  graphics still work.
# It minimizes squared residual over [lo, hi], records each tested point as "mid",
# and snaps the final root to the 0.05 grid (like  bisection).
bisection_with_trace <- function(f, lo, hi, iters = 30L, step = 0.05) {
  obj <- function(x) {
    fx <- f(x)
    if (is.na(fx) || !is.finite(fx)) return(Inf)
    fx * fx
  }
  a <- lo; b <- hi
  phi <- (sqrt(5) - 1) / 2  # golden ratio conjugate
  c <- b - phi * (b - a)
  d <- a + phi * (b - a)
  fc <- obj(c)
  fd <- obj(d)
  
  rec <- data.frame(iter = integer(0), lo = numeric(0), hi = numeric(0),
                    mid = numeric(0), flo = numeric(0), fhi = numeric(0), fmid = numeric(0))
  
  for (i in 1:iters) {
    # record the newly evaluated point as "mid" (choose the better of c/d this step)
    if (fc < fd) {
      # keep [a, d]
      # record c
      rec <- rbind(rec, data.frame(
        iter = i,
        lo   = a,
        hi   = d,
        mid  = c,
        flo  = f(a),
        fhi  = f(d),
        fmid = f(c)
      ))
      b  <- d
      d  <- c; fd <- fc
      c  <- b - phi * (b - a); fc <- obj(c)
    } else {
      # keep [c, b]
      # record d
      rec <- rbind(rec, data.frame(
        iter = i,
        lo   = c,
        hi   = b,
        mid  = d,
        flo  = f(c),
        fhi  = f(b),
        fmid = f(d)
      ))
      a  <- c
      c  <- d; fc <- fd
      d  <- a + phi * (b - a); fd <- obj(d)
    }
    if ((b - a) <= step) break
  }
  
  cand <- (a + b) / 2
  root <- round(cand / step) * step
  list(root = root, trace = rec)
}



 show_si_resolution <- function(Age_star, MTH_obs, latitude, elevation,
                                lo = 5, hi = 60, iters = 15L, step = 0.05) {
   # residual: f(SI) = MTH_obs - CalcMTH_env(SI, Age*, latitude, elevation)
   f_SI <- function(SI) MTH_obs - CalcMTH_env(SI, Age_star, latitude, elevation)
   res  <- bisection_with_trace(f_SI, lo, hi, iters, step)
   si_seq <- res$trace$mid
   si_seq <- si_seq[is.finite(si_seq)]
   if (!length(si_seq)) si_seq <- res$root
   
   ages <- seq(0, 40, by = 0.25)
   cols <- if (length(si_seq) > 1) grDevices::rainbow(length(si_seq)) else "grey50"
   
   oldpar <- par(no.readonly = TRUE); on.exit(par(oldpar))
   # set plot range using first curve
   y0 <- vapply(ages, function(a) CalcMTH_env(si_seq[1], a, latitude, elevation), numeric(1))
   plot(ages, y0, type="l", col=cols[1], lwd=2,
        xlab="Age (years)", ylab="MTH (m)", main="Curve matching: SI via MTH at Age*",
        ylim = c(0, 80))   # <-- set  MTH y-axis here
      # overlay all iteration curves
   if (length(si_seq) > 1) {
     for (k in 2:length(si_seq)) {
       yk <- vapply(ages, function(a) CalcMTH_env(si_seq[k], a, latitude, elevation), numeric(1))
       lines(ages, yk, col = cols[k], lwd = 2)
     }
   }
   # final curve in black
   y_final <- vapply(ages, function(a) CalcMTH_env(res$root, a, latitude, elevation), numeric(1))
   lines(ages, y_final, col = "black", lwd = 3, lty = 2)
   # observed point
   points(Age_star, MTH_obs, pch = 19)
   legend("topleft",
          legend = c("iterations", sprintf("final SI = %.2f", res$root), "observed"),
          lwd = c(2,3,NA), pch = c(NA, NA, 19),
          col = c("grey50", "black", "black"), bty = "n")
   invisible(res)
 }
 
 show_i300_resolution <- function(T, N, DBH_obs_cm, SI, latitude, elevation,
                                  lo = 1.328, hi = 60, iters = 30L, step = 0.05) {
   # residual: g(I300) = DBH(I300, at T,N) - DBH_obs
   g_I300 <- function(I300) {
     CalcDBH_cubic(I300, SI, Age = T, N = N, latitude = latitude, elevation = elevation) - DBH_obs_cm
   }
   res  <- bisection_with_trace(g_I300, lo, hi, iters, step)
   i_seq <- res$trace$mid
   i_seq <- i_seq[is.finite(i_seq)]
   if (!length(i_seq)) i_seq <- res$root
   
   ages <- seq(max(1, T - 25), T + 25, by = 0.25)  # local window around T
   cols <- if (length(i_seq) > 1) grDevices::rainbow(length(i_seq)) else "grey50"
   
   oldpar <- par(no.readonly = TRUE); on.exit(par(oldpar))
   y0 <- vapply(ages, function(a) CalcDBH_cubic(i_seq[1], SI, Age = a, N = N, latitude = latitude, elevation = elevation), numeric(1))
   plot(ages, y0, type="l", col=cols[1], lwd=2,
        xlab="Age (years)", ylab="DBH (cm)", main="Curve matching: I300 via DBH at (T,N)",
        ylim = c(0, 50))   # <-- set  DBH y-axis here
   
   if (length(i_seq) > 1) {
     for (k in 2:length(i_seq)) {
       yk <- vapply(ages, function(a) CalcDBH_cubic(i_seq[k], SI, Age = a, N = N, latitude = latitude, elevation = elevation), numeric(1))
       lines(ages, yk, col = cols[k], lwd = 2)
     }
   }
   y_final <- vapply(ages, function(a) CalcDBH_cubic(res$root, SI, Age = a, N = N, latitude = latitude, elevation = elevation), numeric(1))
   lines(ages, y_final, col = "black", lwd = 3, lty = 2)
   points(T, DBH_obs_cm, pch = 19)
   legend("topleft",
          legend = c("iterations", sprintf("final I300 = %.2f", res$root), "observed"),
          lwd = c(2,3,NA), pch = c(NA, NA, 19),
          col = c("grey50", "black", "black"), bty = "n")
   invisible(res)
 }
 
 # SI resolution (uses  observed MTH at Age)
 i <- 5
 r <- plotsM_with_site[i, ]
 show_si_resolution(
   Age_star = as.numeric(r$Age),
   MTH_obs  = as.numeric(r$MTH),
   latitude      = as.numeric(r$latitude),
   elevation     = as.numeric(r$elevation)
 )
 
 # I300 resolution (uses  solved or provided SI, and observed DBH from BA,N)
 SI_est     <- out$SI[match(r$Plot, out$Plot)]
 DBH_obs_cm <- derive_DBH_cm(BA_m2_ha = as.numeric(r$BA), N_sph = as.numeric(r$SPH))
 show_i300_resolution(
   T           = as.numeric(r$Age),
   N           = as.numeric(r$SPH),
   DBH_obs_cm  = DBH_obs_cm,
   SI          = SI_est,
   latitude         = as.numeric(r$latitude),
   elevation        = as.numeric(r$elevation)
 )
}
