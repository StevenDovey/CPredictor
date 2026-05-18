# ==========================================================================
# Douglas-fir 500 Index Growth Model - R port of VBA Module 4
#
# Monthly timestep stand-level growth model for Douglas-fir.
# Originally by Lars Wichmann Hansen, Forest Research, NZ.
# ==========================================================================

# ---------------------------------------------------------------------------
# MTH model: predict mean top height from site index, age, and latitude
# ---------------------------------------------------------------------------
dfir_mth_calc <- function(SI, T, latitude) {
  A <- -3.7082
  B <- 0.3844
  p <- 0.0338
  q <- -0.00057
  0.25 + (SI - 0.25) *
    ((1 - exp(-exp(A) * T)) / (1 - exp(-exp(A) * 40)))^
    (1 / (B + (p + q * latitude) * SI))
}

# ---------------------------------------------------------------------------
# Mean height from MTH and stocking
# ---------------------------------------------------------------------------
dfir_mean_ht <- function(MTH, N) {
  A <- 0.106
  B <- -0.228
  MTH * (1 - A * (1 - exp(B * (N - 100) / 100)))
}

# ---------------------------------------------------------------------------
# MTH from mean height (inverse)
# ---------------------------------------------------------------------------
dfir_mh_to_mth <- function(MH, N) {
  A <- 0.106
  B <- -0.228
  1 / (1 / MH * (1 - A * (1 - exp(B * (N - 100) / 100))))
}

# ---------------------------------------------------------------------------
# Crown length of unpruned element
# ---------------------------------------------------------------------------
dfir_cl_unpruned <- function(MTH, SPH) {
  if (SPH <= 0) return(0)
  k <- 0.8429
  A <- 6.9833
  B <- 2028
  min(k * (MTH - 0.1), A + B / SPH)
}

# ---------------------------------------------------------------------------
# BA to DBH conversion
# ---------------------------------------------------------------------------
dfir_ba_to_dbh <- function(BA, SPH) {
  if (SPH <= 0) return(0)
  sqrt(BA / SPH / pi * 4) * 100
}

# ---------------------------------------------------------------------------
# DBH to BA
# ---------------------------------------------------------------------------
dfir_dbh_to_ba <- function(DBH) {
  DBH^2 * pi / 40000
}

# ---------------------------------------------------------------------------
# Volume function (Beekhuis)
# ---------------------------------------------------------------------------
dfir_volume <- function(BA, H, SPH) {
  if (BA <= 0) return(0)
  A <- 0.928
  B <- 0.3208
  BA * (A + B * H)
}

# ---------------------------------------------------------------------------
# BA from volume (inverse Beekhuis)
# ---------------------------------------------------------------------------
dfir_ba_from_vol <- function(Vol, MTH) {
  if (Vol <= 0) return(0)
  A <- 0.928
  B <- 0.3208
  Vol / (A + B * MTH)
}

# ---------------------------------------------------------------------------
# Starting DBH from SBAP
# ---------------------------------------------------------------------------
dfir_start_dbh <- function(SBAP, MTH, T, N) {
  if (T < 10) return(0.001)
  A <- 3.4; B <- 0.485; C <- 0.17; D <- 0; Fp <- 1.01; g <- -0.479
  A * (MTH - 1.4)^B * (1 + C * (SBAP - 1.9)) * (1 + Fp * T * N^g)
}

# ---------------------------------------------------------------------------
# Convert 500 Index to SBAP
# ---------------------------------------------------------------------------
dfir_five_to_sbap <- function(FiveIndex, SI) {
  (FiveIndex / (0.0971 * SI^1.344))^1.03
}

# ---------------------------------------------------------------------------
# Mortality/stocking model
# ---------------------------------------------------------------------------
dfir_sph_calc <- function(Ns, Hs, Ds, Ts, Tf, MA = 0) {
  A <- 0.00007; B <- -20.43; C <- 1.517; D <- 0.3714
  X <- exp(B + C * (log(Ns) + D * log(Hs * Ds^2)))
  Y <- A + (1 - A) * X / (1 + X)
  Ns * (1 - Y * (Tf - Ts))
}

# ---------------------------------------------------------------------------
# BA increment model
# ---------------------------------------------------------------------------
dfir_ba_calc <- function(SBAP, BAs, CL, Ts, Tf, t_last_thin_age = 0,
                         t_last_thin_ba_ratio = 1) {
  dt_thin <- Ts - t_last_thin_age
  SBAP_t <- SBAP
  if (t_last_thin_age > 0 && dt_thin > 0 && dt_thin < 12) {
    shock_reduction <- 1 - (0.15 * exp(-0.27 * dt_thin))
    shock_reduction <- max(shock_reduction, 0.1)
    SBAP_t <- SBAP * shock_reduction
  }

  B <- -0.0002059; C <- 3.0955; D <- 50; Fp <- -5.46; g <- 0.1217
  Crwn <- 1 - exp(B * CL)
  Aget <- max(1, C + (1 - C) / D * Ts)
  ratio <- BAs / SBAP_t
  Comp <- 1 - exp(Fp + g * ratio) / (1 + exp(Fp + g * ratio))
  BAs + SBAP_t * Crwn * Aget * Comp * (Tf - Ts)
}

# ---------------------------------------------------------------------------
# Thinning BA calculation
# ---------------------------------------------------------------------------
dfir_thin_calc <- function(BAs, Ns, Nf, A) {
  if (A > 0) BAs * (Nf / Ns)^A else 10
}

# ---------------------------------------------------------------------------
# Thinning indicator
# ---------------------------------------------------------------------------
dfir_thindicator <- function(Tf, Ts, thin) {
  result <- list(TT = 0, SPHf = 0, A = 0, thin_exact_age = 0)
  if (is.null(thin) || nrow(thin) == 0) return(result)
  for (i in seq_len(nrow(thin))) {
    if (thin$Age[i] > Ts && thin$Age[i] <= Tf) {
      if (thin$SPH[i] > 0 && thin$ThinCoeff[i] > 0) {
        result$TT <- 1
        result$SPHf <- thin$SPH[i]
        result$A <- thin$ThinCoeff[i]
        result$thin_exact_age <- thin$Age[i]
      }
    }
  }
  result
}

# ---------------------------------------------------------------------------
# Pruning indicator
# ---------------------------------------------------------------------------
dfir_prundicator <- function(Tf, Ts, prune) {
  if (is.null(prune) || nrow(prune) == 0) return(0)
  PT <- 0
  for (i in seq_len(nrow(prune))) {
    if (prune$Age[i] > Ts && prune$Age[i] <= Tf) {
      PT <- prune$Age[i]
    }
  }
  PT
}

# ---------------------------------------------------------------------------
# Crown height calculation
# ---------------------------------------------------------------------------
dfir_ch_calc <- function(Ts, prune) {
  aCH <- rep(0, 6)
  if (is.null(prune) || nrow(prune) == 0) return(aCH)
  for (n in seq_len(min(nrow(prune), 5))) {
    if (prune$Age[n] <= Ts) {
      aCH[n + 1] <- prune$Height[n]
    }
  }
  aCH
}

# ---------------------------------------------------------------------------
# Pruned stems distribution
# ---------------------------------------------------------------------------
dfir_pruned_stems <- function(Tf, Ts, SPH, prune, TT, pSPH, PT) {
  aSPH <- rep(0, 6)
  if (PT >= 1 || TT >= 1) {
    if (!is.null(prune) && nrow(prune) > 0) {
      for (i in seq_len(min(nrow(prune), 5))) {
        N_idx <- 6 - i
        if (prune$Age[i] <= Tf) {
          Nused <- sum(aSPH)
          if (SPH - prune$N[i] >= 0) {
            Nominal <- prune$N[i] - Nused
          } else {
            if (SPH - Nused > 0) {
              Nominal <- SPH - Nused
            } else {
              Nominal <- 0
            }
          }
        } else {
          Nominal <- 0
        }
        Alive <- if (pSPH[N_idx + 1] > 0) pSPH[N_idx + 1] else {
          if (prune$Age[i] <= Ts - 0.1) 0 else SPH
        }
        aSPH[N_idx + 1] <- min(Nominal, Alive)
      }
    }
    remainder <- SPH - sum(aSPH)
    aSPH[1] <- max(0, remainder)
  } else {
    lastsph <- sum(pSPH)
    if (lastsph != 0) {
      ratio <- SPH / lastsph
      aSPH <- pSPH * ratio
    }
  }
  aSPH
}

# ---------------------------------------------------------------------------
# Crown length calculation
# ---------------------------------------------------------------------------
dfir_cl_calc <- function(MH, MTH, SPH, pSPH, pCH) {
  CLu <- dfir_cl_unpruned(MTH, SPH)
  CL <- 0
  for (i in seq_along(pSPH)) {
    CLp <- MH - pCH[i]
    CL <- CL + pSPH[i] * min(CLu, CLp)
  }
  CL
}

# ---------------------------------------------------------------------------
# Pruned growth distribution
# ---------------------------------------------------------------------------
dfir_pruned_growth <- function(MH, MTH, SPH, pSPH, pCH, CLtotal, dBA) {
  out <- rep(0, 6)
  CLu <- dfir_cl_unpruned(MTH, SPH)
  if (CLtotal <= 0) return(out)
  for (i in seq_along(pSPH)) {
    if (pCH[i] > 0 || i == 1) {
      CLp <- MH - pCH[i]
      CL <- min(CLu, CLp)
      out[i] <- CL * pSPH[i] / CLtotal * dBA
    }
  }
  out
}

# ---------------------------------------------------------------------------
# Cumulative BA by pruned element
# ---------------------------------------------------------------------------
dfir_pba_calc <- function(pBA, pGrowth, pSPH, pSPHb4, BA, SPH, PT, TT) {
  alpha <- 0.705
  out <- pBA + pGrowth

  if (PT >= 1 || TT >= 1) {
    GrowF <- pBA + pGrowth
    for (i in 2:6) {
      if (pSPH[i] - pSPHb4[i] > 4) {
        if (pSPH[i] <= pSPHb4[i - 1] && pSPHb4[i - 1] > 0) {
          out[i] <- GrowF[i - 1] * (pSPH[i] / pSPHb4[i - 1])^alpha
          out[i - 1] <- out[i - 1] - out[i]
        } else {
          out[i] <- pBA[i - 1] + pGrowth[i - 1]
          out[i - 1] <- 0
        }
      }
      if (pSPH[i] == 0) out[i] <- 0
    }
  }

  if (TT >= 1 && abs(pSPH[1] - pSPHb4[1]) >= 1 && pSPHb4[1] > 0) {
    out[1] <- pBA[1] * (pSPH[1] / pSPHb4[1])^alpha
    for (i in 2:6) {
      if (abs(pSPH[i] - pSPHb4[i]) >= 1 && pSPHb4[i] > 0) {
        out[i] <- pBA[1] * (pSPH[i] / pSPHb4[i])^alpha
      }
    }
  }
  out
}

# ---------------------------------------------------------------------------
# SI solver: estimate site index from one height/age measurement
# ---------------------------------------------------------------------------
dfir_si_solver <- function(T1, H1, latitude, tol = 0.02) {
  uSI <- 50; dSI <- 10
  uH <- dfir_mth_calc(uSI, T1, latitude)
  dH <- dfir_mth_calc(dSI, T1, latitude)
  while (uH - dH > tol) {
    if (abs(uH - H1) > abs(dH - H1)) {
      uSI <- uSI - (uSI - dSI) / 4
    } else {
      dSI <- dSI + (uSI - dSI) / 4
    }
    uH <- dfir_mth_calc(uSI, T1, latitude)
    dH <- dfir_mth_calc(dSI, T1, latitude)
  }
  (uSI + dSI) / 2
}

# ---------------------------------------------------------------------------
# SBAP solver: estimate SBAP from one DBH measurement
# ---------------------------------------------------------------------------
dfir_sbap_solver <- function(SI, T1, D1, N0, rotlength, latitude,
                             thin = NULL, prune = NULL, tol = 0.02) {
  dI500 <- 1; uI500 <- 40
  while (abs(dI500 - uI500) > tol) {
    mI500 <- (dI500 + uI500) / 2
    SBAP <- dfir_five_to_sbap(mI500, SI)
    result <- dfir_grow(SBAP, SI, N0, rotlength, thin, prune, latitude)
    start_age <- result$monthly$Age[1]
    lookup_idx <- T1 * 12 - as.integer(start_age * 12) + 1
    lookup_idx <- max(1, min(lookup_idx, nrow(result$monthly)))
    DBH_lookup <- result$monthly$DBH[lookup_idx]
    if (DBH_lookup < D1) {
      dI500 <- mI500
    } else {
      uI500 <- mI500
    }
  }
  (dI500 + uI500) / 2
}

# ---------------------------------------------------------------------------
# Early yield correction for young stands
# ---------------------------------------------------------------------------
dfir_early_yield <- function(annual, SI, latitude) {
  initialvol <- 0.0000064
  ddbh <- 0
  for (age in 1:min(20, nrow(annual))) {
    prevdbh <- ddbh
    ddbh <- annual$DBH[age + 1]
    if (!is.na(ddbh) && ddbh >= 2 && prevdbh < 2) {
      N_age <- annual$SPHA[age + 1]
      Vol_age <- annual$Vol[age + 1]
      if (N_age > 0) {
        treevolinc <- Vol_age / N_age - initialvol
        k <- treevolinc / (age^2.7)
        for (j in seq(age - 1, 0, -1)) {
          idx <- j + 1
          if (idx < 1 || idx > nrow(annual)) next
          N_j <- annual$SPHA[idx]
          annual$Vol[idx] <- (initialvol + k * j^2.7) * N_j
          annual$MTH[idx] <- dfir_mth_calc(SI, j, latitude)
          if (annual$MTH[idx] < 1.4) {
            annual$BA[idx] <- 0
          } else {
            annual$BA[idx] <- dfir_ba_from_vol(annual$Vol[idx], annual$MTH[idx])
          }
          if (N_j > 0) {
            annual$DBH[idx] <- 200 * sqrt((annual$BA[idx] / N_j) / pi)
          }
        }
      }
      break
    }
  }
  annual
}

# ---------------------------------------------------------------------------
# Grow: main monthly simulation loop
# ---------------------------------------------------------------------------
dfir_grow <- function(SBAP, SI, SPHin, Rotation, thin = NULL, prune = NULL,
                      latitude = -42, MA = 0) {
  if (SBAP < 0.2 || SBAP >= 5 || SI <= 15 || SI >= 50 ||
      Rotation <= 20 || Rotation > 100) {
    warning("Stand parameters outside valid range for Douglas-fir model")
    return(NULL)
  }

  # Find start time via bisection
  TsU <- 14; TsD <- 0
  TsM <- (TsU + TsD) / 2
  MTH_TsM_Old <- 0
  MTH_TsM <- dfir_mth_calc(SI, TsM, latitude)
  MH_TsM <- dfir_mean_ht(MTH_TsM, SPHin)
  MTH_start <- 4

  while (abs(MTH_TsM_Old - MTH_TsM) > 0.05 &&
         (MH_TsM >= MTH_start + 0.05 || MH_TsM <= MTH_start)) {
    if (MH_TsM < MTH_start) {
      TsD <- TsM
    } else {
      TsU <- TsM
    }
    TsM <- (TsU + TsD) / 2
    MTH_TsM_Old <- MTH_TsM
    MTH_TsM <- dfir_mth_calc(SI, TsM, latitude)
    MH_TsM <- dfir_mean_ht(MTH_TsM, SPHin)
  }

  if (TsM >= 12) {
    TsM <- 12
  } else {
    TsM <- round(TsM * 12) / 12
  }

  month_s <- as.integer(TsM * 12)
  Ts <- month_s / 12
  MTHs <- dfir_mth_calc(SI, Ts, latitude)
  SPHs <- SPHin
  MHs <- dfir_mean_ht(MTHs, SPHs)
  DBHs <- dfir_start_dbh(SBAP, MTHs, Ts, SPHs)
  BAs <- DBHs^2 * pi / 40000 * SPHin

  pSPH <- rep(0, 6); pSPH[1] <- SPHs
  pCH <- dfir_ch_calc(Ts, prune)
  CLs <- dfir_cl_calc(MHs, MTHs, SPHs, pSPH, pCH)
  dBA <- 0
  pGrowth <- dfir_pruned_growth(MHs, MTHs, SPHs, pSPH, pCH, CLs, dBA)
  dummy_growth <- rep(0, 6)
  pBA <- dfir_pba_calc(rep(0, 6), dummy_growth, pSPH, rep(0, 6), BAs, SPHs, 1, 1)
  pBA[1] <- BAs + pBA[1]

  t_last_thin_age <- 0
  t_last_thin_ba_ratio <- 1
  TAge <- 0; TSPH <- 0; TBA <- 0
  ic <- 0

  out_rows <- list()

  while (Ts <= Rotation) {
    month_f <- month_s + 1
    Tf <- month_f / 12

    BAb4 <- BAs
    MTHprev <- MTHs
    pSPHb4 <- pSPH

    # Thinning check
    temp <- dfir_thindicator(Tf, Ts, thin)
    TT <- temp$TT
    SPHf <- temp$SPHf
    A_thin <- temp$A
    if (TT == 1) Tf <- temp$thin_exact_age

    MTHs <- dfir_mth_calc(SI, Tf, latitude)
    PT <- dfir_prundicator(Tf, Ts, prune)

    SPHs <- dfir_sph_calc(SPHs, MTHs, DBHs, Ts, Tf, MA)
    BAs <- dfir_ba_calc(SBAP, BAs, CLs, Ts, Tf, t_last_thin_age, t_last_thin_ba_ratio)
    dBA <- BAs - BAb4
    DBHs <- dfir_ba_to_dbh(BAs, SPHs)
    MHs <- dfir_mean_ht(MTHs, SPHs)
    pCH <- dfir_ch_calc(Tf, prune)
    pSPH <- dfir_pruned_stems(Tf, Ts, SPHs, prune, TT, pSPH, PT)
    CLs <- dfir_cl_calc(MHs, MTHs, SPHs, pSPH, pCH)

    if (TT > 0) {
      # Record pre-thin output
      ic <- ic + 1
      out_rows[[ic]] <- data.frame(
        Age = Tf, DBH = DBHs, MTH = MTHs, MH = MHs,
        SPHA = SPHs, BA = BAs, CL = CLs,
        ThinAge = TAge, ThinSPH = TSPH, ThinBA = TBA,
        Vol = dfir_volume(BAs, MTHs, SPHs),
        stringsAsFactors = FALSE
      )
      month_f <- month_f + 1
      Tf <- month_f / 12

      t_last_thin_ba_ratio <- BAs
      BAs <- dfir_thin_calc(BAs, SPHs, SPHf, A_thin)
      t_last_thin_ba_ratio <- t_last_thin_ba_ratio / BAs
      TAge <- Ts
      TBA <- BAb4 - BAs
      TSPH <- SPHs - SPHf
      SPHs <- SPHf
      DBHs <- dfir_ba_to_dbh(BAs, SPHs)
      pCH <- dfir_ch_calc(Tf, prune)
      pSPH <- dfir_pruned_stems(Tf, Ts, SPHs, prune, TT, pSPH, PT)
      CLs <- dfir_cl_calc(MHs, MTHs, SPHs, pSPH, pCH)
      t_last_thin_age <- Ts
    }

    pGrowth <- dfir_pruned_growth(MHs, MTHs, SPHs, pSPH, pCH, CLs, dBA)
    pBA <- dfir_pba_calc(pBA, pGrowth, pSPH, pSPHb4, BAs, SPHs, PT, TT)

    month_s <- month_f
    Ts <- Tf

    CLu_val <- dfir_cl_unpruned(MTHs, SPHs)
    ic <- ic + 1
    out_rows[[ic]] <- data.frame(
      Age = Tf, DBH = DBHs, MTH = MTHs, MH = MHs,
      SPHA = SPHs, BA = BAs, CL = CLs,
      ThinAge = TAge, ThinSPH = TSPH, ThinBA = TBA,
      Vol = dfir_volume(BAs, MTHs, SPHs),
      stringsAsFactors = FALSE
    )
    TBA <- 0; TSPH <- 0; TAge <- 0

    if (ic > 1350) {
      warning("Stand outside the possible range of the calculator!")
      break
    }
  }

  monthly <- do.call(rbind, out_rows)
  rownames(monthly) <- NULL
  list(monthly = monthly, start_age = TsM)
}

# ==========================================================================
# dfir_yield: main entry point matching the VBA Dfir_Yield function
# ==========================================================================
dfir_yield <- function(I500 = NULL, H30 = NULL, SI = NULL,
                       N0 = 1000, rotlength = 40,
                       latitude = -42, MAT = 12,
                       Soil_C = 5.57, Soil_N = 0.296,
                       T1 = NULL, H1 = NULL, D1 = NULL,
                       thin_schedule = NULL,
                       prune_schedule = NULL) {
  # Resolve SI
  if (is.null(SI) && !is.null(H30)) SI <- H30
  if (is.null(SI) && !is.null(T1) && !is.null(H1)) {
    SI <- dfir_si_solver(T1, H1, latitude)
  }
  if (is.null(SI)) stop("Must supply SI, H30, or T1+H1 for site index")

  # Resolve I500
  if (is.null(I500) && !is.null(D1) && !is.null(T1)) {
    I500 <- dfir_sbap_solver(SI, T1, D1, N0, rotlength, latitude, thin_schedule, prune_schedule)
  }
  if (is.null(I500)) stop("Must supply I500 or T1+D1 for productivity index")

  SBAP <- dfir_five_to_sbap(I500, SI)

  result <- dfir_grow(SBAP, SI, N0, rotlength, thin_schedule, prune_schedule, latitude)
  if (is.null(result)) return(NULL)

  # Extract annual values from monthly output
  monthly <- result$monthly
  start_age <- result$start_age

  annual_ages <- 0:rotlength
  annual <- data.frame(
    Age = annual_ages,
    SPHA = NA_real_, DBH = NA_real_, BA = NA_real_,
    MTH = NA_real_, Vol = NA_real_,
    stringsAsFactors = FALSE
  )
  annual$SPHA[1] <- N0

  for (t_val in seq_len(rotlength)) {
    ic_idx <- t_val * 12 - as.integer(start_age * 12) + 1
    if (ic_idx >= 1 && ic_idx <= nrow(monthly)) {
      annual$SPHA[t_val + 1] <- monthly$SPHA[ic_idx]
      annual$MTH[t_val + 1] <- monthly$MTH[ic_idx]
      annual$DBH[t_val + 1] <- monthly$DBH[ic_idx]
      annual$BA[t_val + 1] <- monthly$BA[ic_idx]
      annual$Vol[t_val + 1] <- monthly$Vol[ic_idx]
    }
  }

  # Apply early yield correction
  annual <- dfir_early_yield(annual, SI, latitude)

  # Wood density model
  Adj_Soil_C_N <- if (Soil_N > 0.14) Soil_C / (Soil_N - 0.14) else 50
  BH_Age <- -(1 / exp(-3.7082)) * log(1 - ((1.4 - 0.25) / (SI - 0.25))^
    (0.3844 + (0.0338 - 0.00057 * latitude) * SI) *
    (1 - exp(-exp(-3.7082) * 40)))
  Age_BH_ring_30_formed <- 30.5 + BH_Age
  Density_ring_30 <- 136.5 + 23.3 * MAT + 3.09 * Adj_Soil_C_N
  local_parameter <- (Density_ring_30 - 432.6 - (1.22 - 30) /
    (0.0235 + 0.0125 * exp(0.221 * 30))) /
    (1 - 0.814 * exp(-30 * 0.258))

  annual$WoodDensity <- NA_real_
  for (t_val in 0:rotlength) {
    BH_ring_from_pith <- t_val - BH_Age - 0.5
    if (BH_ring_from_pith < 0) BH_ring_from_pith <- 0
    Sheath_ratio <- 1.094 - 0.0277 * sqrt(BH_ring_from_pith)
    BH_ring_density <- 432.6 + (1.22 - BH_ring_from_pith) /
      (0.0235 + 0.0125 * exp(0.221 * BH_ring_from_pith)) +
      (1 - 0.814 * exp(-0.258 * BH_ring_from_pith)) * local_parameter
    annual$WoodDensity[t_val + 1] <- Sheath_ratio * BH_ring_density
  }

  list(
    annual = annual,
    monthly = monthly,
    SI = SI, I500 = I500, SBAP = SBAP,
    start_age = start_age,
    BH_Age = BH_Age,
    Density_ring_30 = Density_ring_30
  )
}
