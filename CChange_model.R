# ==========================================================================
# CChange Carbon Model - R port of VBA Module 5 (originally FORTRAN)
#
# Compartment model for dynamically simulating dry matter and nutrient
# content of managed radiata pine stands over a rotation.
#
# State variables X(1..16):
#   X(1)  = abstract (equivalent to Z, net primary production)
#   X(2)  = abstract (equivalent to Z minus mortality flows)
#   X(3)  = 0-1 year old needle
#   X(4)  = 1-2 year old needle
#   X(5)  = 2+ year old needle
#   X(6)  = live branch
#   X(7)  = dead branch
#   X(8)  = stem wood
#   X(9)  = coarse root
#   X(10) = needle litter
#   X(11) = branch litter
#   X(12) = stem litter
#   X(13) = coarse root litter
#   X(14) = live fine root
#   X(15) = fine root litter
#   X(16) = stem bark
# ==========================================================================

# ---------------------------------------------------------------------------
# Nutrient concentration functions (Puruki reference site)
# ---------------------------------------------------------------------------

puruki_stemwd_nconc <- function(t) 1.0044 * (t + 2.753)^(-0.848)
puruki_stembk_nconc <- function(t) 9.8398 * (t + 6.199)^(-1.0672)
puruki_livebr_nconc <- function(t) 1.4806 - 1.2025 * (1 - exp(-0.5452 * t))
puruki_deadbr_nconc <- function(t) 3.5524 - 3.2556 * (1 - exp(-0.5595 * t))
puruki_fol_1yr_nconc <- function(t) 1.5984
puruki_fol_2yr_nconc <- function(t) {
  val <- 1.4859 - 1 / (2.232 + 0.2803 * (t - 5) + 0.8822 * (t - 5)^2)
  max(val, 0.975)
}
puruki_fol_3yr_nconc <- function(t) {
  val <- 1.2939 - 1 / (2.277 - 4.0399 * (t - 5) + 5.6249 * (t - 5)^2)
  max(val, 0.975)
}
puruki_needle_fall_nconc <- function(t) 0.975

puruki_stemwd_pconc <- function(t) 0.358 * (t + 3.1561)^(-1.1618)
puruki_stembk_pconc <- function(t) 1.1094 * (t + 3.7653)^(-1.1175)
puruki_livebr_pconc <- function(t) 0.2907 - 0.2435 * (1 - exp(-0.5442 * t))
puruki_deadbr_pconc <- function(t) 0.162 - 0.1376 * (1 - exp(-0.2642 * t))
puruki_fol_1yr_pconc <- function(t) 0.1848
puruki_fol_2yr_pconc <- function(t) {
  val <- 0.1546 - 1 / (29.8995 - 42.7266 * (t - 5) + 39.5945 * (t - 5)^2)
  max(val, 0.137)
}
puruki_fol_3yr_pconc <- function(t) 0.137
puruki_needle_fall_pconc <- function(t) 0.0827

# ---------------------------------------------------------------------------
# Nconc: nitrogen concentration (%) of component i at age t
# ---------------------------------------------------------------------------
nconc <- function(i, t, soil_c = 5.57, soil_n = 0.296) {
  if (soil_c == 0) soil_c <- 5.57
  if (soil_n == 0) soil_n <- 0.296
  cn_adj <- if (soil_n <= 0.014) 50 else soil_c / (soil_n - 0.014)
  cn_adj <- min(cn_adj, 50)

  if (i == 3) {
    age13 <- 1.9015 - 0.01791 * cn_adj
    return(puruki_fol_1yr_nconc(t) * age13 / puruki_fol_1yr_nconc(13))
  }
  if (i == 4 || i == 14) {
    age13 <- 1.6725 - 0.016149 * cn_adj
    return(puruki_fol_2yr_nconc(t) * age13 / puruki_fol_2yr_nconc(13))
  }
  if (i == 5) {
    age13 <- 1.6725 - 0.016149 * cn_adj
    nc2 <- puruki_fol_2yr_nconc(t) * age13 / puruki_fol_2yr_nconc(13)
    nc_nf <- (1.5105 - 0.019599 * cn_adj) * 0.828
    denom <- puruki_fol_2yr_nconc(t) - puruki_needle_fall_nconc(t)
    if (abs(denom) < 1e-12) return(nc2)
    return(nc2 - (nc2 - nc_nf) *
      (puruki_fol_2yr_nconc(t) - puruki_fol_3yr_nconc(t)) / denom)
  }
  if (i == 10) {
    return((1.5105 - 0.019599 * cn_adj) * 0.828)
  }
  if (i == 16) {
    age13 <- 0.4566 - 0.0043832 * cn_adj
    return(puruki_stembk_nconc(t) * age13 / puruki_stembk_nconc(13))
  }
  if (i == 6) {
    age13 <- 0.0858 - 0.001193 * cn_adj
    return((age13 / puruki_stemwd_nconc(13)) * puruki_livebr_nconc(t))
  }
  if (i %in% c(7, 11, 13)) {
    age13 <- 0.0858 - 0.001193 * cn_adj
    return((age13 / puruki_stemwd_nconc(13)) * puruki_deadbr_nconc(t))
  }
  if (i %in% c(8, 9)) {
    age13 <- 0.0858 - 0.001193 * cn_adj
    return(puruki_stemwd_nconc(t) * age13 / puruki_stemwd_nconc(13))
  }
  0
}

# ---------------------------------------------------------------------------
# Pconc: phosphorus concentration (%) of component i at age t
# ---------------------------------------------------------------------------
pconc <- function(i, t, soil_c = 5.57, soil_organic_p = 333) {
  if (soil_c == 0) soil_c <- 5.57
  if (soil_organic_p == 0) soil_organic_p <- 333
  ln_cp <- log(soil_c / soil_organic_p / 10000)

  if (i == 3) {
    age13 <- -0.35538 - 0.037992 * ln_cp
    return(puruki_fol_1yr_pconc(t) * age13 / puruki_fol_1yr_pconc(13))
  }
  if (i == 4 || i == 14) {
    age13 <- -0.26621 - 0.030253 * ln_cp
    return(puruki_fol_2yr_pconc(t) * age13 / puruki_fol_2yr_pconc(13))
  }
  if (i == 5) {
    age13 <- -0.26621 - 0.030253 * ln_cp
    pc2 <- puruki_fol_2yr_pconc(t) * age13 / puruki_fol_2yr_pconc(13)
    pc_nf <- 0.0827
    denom <- puruki_fol_2yr_pconc(t) - puruki_needle_fall_pconc(t)
    if (abs(denom) < 1e-12) return(pc2)
    return(pc2 - (pc2 - pc_nf) *
      (puruki_fol_2yr_pconc(t) - puruki_fol_3yr_pconc(t)) / denom)
  }
  if (i == 10) return(0.0827)
  if (i == 16) return(puruki_stembk_pconc(t))
  if (i == 6) return(puruki_livebr_pconc(t))
  if (i %in% c(7, 11, 13)) return(puruki_deadbr_pconc(t))
  if (i %in% c(8, 9)) return(puruki_stemwd_pconc(13))
  0
}

# ---------------------------------------------------------------------------
# EULER: forward Euler numerical integration of state transition equations
# ---------------------------------------------------------------------------
cc_euler <- function(A, X, NAP) {
  DT <- 1
  NSOL <- as.integer(1 / DT)
  z <- NAP * 3

  for (step in seq_len(NSOL)) {
    DX <- numeric(16)
    DX[1]  <- DT * (z + A[1, 1] * X[1])
    DX[2]  <- DT * (A[2, 1] * X[1] + A[2, 2] * X[2])
    DX[3]  <- DT * (A[3, 2] * X[2] + A[3, 3] * X[3])
    DX[4]  <- DT * (A[4, 3] * X[3] + A[4, 4] * X[4])
    DX[5]  <- DT * (A[5, 4] * X[4] + A[5, 5] * X[5])
    DX[6]  <- DT * (A[6, 2] * X[2] + A[6, 6] * X[6])
    DX[7]  <- DT * (A[7, 6] * X[6] + A[7, 7] * X[7])
    DX[8]  <- DT * (A[8, 2] * X[2] + A[8, 8] * X[8])
    DX[9]  <- DT * (A[9, 2] * X[2] + A[9, 9] * X[9])
    DX[10] <- DT * (A[10, 3] * X[3] + A[10, 4] * X[4] + A[10, 5] * X[5] + A[10, 10] * X[10])
    DX[11] <- DT * (A[11, 7] * X[7] + A[11, 11] * X[11])
    DX[12] <- DT * (A[12, 8] * X[8] + A[12, 16] * X[16] + A[12, 12] * X[12])
    DX[13] <- DT * (A[13, 9] * X[9] + A[13, 13] * X[13])
    DX[14] <- DT * (A[14, 2] * X[2] + A[14, 14] * X[14])
    DX[15] <- DT * (A[15, 14] * X[14] + A[15, 15] * X[15])
    DX[16] <- DT * (A[16, 2] * X[2] + A[16, 16] * X[16])
    X <- X + DX
  }
  X
}

# ---------------------------------------------------------------------------
# DMAMAT: compute rate matrix A from flow matrix F and state vector X
# ---------------------------------------------------------------------------
cc_dmamat <- function(F, X, Y) {
  A <- matrix(0, 16, 16)

  safe_div <- function(num, den) if (den != 0) num / den else 0

  # Compartment 1
  if (X[1] != 0) {
    A[1, 1] <- -1 * (F[2, 1] + Y[1]) / X[1]
    A[2, 1] <- F[2, 1] / X[1]
  }
  # Compartment 2
  if (X[2] != 0) {
    A[2, 2] <- -1 * (F[3, 2] + F[6, 2] + F[8, 2] + F[9, 2] + F[14, 2] + Y[2]) / X[2]
    A[3, 2] <- F[3, 2] / X[2]
    A[6, 2] <- F[6, 2] / X[2]
    A[8, 2] <- F[8, 2] / X[2]
    A[9, 2] <- F[9, 2] / X[2]
    A[14, 2] <- F[14, 2] / X[2]
    A[16, 2] <- F[16, 2] / X[2]
  }
  # Compartment 3
  if (X[3] != 0) {
    A[3, 3] <- -1 * (F[4, 3] + F[10, 3] + Y[3]) / X[3]
    A[4, 3] <- F[4, 3] / X[3]
    A[10, 3] <- F[10, 3] / X[3]
  }
  # Compartment 4
  if (X[4] != 0) {
    A[4, 4] <- -1 * ((F[5, 4] + F[10, 4]) / X[4])
    A[5, 4] <- F[5, 4] / X[4]
    A[10, 4] <- F[10, 4] / X[4]
  }
  # Compartment 5
  if (X[5] != 0) {
    A[5, 5] <- -1 * (F[10, 5] / X[5])
    A[10, 5] <- F[10, 5] / X[5]
  }
  # Compartment 6
  if (X[6] != 0) {
    A[6, 6] <- -1 * (F[7, 6] / X[6])
    A[7, 6] <- F[7, 6] / X[6]
  }
  # Compartment 7
  if (X[7] != 0) {
    A[7, 7] <- -1 * (F[11, 7] + Y[7]) / X[7]
    A[11, 7] <- F[11, 7] / X[7]
  }
  # Compartment 8
  if (X[8] != 0) {
    A[8, 8] <- -1 * (F[12, 8] / X[8])
    A[12, 8] <- F[12, 8] / X[8]
  }
  # Compartment 9
  if (X[9] != 0) {
    A[9, 9] <- -1 * (F[13, 9] / X[9])
    A[13, 9] <- F[13, 9] / X[9]
  }
  # Compartment 10
  if (X[10] != 0) A[10, 10] <- -1 * (Y[10] / X[10])
  # Compartment 11
  if (X[11] != 0) A[11, 11] <- -1 * (Y[11] / X[11])
  # Compartment 12
  if (X[12] != 0) A[12, 12] <- -1 * (Y[12] / X[12])
  if (X[16] != 0) A[12, 16] <- F[12, 16] / X[16]
  # Compartment 13
  if (X[13] != 0) A[13, 13] <- -1 * (Y[13] / X[13])
  # Compartment 14
  if (X[14] != 0) {
    A[14, 14] <- -1 * (F[15, 14] / X[14])
    A[15, 14] <- F[15, 14] / X[14]
  }
  # Compartment 15
  if (X[15] != 0) A[15, 15] <- -1 * (Y[15] / X[15])
  # Compartment 16
  if (X[16] != 0) A[16, 16] <- -1 * (F[12, 16] / X[16])

  A
}

# ---------------------------------------------------------------------------
# DMFMAT: generate flow matrix F, partitioning production to components
# ---------------------------------------------------------------------------
cc_dmfmat <- function(X, NAP, t, SPHA, BA, HT, GCL, HTBGC,
                      CONS, SX3, INDEAD, INPNAP, F128SP_val,
                      X6MAX, AM1, AM2, BM1, BM2,
                      AY1A, AY1B, AY1C,
                      MATEMP = 12) {
  F <- matrix(0, 16, 16)
  Y <- numeric(16)

  # Flow from compartment 1 (NAP buffer) to compartment 2 (allocation pool)
  F[2, 1] <- NAP
  PT <- t + 1
  PROOT <- 0.3
  BG <- NAP * PROOT
  TNAPAG <- NAP - BG

  # Partition above-ground production to new needles
  PFOL <- ((39.5801 - 12) / (1 + (PT / 7.867)^3.3371) + 12) / 100
  F[3, 2] <- TNAPAG * PFOL

  # Partition to stem wood (excl bark)
  PSTMW <- ((65 - 28.5213) / (1 + (PT / 8.3314)^(-2.9473)) + 28.5213) / 100
  F[8, 2] <- TNAPAG * PSTMW

  # Partition to stem wood + bark
  PSTMWB <- ((72 - 33.8497) / (1 + (PT / 8.7097)^(-2.9383)) + 33.8497) / 100
  F[16, 2] <- TNAPAG * PSTMWB - F[8, 2]

  # Scaler for stem bark mortality
  STMMORT <- if (F128SP_val != 0) (TNAPAG * PSTMWB) / F128SP_val else 1

  # Partition to branches + reproduction (by difference)
  PBRRP <- 1 - PFOL - PSTMWB
  F[6, 2] <- TNAPAG * PBRRP

  # Partition BG to coarse and fine root
  F[9, 2] <- (F[8, 2] + F[6, 2]) * 0.2
  F[14, 2] <- BG - F[9, 2]

  # Stem mortality from external model
  F[12, 8] <- 0
  F[12, 16] <- 0
  if (INPNAP == 1) {
    F[12, 8] <- F128SP_val
    F[12, 16] <- (F128SP_val * STMMORT) - F128SP_val
  }

  # Consumption of 1yr fascicles
  Y[3] <- CONS * X[3]

  # Fascicle retention
  X3RET <- SX3
  X4RET <- SX3 / 2
  X3LOSS <- 1 - X3RET
  X4LOSS <- 1 - X4RET

  # Litter flows from needle age classes
  F[10, 3] <- X[3] * X3LOSS
  F[10, 4] <- X[4] * X4LOSS
  F[10, 5] <- X[5]
  F[11, 7] <- 0

  # Recruitment to older needle age classes
  F[4, 3] <- X[3] - F[10, 3]
  F[5, 4] <- X[4] - F[10, 4]

  # Branch mortality
  if (X[6] >= X6MAX) {
    F[7, 6] <- F[6, 2]
  } else {
    RAT <- X[6] / X6MAX
    F[7, 6] <- F[6, 2] * RAT^4
  }

  # Coarse root mortality (10% of stem mortality)
  F[13, 9] <- F[12, 8] * 0.1
  # Fine root turnover
  F[15, 14] <- X[14] * 1.5

  # Decomposition: temperature-dependent decay constants
  if (MATEMP == 0) MATEMP <- 12
  DECAY_NEEDLE <- 0.081 * exp(0.093 * MATEMP)
  DECAY_BRANCH <- 0.0429 * exp(0.093 * MATEMP)
  DECAY_STEM <- 0.0376 * exp(0.093 * MATEMP)
  DECAY_CROOT <- 0.0684 * exp(0.093 * MATEMP)
  PROP_NEEDLE <- 1 - exp(-DECAY_NEEDLE)
  PROP_BRANCH <- 1 - exp(-DECAY_BRANCH)
  PROP_ABRANCH <- PROP_BRANCH * 0.0607 / 0.0936
  PROP_STEM <- 1 - exp(-DECAY_STEM)
  PROP_CROOT <- 1 - exp(-DECAY_CROOT)

  Y[10] <- X[10] * PROP_NEEDLE
  Y[11] <- X[11] * PROP_BRANCH
  Y[12] <- X[12] * PROP_STEM
  Y[13] <- X[13] * PROP_CROOT
  Y[7]  <- X[7] * PROP_ABRANCH
  Y[15] <- X[15] * 0.52

  # Adjust abstract compartments
  TMORT <- F[7, 6] + F[12, 8] + F[13, 9] + F[15, 14] + Y[3] +
           F[10, 3] + F[10, 4] + F[10, 5]
  Y[1] <- X[1]
  Y[2] <- X[2] + TMORT

  list(
    F = F, Y = Y,
    PROP_NEEDLE = PROP_NEEDLE,
    PROP_BRANCH = PROP_BRANCH,
    PROP_STEM = PROP_STEM,
    PROP_CROOT = PROP_CROOT,
    PROP_ABRANCH = PROP_ABRANCH
  )
}

# ---------------------------------------------------------------------------
# STRUCT: simulate stand structural features
# ---------------------------------------------------------------------------
cc_struct <- function(X, t, SPHA, BA, HT, GCL, HTBGC, HINC,
                      B31, B32, S, SX3, AM1, AM2, BM1, BM2,
                      AY1A, AY1B, AY1C, SFINT, SFSPE,
                      HTSP, BASP, SPHSP, TVOLSP,
                      ADIST, ID, DENSREG, DENWH, DENINC,
                      X8GAIN, INPNAP) {
  HOLD <- HT
  t_new <- t + 1
  ISPT <- as.integer(t_new)

  BSS <- S / (1 - exp(B31 * S * 20))^B32
  HT_new <- BSS * (1 - exp(B31 * S * t_new))^B32

  if (length(HTSP) >= ISPT && HTSP[ISPT] > 0) HT_new <- HTSP[ISPT + ID - 1]
  if (length(SPHSP) >= ISPT && SPHSP[ISPT] > 0) SPHA <- SPHSP[ISPT + ID - 1]
  if (length(BASP) >= ISPT && BASP[ISPT] > 0) BA <- BASP[ISPT + ID - 1]

  HINC_new <- HT_new - HOLD

  X3RET <- SX3
  X3MAX <- AM1 * exp(AM2 * (t_new + 1))
  X4MAX <- X3MAX * X3RET
  X5MAX <- X4MAX * X3RET * 0.5

  LA1Y <- (AY1A * X[3]^2 + AY1B * X[3] + AY1C) * X[3]
  LA2Y <- (AY1A * X[4]^2 + AY1B * X[4] + AY1C) * X[4]
  LA3Y <- (AY1A * X[5]^2 + AY1B * X[5] + AY1C) * X[5]
  MAXL1Y <- (AY1A * X3MAX^2 + AY1B * X3MAX + AY1C) * X3MAX

  GCLN <- GCL + HINC_new
  if (SPHA >= 1000) {
    GCLMAX <- HT_new^0.45077 * (sqrt(10000 / SPHA))^0.32111 * 2.43641
  } else {
    GCLMAX <- HT_new^0.8888 * (sqrt(10000 / SPHA))^0.20861 * 0.77889
  }
  GCL_new <- min(GCLN, GCLMAX)
  HTBGC_new <- HT_new - GCL_new

  if (BASP[min(ISPT, length(BASP))] == 0 && HT_new > 0) {
    BA <- ((X[8] + X[16]) - SFINT * SPHA) / (HT_new * SFSPE)
    BA <- max(BA, 0)
  }

  X6MAX <- BM1 * (1 / SPHA)^BM2
  TLA <- LA1Y + LA2Y + LA3Y

  dens <- 0
  Vol <- 0
  TVOL <- 0
  if (HT_new > 1.4 && SPHA > 0) {
    DBHBAR <- sqrt(BA * 4 / (SPHA * 3.1416))
    C <- 3.06
    VOLBAR <- (DBHBAR^2) * ((HT_new - 1.4)^(1 - C)) * (HT_new^C) / C
    TVOL <- VOLBAR * SPHA

    if (DENSREG != 6) {
      if (dens > 0) Vol <- X[8] / dens
    }
    if (INPNAP == 1 && DENSREG == 6 && length(DENINC) >= ISPT) {
      Vol <- Vol + X8GAIN / DENINC[ISPT + ID - 1]
    }
    if (length(TVOLSP) >= ISPT && TVOLSP[ISPT] > 0) {
      Vol <- TVOLSP[ISPT + ID - 1]
    }
    if (Vol > 0 && DENSREG == 6) dens <- X[8] / Vol
  }

  list(
    HT = HT_new, GCL = GCL_new, HTBGC = HTBGC_new,
    HINC = HINC_new, BA = BA, SPHA = SPHA,
    X3MAX = X3MAX, X6MAX = X6MAX, TLA = TLA,
    LA1Y = LA1Y, LA2Y = LA2Y, LA3Y = LA3Y,
    MAXL1Y = MAXL1Y, GCLMAX = GCLMAX,
    dens = dens, Vol = Vol, TVOL = TVOL,
    t = t_new
  )
}

# ---------------------------------------------------------------------------
# DISTUR: initialise or perturb compartments at disturbance events
# ---------------------------------------------------------------------------
cc_distur_init <- function(SPHA_init, AVERX, DMINI, FLOSS,
                           B31, B32, S, CONS, SX3, RADIA, IROT,
                           floor_init = NULL,
                           YESN = 0, soil_c = 5.57, soil_n = 0.296,
                           soil_organic_p = 333) {
  X <- numeric(16)
  XN <- numeric(16)
  XP <- numeric(16)

  for (i in 1:7) X[i + 2] <- AVERX[i] * SPHA_init / 1000
  X[14] <- AVERX[8] * SPHA_init / 1000
  X[8] <- AVERX[6] * SPHA_init * 0.75 / 1000
  X[16] <- AVERX[6] * SPHA_init * 0.25 / 1000
  X[1] <- 5
  X[2] <- 5

  if (IROT == 1 && !is.null(floor_init)) {
    X[10] <- floor_init[1]
    X[11] <- floor_init[2]
    X[12] <- floor_init[3]
    X[13] <- floor_init[4]
    X[15] <- floor_init[5]
  }

  AY1A <- 0.009617
  AY1B <- -0.08771
  AY1C <- 1.7122

  LA1Y <- (AY1A * X[3]^2 + AY1B * X[3] + AY1C) * X[3]
  LA2Y <- (AY1A * X[4]^2 + AY1B * X[4] + AY1C) * X[4]
  LA3Y <- (AY1A * X[5]^2 + AY1B * X[5] + AY1C) * X[5]
  TLA <- LA1Y + LA2Y + LA3Y

  AM1 <- 10.5
  AM2 <- -0.0301704

  if (YESN == 1) {
    for (i in 3:9) {
      XN[i] <- X[i] * nconc(i, 0, soil_c, soil_n) * 10
      XP[i] <- X[i] * pconc(i, 0, soil_c, soil_organic_p) * 10
    }
    XN[14] <- X[14] * nconc(14, 0, soil_c, soil_n) * 10
    XN[16] <- X[16] * nconc(16, 0, soil_c, soil_n) * 10
    XP[14] <- X[14] * pconc(14, 0, soil_c, soil_organic_p) * 10
    XP[16] <- X[16] * pconc(16, 0, soil_c, soil_organic_p) * 10
  }

  list(
    X = X, XN = XN, XP = XP,
    AY1A = AY1A, AY1B = AY1B, AY1C = AY1C,
    AM1 = AM1, AM2 = AM2,
    LA1Y = LA1Y, LA2Y = LA2Y, LA3Y = LA3Y,
    TLA = TLA,
    DMINI = DMINI, FLOSS = FLOSS,
    DMLOSI = DMINI * FLOSS,
    DMLOSD = DMINI - DMINI * FLOSS
  )
}

# ---------------------------------------------------------------------------
# cc_distur_thin: apply thinning disturbance to compartments
# ---------------------------------------------------------------------------
cc_distur_thin <- function(X, XN, XP, BA, SPHA, BAN_val, XTRACT_val,
                           XTRACC_val, XTRACF_val, YESN = 0, t = 0,
                           soil_c = 5.57, soil_n = 0.296,
                           soil_organic_p = 333) {
  if (SPHA == 0) BAN_val <- 0
  if (BA <= 0) return(list(X = X, XN = XN, XP = XP, XTR = 0, XTF = 0, BA = BA, Vol = 0))

  TV <- numeric(16)
  TV[3]  <- (X[3]  / BA) * BAN_val
  TV[4]  <- (X[4]  / BA) * BAN_val
  TV[5]  <- (X[5]  / BA) * BAN_val
  TV[6]  <- (X[6]  / BA) * BAN_val
  TV[7]  <- (X[7]  / BA) * BAN_val
  TV[8]  <- (X[8]  / BA) * BAN_val
  TV[9]  <- (X[9]  / BA) * BAN_val
  TV[14] <- (X[14] / BA) * BAN_val
  TV[16] <- (X[16] / BA) * BAN_val

  XTRFOL <- 0; XTRBRC <- 0; XTRSTM <- 0
  if (XTRACC_val > 0 || XTRACC_val == -1) {
    XTRFOL <- (X[3] + X[4] + X[5] - (TV[3] + TV[4] + TV[5])) * XTRACC_val
  }
  X[10] <- X[10] + X[3] + X[4] + X[5] - (TV[3] + TV[4] + TV[5] + XTRFOL)

  if (XTRACC_val > 0 || XTRACC_val == -1) {
    XTRBRC <- (X[6] + X[7] - (TV[6] + TV[7])) * XTRACC_val
  }
  X[11] <- X[11] + X[6] + X[7] - (TV[6] + TV[7] + XTRBRC)

  if (XTRACT_val > 0) {
    XTRSTM <- (X[8] + X[16] - TV[8] - TV[16]) * XTRACT_val
  }
  X[12] <- X[12] + X[8] + X[16] - (TV[8] + TV[16] + XTRSTM)
  X[13] <- X[13] + X[9] - TV[9]
  X[15] <- X[15] + X[14] - TV[14]

  XTR <- XTRFOL + XTRBRC + XTRSTM

  XTF <- 0
  if (XTRACF_val > 0) {
    XTF <- (X[10] + X[11] + X[12]) * XTRACF_val
    X[10] <- X[10] * (1 - XTRACF_val)
    X[11] <- X[11] * (1 - XTRACF_val)
    X[12] <- X[12] * (1 - XTRACF_val)
  }

  X[3] <- TV[3]; X[4] <- TV[4]; X[5] <- TV[5]
  X[6] <- TV[6]; X[7] <- TV[7]; X[8] <- TV[8]
  X[9] <- TV[9]; X[14] <- TV[14]; X[16] <- TV[16]

  list(X = X, XN = XN, XP = XP, XTR = XTR, XTF = XTF, BA = BAN_val)
}

# ---------------------------------------------------------------------------
# OUTPUT: compute carbon summaries from dry matter
# ---------------------------------------------------------------------------
cc_output <- function(X, t, DMLOSI, DMLOSD, DMINI, ICOUNT, CSHRUB_prev = 0) {
  DMDEF <- 0
  if (DMINI > 0) DMDEF <- 0
  PRVEGR <- DMLOSD * exp(-0.18 * t)
  CSHRUB <- (PRVEGR + DMDEF + 0.1778 * t + 0.0064 * t^2) / 2

  if (ICOUNT > 1) CSHRUB <- CSHRUB_prev

  CX <- numeric(16)
  CTREES <- 0
  for (i in 3:16) {
    CX[i] <- X[i] / 2
    CTREES <- CTREES + CX[i]
  }
  CSTAND <- CTREES + CSHRUB
  CFAS <- CX[3] + CX[4] + CX[5]
  CSTEM <- CX[8] + CX[16]
  CROOTL <- CX[9] + CX[14]
  CROOTD <- CX[13] + CX[15]

  list(
    CX = CX, CTREES = CTREES, CSHRUB = CSHRUB, CSTAND = CSTAND,
    CFAS = CFAS, CSTEM = CSTEM, CROOTL = CROOTL, CROOTD = CROOTD,
    C_branch_live = CX[6], C_branch_dead = CX[7],
    C_needle_litter = CX[10], C_branch_litter = CX[11],
    C_stem_litter = CX[12]
  )
}

# ---------------------------------------------------------------------------
# NUTRIENT: simulate nutrient content and uptake
# ---------------------------------------------------------------------------
cc_nutrient <- function(F, X, XMN1, XN, XP, XNMN1, XPMN1, t,
                        PROP_NEEDLE, PROP_BRANCH, PROP_STEM,
                        PROP_CROOT, PROP_ABRANCH,
                        soil_c = 5.57, soil_n = 0.296,
                        soil_organic_p = 333) {
  # Update nitrogen content from current state and concentrations
  for (i in c(3, 4, 5)) {
    XN[i] <- X[i] * nconc(i, t, soil_c, soil_n) * 10
    XP[i] <- X[i] * pconc(i, t, soil_c, soil_organic_p) * 10
  }
  XN[6] <- X[6] * nconc(6, t, soil_c, soil_n) * 10
  XP[6] <- X[6] * pconc(6, t, soil_c, soil_organic_p) * 10
  XN[8] <- X[8] * nconc(8, t, soil_c, soil_n) * 10
  XP[8] <- X[8] * pconc(8, t, soil_c, soil_organic_p) * 10
  XN[16] <- X[16] * nconc(16, t, soil_c, soil_n) * 10
  XP[16] <- X[16] * pconc(16, t, soil_c, soil_organic_p) * 10
  XN[7] <- X[7] * nconc(16, t, soil_c, soil_n) * 10
  XP[7] <- X[7] * pconc(16, t, soil_c, soil_organic_p) * 10
  XN[9] <- X[9] * nconc(9, t, soil_c, soil_n) * 10
  XN[14] <- X[14] * nconc(14, t, soil_c, soil_n) * 10
  XP[9] <- X[9] * pconc(9, t, soil_c, soil_organic_p) * 10
  XP[14] <- X[14] * pconc(14, t, soil_c, soil_organic_p) * 10

  # Decay-related nutrient losses
  FLC <- nconc(10, t, soil_c, soil_n)
  FLCP <- pconc(10, t, soil_c, soil_organic_p)

  decayN <- numeric(16)
  decayP <- numeric(16)
  if (X[10] != 0) decayN[10] <- XN[10] * PROP_NEEDLE * 0.5
  if (X[11] != 0) decayN[11] <- XN[11] * PROP_BRANCH * 0.5
  if (X[12] != 0) decayN[12] <- XN[12] * PROP_STEM * 0.5
  if (X[13] != 0) decayN[13] <- XN[13] * PROP_CROOT * 0.5
  DecayN_total <- decayN[10] + decayN[11] + decayN[12] + decayN[13]

  if (X[10] != 0) decayP[10] <- XP[10] * PROP_NEEDLE * 0.76
  if (X[11] != 0) decayP[11] <- XP[11] * PROP_BRANCH * 0.76
  if (X[12] != 0) decayP[12] <- XP[12] * PROP_STEM * 0.76
  if (X[13] != 0) decayP[13] <- XP[13] * PROP_CROOT * 0.76
  DecayP_total <- decayP[10] + decayP[11] + decayP[12] + decayP[13]

  # Update litter N and P pools
  XN[10] <- XN[10] + (F[10, 3] * FLC + F[10, 4] * FLC + F[10, 5] * FLC) * 10 - decayN[10]
  XN[11] <- XN[11] - decayN[11]
  DSN <- 0
  if (X[8] != 0) DSN <- XN[8] * (F[12, 8] / X[8])
  XN[12] <- XN[12] + DSN - decayN[12]
  DRNC <- 0
  if (X[9] != 0) DRNC <- XN[9] * (F[13, 9] / X[9])
  XN[13] <- XN[13] + DRNC - decayN[13]

  XP[10] <- XP[10] + (F[10, 3] * FLCP + F[10, 4] * FLCP + F[10, 5] * FLCP) * 10 - decayP[10]
  XP[11] <- XP[11] - decayP[11]
  DSP <- 0
  if (X[8] != 0) DSP <- XP[8] * (F[12, 8] / X[8])
  XP[12] <- XP[12] + DSP - decayP[12]
  DRPC <- 0
  if (X[9] != 0) DRPC <- XP[9] * (F[13, 9] / X[9])
  XP[13] <- XP[13] + DRPC - decayP[13]

  # Nutrient uptake
  UPNAG <- sum(XN[c(3:8, 16)]) - sum(XNMN1[c(3:8, 16)])
  UPNBG <- sum(XN[c(9, 14)]) - sum(XNMN1[c(9, 14)])
  UPNFF <- sum(XN[10:13]) - sum(XNMN1[10:13])
  DeficitN <- UPNAG + UPNBG + UPNFF

  UPPAG <- sum(XP[c(3:8, 16)]) - sum(XPMN1[c(3:8, 16)])
  UPPBG <- sum(XP[c(9, 14)]) - sum(XPMN1[c(9, 14)])
  UPPFF <- sum(XP[10:13]) - sum(XPMN1[10:13])
  DeficitP <- UPPAG + UPPBG + UPPFF

  list(
    XN = XN, XP = XP,
    UPNAG = UPNAG, UPPAG = UPPAG,
    UPNBG = UPNBG, UPPBG = UPPBG,
    UPNFF = UPNFF, UPPFF = UPPFF,
    DecayN = DecayN_total, DecayP = DecayP_total,
    DeficitN = DeficitN, DeficitP = DeficitP
  )
}

# ==========================================================================
# DRYMAT: main entry point - run the CChange simulation
# ==========================================================================
run_cchange <- function(
  growth_table,
  disturbance_schedule = NULL,
  AVERX = c(0, 0, 0, 0, 0, 0, 0, 0),
  DMINI = 0, FLOSS = 0,
  B31 = -0.03, B32 = 1.5, S = 25,
  CONS = 0.04, SX3 = 0.5, RADIA = 15,
  IROT = 1, YESN = 0,
  soil_c = 5.57, soil_n = 0.296,
  soil_organic_p = 333, soil_bray2_p = 0,
  MATEMP = 12,
  DENSREG = 5,
  floor_init = c(0, 0, 0, 0, 0)
) {
  # growth_table must have columns: Age, SPHA, MTH, Vol, GrossVol, BA, WholeStemDens, RingDens
  required <- c("Age", "SPHA", "MTH", "BA")
  miss <- setdiff(required, names(growth_table))
  if (length(miss)) stop(paste("Missing growth_table columns:", paste(miss, collapse = ", ")))

  # Defaults for optional columns
  if (!"Vol" %in% names(growth_table)) growth_table$Vol <- 0
  if (!"GrossVol" %in% names(growth_table)) growth_table$GrossVol <- 0
  if (!"WholeStemDens" %in% names(growth_table)) growth_table$WholeStemDens <- 0
  if (!"RingDens" %in% names(growth_table)) growth_table$RingDens <- 0

  NYRSD <- nrow(growth_table)
  HTSP <- rep(0, NYRSD + 10)
  SPHSP <- rep(0, NYRSD + 10)
  BASP <- rep(0, NYRSD + 10)
  TVOLSP <- rep(0, NYRSD + 10)
  DENWH <- rep(0, NYRSD + 10)
  DENINC <- rep(0, NYRSD + 10)
  F128SP <- rep(0, NYRSD + 10)

  for (i in seq_len(NYRSD)) {
    idx <- as.integer(growth_table$Age[i]) + 1
    if (idx > 0 && idx <= length(HTSP)) {
      HTSP[idx] <- growth_table$MTH[i]
      SPHSP[idx] <- growth_table$SPHA[i]
      BASP[idx] <- growth_table$BA[i]
      TVOLSP[idx] <- growth_table$Vol[i]
      DENWH[idx] <- growth_table$WholeStemDens[i]
      DENINC[idx] <- growth_table$RingDens[i]
    }
  }

  # Compute mortality volume (gross - net)
  for (i in 2:NYRSD) {
    idx <- as.integer(growth_table$Age[i]) + 1
    if (idx > 1 && growth_table$GrossVol[i] > 0) {
      F128SP[idx] <- max(0, growth_table$GrossVol[i] - growth_table$Vol[i])
    }
  }

  # Set up disturbance arrays
  NDIST <- 0
  ADIST <- rep(10000, 11)
  SPHAN <- rep(-1, 10)
  BAN <- rep(-1, 10)
  PRNHT <- rep(-1, 10)
  XTRACT <- rep(0, 10)
  XTRACC <- rep(0, 10)
  XTRACF <- rep(0, 10)

  if (!is.null(disturbance_schedule) && nrow(disturbance_schedule) > 0) {
    NDIST <- nrow(disturbance_schedule)
    for (i in seq_len(NDIST)) {
      ADIST[i] <- disturbance_schedule$Age[i]
      if ("SPHA" %in% names(disturbance_schedule)) SPHAN[i] <- disturbance_schedule$SPHA[i]
      if ("BA" %in% names(disturbance_schedule)) BAN[i] <- disturbance_schedule$BA[i]
      if ("PruneHt" %in% names(disturbance_schedule)) PRNHT[i] <- disturbance_schedule$PruneHt[i]
      if ("StemExtract" %in% names(disturbance_schedule)) XTRACT[i] <- disturbance_schedule$StemExtract[i]
      if ("CrownExtract" %in% names(disturbance_schedule)) XTRACC[i] <- disturbance_schedule$CrownExtract[i]
      if ("FloorExtract" %in% names(disturbance_schedule)) XTRACF[i] <- disturbance_schedule$FloorExtract[i]
    }
  }

  if (NDIST == 0) {
    NDIST <- 1
    ADIST[1] <- 0   # VBA starts at age 0 (planting year)
    SPHAN[1] <- growth_table$SPHA[1]
    BAN[1] <- growth_table$BA[1]
  }
  ADIST[NDIST + 1] <- 10000

  # Constants
  AM1 <- 10.5
  AM2 <- -0.0301704
  SFINT <- 0.000725
  SFSPE <- 0.156
  BM1 <- 181.04
  BM2 <- 0.3
  AY1A <- 0.009617
  AY1B <- -0.08771
  AY1C <- 1.7122
  INPNAP <- 1
  INDEAD <- 0

  AGEND <- max(growth_table$Age)

  # Initialise compartments
  init <- cc_distur_init(
    SPHA_init = SPHAN[1], AVERX = AVERX,
    DMINI = DMINI, FLOSS = FLOSS,
    B31 = B31, B32 = B32, S = S,
    CONS = CONS, SX3 = SX3, RADIA = RADIA,
    IROT = IROT, floor_init = floor_init,
    YESN = YESN, soil_c = soil_c, soil_n = soil_n,
    soil_organic_p = soil_organic_p
  )

  X <- init$X
  XN <- init$XN
  XP <- init$XP

  t_start <- ADIST[1]
  T <- t_start
  ID <- 1
  SPHA <- SPHAN[1]
  BA <- if (BAN[1] > 0) BAN[1] else 0
  HT <- 0.2
  GCL <- HT
  HTBGC <- 0
  HINC <- 0
  GCFRE <- HT
  TBRCH <- T
  GCLMAX <- 0
  X3MAX <- AM1 * exp(AM2 * (T + 1))
  X6MAX <- if (SPHA > 0) BM1 * (1 / SPHA)^BM2 else 100

  Vol <- 0
  dens <- 0

  # Output storage
  out_list <- list()

  # Initial output
  out0 <- cc_output(X, T, init$DMLOSI, init$DMLOSD, init$DMINI, 1)
  out_list[[1]] <- data.frame(
    Age = T, `Needle_0to1yr(X3)` = X[3], `Needle_1to2yr(X4)` = X[4], `Needle_2plus_yr(X5)` = X[5], `Live_branch(X6)` = X[6], `Dead_branch(X7)` = X[7],
    `Stem_wood(X8)` = X[8], `Coarse_root(X9)` = X[9], `Needle_litter(X10)` = X[10], `Branch_litter(X11)` = X[11], `Stem_litter(X12)` = X[12],
    `Coarse_root_litter(X13)` = X[13], `Live_fine_root(X14)` = X[14], `Fine_root_litter(X15)` = X[15], `Stem_bark(X16)` = X[16],
    SPHA = SPHA, BA = BA, HT = HT, GCL = GCL, Vol = Vol, dens = dens,
    CTREES = out0$CTREES, CSHRUB = out0$CSHRUB, CSTAND = out0$CSTAND,
    CFAS = out0$CFAS, CSTEM = out0$CSTEM,
    C_branch_live = out0$C_branch_live, C_branch_dead = out0$C_branch_dead,
    CROOTL = out0$CROOTL, CROOTD = out0$CROOTD,
    C_needle_litter = out0$C_needle_litter, C_branch_litter = out0$C_branch_litter,
    C_stem_litter = out0$C_stem_litter,
    stringsAsFactors = FALSE, check.names = FALSE
  )

  # Annual loop
  while (T < AGEND) {
    j <- as.integer(T)

    # Compute NAP from growth table volume increments and density
    ISPT <- as.integer(T) + 1
    NAP <- 0
    if (ISPT > 0 && ISPT < length(TVOLSP) && TVOLSP[ISPT] > 0) {
      vol_now <- TVOLSP[min(ISPT + 1, length(TVOLSP))]
      vol_prev <- TVOLSP[ISPT]
      dens_now <- if (DENSREG == 6 && DENINC[ISPT] > 0) DENINC[ISPT] else
                  if (DENWH[ISPT] > 0) DENWH[ISPT] else 0.42
      stem_inc <- max(0, (vol_now - vol_prev)) * dens_now
      NAP <- stem_inc / 0.4
    }

    # Generate flow matrix
    F128SP_val <- if (j + 2 <= length(F128SP)) F128SP[j + 2] else 0
    fmat <- cc_dmfmat(
      X = X, NAP = NAP, t = T,
      SPHA = SPHA, BA = BA, HT = HT, GCL = GCL, HTBGC = HTBGC,
      CONS = CONS, SX3 = SX3, INDEAD = INDEAD, INPNAP = INPNAP,
      F128SP_val = F128SP_val,
      X6MAX = X6MAX, AM1 = AM1, AM2 = AM2, BM1 = BM1, BM2 = BM2,
      AY1A = AY1A, AY1B = AY1B, AY1C = AY1C,
      MATEMP = MATEMP
    )

    # Compute rate matrix
    A <- cc_dmamat(fmat$F, X, fmat$Y)

    # Store pre-integration state for nutrients
    XMN1 <- X
    XNMN1 <- XN
    XPMN1 <- XP

    # Euler integration
    X8HOLD <- X[8]
    X <- cc_euler(A, X, NAP)
    X8GAIN <- X[8] - X8HOLD

    # Density and structural update
    ATD <- T + 1
    if (DENSREG <= 3) {
      DRA <- -370; DRB <- -0.035; DRC <- 510
      dens <- (DRC + DRA * (1 - exp(DRB * ATD))) / 1000
    }
    IATD <- as.integer(ATD) + 1
    if (DENSREG == 5 && IATD <= length(DENWH) && DENWH[IATD] > 0) {
      dens <- DENWH[IATD]
    }

    # Structural update
    str_out <- cc_struct(
      X, T, SPHA, BA, HT, GCL, HTBGC, HINC,
      B31, B32, S, SX3, AM1, AM2, BM1, BM2,
      AY1A, AY1B, AY1C, SFINT, SFSPE,
      HTSP, BASP, SPHSP, TVOLSP,
      ADIST, ID, DENSREG, DENWH, DENINC,
      X8GAIN, INPNAP
    )

    T <- str_out$t
    HT <- str_out$HT
    GCL <- str_out$GCL
    HTBGC <- str_out$HTBGC
    HINC <- str_out$HINC
    BA <- str_out$BA
    SPHA <- str_out$SPHA
    X3MAX <- str_out$X3MAX
    X6MAX <- str_out$X6MAX
    Vol <- str_out$Vol
    if (str_out$dens > 0) dens <- str_out$dens

    # Nutrient simulation
    if (YESN == 1) {
      nut <- cc_nutrient(
        fmat$F, X, XMN1, XN, XP, XNMN1, XPMN1, T,
        fmat$PROP_NEEDLE, fmat$PROP_BRANCH, fmat$PROP_STEM,
        fmat$PROP_CROOT, fmat$PROP_ABRANCH,
        soil_c, soil_n, soil_organic_p
      )
      XN <- nut$XN
      XP <- nut$XP
    }

    # Carbon output
    cout <- cc_output(X, T, init$DMLOSI, init$DMLOSD, init$DMINI, 1)

    out_list[[length(out_list) + 1]] <- data.frame(
      Age = T, `Needle_0to1yr(X3)` = X[3], `Needle_1to2yr(X4)` = X[4], `Needle_2plus_yr(X5)` = X[5], `Live_branch(X6)` = X[6], `Dead_branch(X7)` = X[7],
      `Stem_wood(X8)` = X[8], `Coarse_root(X9)` = X[9], `Needle_litter(X10)` = X[10], `Branch_litter(X11)` = X[11], `Stem_litter(X12)` = X[12],
      `Coarse_root_litter(X13)` = X[13], `Live_fine_root(X14)` = X[14], `Fine_root_litter(X15)` = X[15], `Stem_bark(X16)` = X[16],
      SPHA = SPHA, BA = BA, HT = HT, GCL = GCL, Vol = Vol, dens = dens,
      CTREES = cout$CTREES, CSHRUB = cout$CSHRUB, CSTAND = cout$CSTAND,
      CFAS = cout$CFAS, CSTEM = cout$CSTEM,
      C_branch_live = cout$C_branch_live, C_branch_dead = cout$C_branch_dead,
      CROOTL = cout$CROOTL, CROOTD = cout$CROOTD,
      C_needle_litter = cout$C_needle_litter, C_branch_litter = cout$C_branch_litter,
      C_stem_litter = cout$C_stem_litter,
      stringsAsFactors = FALSE, check.names = FALSE
    )

    # Check for disturbance
    if (T >= ADIST[ID] && ID <= NDIST) {
      thin_result <- cc_distur_thin(
        X, XN, XP, BA, SPHA, BAN[ID], XTRACT[ID],
        XTRACC[ID], XTRACF[ID], YESN, T,
        soil_c, soil_n, soil_organic_p
      )
      X <- thin_result$X
      XN <- thin_result$XN
      XP <- thin_result$XP
      BA <- thin_result$BA
      ID <- ID + 1
    }

    if (T >= AGEND) break
  }

  result_df <- do.call(rbind, out_list)
  rownames(result_df) <- NULL

  list(
    annual_carbon = result_df,
    carbon_summary = data.frame(
      FinalAge = max(result_df$Age),
      FinalStandCarbon_tC_ha = tail(result_df$CSTAND, 1),
      PeakStandCarbon_tC_ha = max(result_df$CSTAND, na.rm = TRUE),
      FinalTreeCarbon_tC_ha = tail(result_df$CTREES, 1),
      FinalStemCarbon_tC_ha = tail(result_df$CSTEM, 1),
      FinalFoliageCarbon_tC_ha = tail(result_df$CFAS, 1),
      FinalLiveRootCarbon_tC_ha = tail(result_df$CROOTL, 1),
      stringsAsFactors = FALSE
    )
  )
}
