# ---------------------------------------------------------------------------
# MultiSpecies_Growth.R
# Port of VBA Module 2 (MultiSpecies_GrowthModel.bas) functions that are
# called by run_model() in TreeLevel_Input.R but were not yet implemented.
#
# These functions complete the multi-species growth model for all 11+ species.
# ---------------------------------------------------------------------------

# Breast height constant
BH_CONST <- 1.4

# ===========================
# Core growth model functions
# ===========================

#' Stand-level volume function
#' @param MTH Mean top height (m)
#' @param DBH Quadratic mean DBH (cm)
#' @param N Stocking (stems/ha)
#' @param VOL_type Volume function type (1, 2, or 3)
#' @param VOL_u,VOL_v,VOL_w,VOL_z Volume function coefficients
Vol_stand <- function(MTH, DBH, N, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z) {
  BA <- N * pi * (DBH / 200)^2
  if (DBH <= 0 || MTH <= 1.4) {
    return(0)
  } else if (VOL_type == 1) {
    return(MTH * BA * (VOL_v * (MTH - 1.4)^(-VOL_w) + VOL_u))
  } else if (VOL_type == 2) {
    return(exp(VOL_u + VOL_v * log(BA) + VOL_w * log(MTH) + VOL_z * (log(MTH))^2))
  } else if (VOL_type == 3) {
    return(BA * (VOL_v + VOL_u * MTH))
  }
  return(0)
}

#' Mortality calculation using SDI-based model
#' @param N0 Current stocking
#' @param DBH Current DBH
#' @param deltaT Time step
#' @param MORT_k,MORT_m,MORT_n Mortality model coefficients
N_Mort <- function(N0, DBH, deltaT, MORT_k, MORT_m, MORT_n) {
  if (DBH <= 0) {
    mort <- 100 * MORT_k
  } else {
    sdi <- (0.405 * N0 * (0.394 * DBH / 10)^1.605) / 1000
    mort <- 100 * (MORT_k + MORT_m * sdi^MORT_n)
  }
  return(N0 * (1 - mort / 100)^deltaT)
}

#' Predict MTH from age T and Site Index SI
MTH_mod <- function(T, SI, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) {
  return(0.3 + Y(T, MTH_model, MTH_form, SI - 0.3, 30, MTH_a, MTH_b, MTH_c))
}

#' Predict SI from MTH at age T
SI_eqn <- function(T, MTH, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) {
  return(0.3 + Y0_fn(30, MTH_model, MTH_form, MTH - 0.3, T, MTH_a, MTH_b, MTH_c))
}

#' Determine breast height age
AgeBH <- function(T, H, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) {
  return(TP_fn(1.4 - 0.3, MTH_model, MTH_form, H - 0.3, T, MTH_a, MTH_b, MTH_c))
}

#' Estimate mean height from MTH and stocking
MnHt_from_MTH <- function(MTH, N, MTH_MnHt_a, MTH_MnHt_b) {
  return(MTH * (1 - MTH_MnHt_a * (1 - exp(MTH_MnHt_b * (N - 100)))))
}

#' Crown height calculation (redwood)
CrownHeight <- function(Meanht, N, pCH = -39.1, qch = 0.78, rch = 4.87) {
  ch <- pCH + qch * Meanht + rch * log(N)
  if (ch < 0) ch <- 0
  return(ch)
}

#' Calculate D300_30 from 300 Index and SI
D300_30_from_I300_SI <- function(I300, SI, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z) {
  if (VOL_type == 1) {
    BA <- (I300 * 30) / (SI * (VOL_v * (SI - 1.4)^(-VOL_w) + VOL_u))
  } else {
    BA <- exp((log(I300 * 30) - VOL_u - VOL_w * log(SI) - VOL_z * (log(SI))^2) / VOL_v)
  }
  return(200 * sqrt(BA / 300 / pi))
}

#' Calculate 300 Index from D300_30 and SI
I300_from_SI_D300_30 <- function(SI, D300_30, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z) {
  return(Vol_stand(SI, D300_30, 300, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z) / 30)
}

# ===========================
# Inverse growth model (Y0) and time prediction (TP)
# ===========================

#' Inverse growth model: given Y at age T, predict Y0 at age t0
Y0_fn <- function(t0, model, form, Y_val, T, A, B, C) {
  if (model == "Richards") {
    if (form == "Anamorphic") {
      return(Y_val * ((1 - exp(-B * T)) / (1 - exp(-B * t0)))^(-C))
    } else if (form == "CA") {
      return(A * (1 - (1 - (Y_val / A)^(1 / C))^(t0 / T))^C)
    } else if (form == "GADA") {
      R <- (log(Y_val) - C * log(1 - exp(-B * T))) / (1 + A * log(1 - exp(-B * T)))
      return(exp(R) * (1 - exp(-B * t0))^(A * R + C))
    } else {
      return(0)
    }
  } else if (model == "Korf") {
    if (form == "Anamorphic") {
      return(Y_val * exp(-B * (t0)^(-C)) / exp(-B * (T)^(-C)))
    } else if (form == "CA") {
      return(A * (Y_val / A)^((T / t0)^C))
    } else if (form == "GADA") {
      R <- log(Y_val) + sqrt((log(Y_val))^2 + 4 * B * (T)^(-C))
      return(exp(R / 2 - 2 * B * (t0)^(-C) / R))
    } else {
      return(0)
    }
  } else if (model == "Hossfeld") {
    if (form == "Anamorphic") {
      return(1 / (1 / Y_val + B * (1 / (t0)^C - 1 / (T)^C)))
    } else if (form == "CA") {
      return((t0^C) / ((T^C) / Y_val - (T^C - t0^C) / A))
    } else if (form == "GADA") {
      R <- (Y_val - A + sqrt((Y_val - A)^2 + 4 * Y_val * B * (T)^(-C))) / 2
      return(R * (A + R) / (R + B * (t0)^(-C)))
    } else {
      return(0)
    }
  } else {
    return(0)
  }
}

#' Time prediction: given Y1 at unknown age, Y0 at age t0, predict age
TP_fn <- function(Y1, model, form, Y0_val, t0, A, B, C) {
  if (model == "Richards") {
    if (form == "Anamorphic") {
      return(-(1 / B) * log(1 - (1 - exp(-B * t0)) * (Y1 / Y0_val)^(1 / C)))
    } else if (form == "CA") {
      return(t0 * log(1 - (Y1 / A)^(1 / C)) / log(1 - (Y0_val / A)^(1 / C)))
    } else if (form == "GADA") {
      R0 <- (log(Y0_val) - log((1 - exp(-B * t0))^C)) / (1 + log((1 - exp(-B * t0))^A))
      return(-(1 / B) * log(1 - exp((log(Y1) - R0) / (C + A * R0))))
    } else {
      return(0)
    }
  } else if (model == "Korf") {
    if (form == "Anamorphic") {
      return((-1 / B * log(Y1 / Y0_val) + (t0)^(-C))^(-1 / C))
    } else if (form == "CA") {
      return(t0 * ((log(Y1 / A)) / (log(Y0_val / A)))^(-1 / C))
    } else if (form == "GADA") {
      R0 <- (log(Y0_val) + sqrt((log(Y0_val))^2 + 4 * B * (t0)^(-C)))
      return((-2 * B / (R0 * log(Y1) - ((R0)^2) / 2))^(1 / C))
    } else {
      return(0)
    }
  } else if (model == "Hossfeld") {
    if (form == "Anamorphic") {
      return((1 / B * (1 / Y1 - 1 / Y0_val + B / ((t0)^C)))^(-1 / C))
    } else if (form == "CA") {
      return(t0 * ((1 / Y0_val - 1 / A) / (1 / Y1 - 1 / A))^(1 / C))
    } else if (form == "GADA") {
      R0 <- (Y0_val - A + sqrt((Y0_val - A)^2 + 4 * Y0_val * B * (t0)^(-C))) / 2
      return((R0 / B * ((A + R0) / Y1 - 1))^(-1 / C))
    } else {
      return(0)
    }
  }
  return(0)
}

#' Estimate b-parameter of Korf MTH model from two MTH measurements (bisection)
MTHmodel_b <- function(age1, MTH1, Age2, MTH2) {
  blo <- 10
  bup <- 100
  for (i in 1:16) {
    bmid <- (blo + bup) / 2
    pred_MTH2lo <- 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, blo, MTH_c)
    pred_MTH2up <- 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, bup, MTH_c)
    pred_MTH2mid <- 0.3 + Y(Age2, MTH_model, MTH_form, MTH1 - 0.3, age1, MTH_a, bmid, MTH_c)
    flo <- pred_MTH2lo - MTH2
    fup <- pred_MTH2up - MTH2
    fmid <- pred_MTH2mid - MTH2
    if (fmid * flo < 0) bup <- bmid else blo <- bmid
  }
  return(bmid)
}

# ===========================
# Thinning age shift (bisection)
# ===========================

#' Estimate initial thinning age-shift using bisection
Thin_age_shift <- function(Age, N1, N2, Thin_coeff, Age_BH, D300_30, SI30) {
  DBH_pre_thin <- DBH_mod(Age, D300_30, SI30, Age_BH, N1,
                           DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                           DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                           MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
  DBH_post_thin <- DBH_pre_thin * ((N1 / N2)^((1 - Thin_coeff) / 2))

  Agelo <- Age_BH + 1
  Ageup <- Age - Age_BH + 20
  for (i in 1:16) {
    Agemid <- (Agelo + Ageup) / 2
    flo <- DBH_mod(Agelo, D300_30, SI30, Age_BH, N2,
                   DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                   DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                   MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
    fup <- DBH_mod(Ageup, D300_30, SI30, Age_BH, N2,
                   DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                   DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                   MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
    fmid <- DBH_mod(Agemid, D300_30, SI30, Age_BH, N2,
                    DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                    DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                    MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) - DBH_post_thin
    if (flo * fmid < 0) Ageup <- Agemid else Agelo <- Agemid
  }
  return(Age - Agemid)
}

# ===========================
# Tree list functions
# ===========================

#' Scale tree list to match stand-level DBH
Scale_tree_list <- function(Treelist, nstems, stocking, DBH) {
  sum_sqd <- 0
  for (tree in 1:nstems) {
    sum_sqd <- sum_sqd + Treelist[tree, 3]^2
  }
  DBH_scaling_factor <- DBH / sqrt(sum_sqd / nstems)
  for (tree in 1:nstems) {
    Treelist[tree, 2] <- stocking / nstems
    Treelist[tree, 3] <- Treelist[tree, 3] * DBH_scaling_factor
  }
  return(Treelist)
}

#' Predict heights for tree list using Petterson Type 1 curve
Predict_height <- function(Treelist, nstems, StandDensity, MnDBH, MTH_val, PRUNEHT_val) {
  for (tree in 1:nstems) {
    Treelist[tree, 5] <- 1.4 + (petA + petB / Treelist[tree, 3])^(-2.5)
    Treelist[tree, 6] <- vub(Species, Treelist[tree, 3], Treelist[tree, 5], 0,
                             StandDensity, MnDBH, MTH_val, PRUNEHT_val)
  }
  return(Treelist)
}

#' Process felled stems
Felled_stems <- function(Species_name, Treelist, nstems, Age, stocking, DBH,
                         Vol_val, StandDensity, MnDBH, MTH_val, PRUNEHT_val,
                         Stemlist, stemno) {
  Treelist <- Scale_tree_list(Treelist, nstems, stocking, DBH)
  stem_vol <- 0
  for (i in 1:nstems) {
    stem_vol <- stem_vol + vub(Species_name, Treelist[i, 3], Treelist[i, 5], 0,
                               StandDensity, MnDBH, MTH_val, PRUNEHT_val)
  }
  stem_vol <- stem_vol * stocking / nstems

  for (i in 1:nstems) {
    Stemlist[stemno, 1] <- Age
    Stemlist[stemno, 2] <- i
    Stemlist[stemno, 3] <- stocking / nstems
    Stemlist[stemno, 4] <- Treelist[i, 3]
    Stemlist[stemno, 5] <- Treelist[i, 5]
    Stemlist[stemno, 6] <- vub(Species_name, Treelist[i, 3], Treelist[i, 5], 0,
                               StandDensity, MnDBH, MTH_val, PRUNEHT_val) * Vol_val / stem_vol
    stemno <- stemno + 1
  }
  return(list(Stemlist = Stemlist, stemno = stemno))
}

#' Make logs from felled stems
Make_Logs <- function(Species_name, Treelist, nstems, Age, stocking, DBH,
                      Vol_val, StandDensity, MnDBH, MTH_val, PRUNEHT_val,
                      Logs, logno, log_length, min_SED, break_height, log_losses) {
  Treelist <- Scale_tree_list(Treelist, nstems, stocking, DBH)
  Harvest_volume <- 0
  stem_vol <- 0
  for (i in 1:nstems) {
    stem_vol <- stem_vol + vub(Species_name, Treelist[i, 3], Treelist[i, 5], 0,
                               StandDensity, MnDBH, MTH_val, PRUNEHT_val)
  }
  stem_vol <- stem_vol * stocking / nstems

  for (i in 1:nstems) {
    for (j in 1:100) {
      log_top <- 0.3 + log_length * j
      if (log_top >= Treelist[i, 5]) break
      sed <- dub(Species_name, Treelist[i, 3], Treelist[i, 5], log_top,
                 StandDensity, MnDBH, MTH_val, PRUNEHT_val)
      if (sed < min_SED / 10) break
      if (log_top >= break_height * Treelist[i, 5]) break
      Volume <- vol_ub(Species_name, Treelist[i, 3], Treelist[i, 5],
                       0.3 + log_length * (j - 1), log_top,
                       StandDensity, MnDBH, MTH_val, PRUNEHT_val) * Vol_val / stem_vol
      Logs[logno, 1] <- Age
      Logs[logno, 2] <- stocking / nstems * (100 - log_losses) / 100
      Logs[logno, 3] <- i
      Logs[logno, 4] <- j
      Logs[logno, 5] <- sed * 10
      Logs[logno, 6] <- 10 * dub(Species_name, Treelist[i, 3], Treelist[i, 5],
                                  0.3 + log_length * (j - 1),
                                  StandDensity, MnDBH, MTH_val, PRUNEHT_val)
      Logs[logno, 7] <- log_length
      Logs[logno, 8] <- Volume
      Harvest_volume <- Harvest_volume + Volume * Logs[logno, 2]
      logno <- logno + 1
    }
  }
  return(list(Logs = Logs, logno = logno, Harvest_volume = Harvest_volume))
}

#' Summarize harvest data
harvest_summary_fn <- function(Logs, logno) {
  harvest_sum <- matrix(0, nrow = 30, ncol = 12)
  if (logno <= 1) return(harvest_sum)

  last_log_age <- Logs[1, 1]
  row <- 1
  for (i in 1:(logno - 1)) {
    log_age <- Logs[i, 1]
    if (log_age != last_log_age) {
      row <- row + 4
    }
    last_log_age <- log_age
    log_size <- as.integer(Logs[i, 5] / 100)
    if (log_size > 10) log_size <- 10
    log_size <- log_size + 1  # 1-indexed column
    log_height <- Logs[i, 4]
    if (log_height > 3) log_height <- 3

    harvest_sum[row, 1] <- log_age
    harvest_sum[row + log_height - 1, log_size] <- harvest_sum[row + log_height - 1, log_size] + Logs[i, 8] * Logs[i, 2]
    harvest_sum[row + 3, log_size] <- harvest_sum[row + 3, log_size] + Logs[i, 8] * Logs[i, 2]
    harvest_sum[row + log_height - 1, 12] <- harvest_sum[row + log_height - 1, 12] + Logs[i, 8] * Logs[i, 2]
    harvest_sum[row + 3, 12] <- harvest_sum[row + 3, 12] + Logs[i, 8] * Logs[i, 2]
  }
  return(harvest_sum)
}

#' Early yield correction for multi-species
#' Corrects early volume predictions when DBH < 2 cm
ms_earlyield <- function(N_vec, DBH_vec, Vol_vec, rotlength) {
  initialvol <- 0.0000064
  ddbh <- 0
  for (Age in 1:min(15, rotlength)) {
    prevdbh <- ddbh
    ddbh <- DBH_vec[Age + 1]  # 1-indexed, Age=0 is index 1
    if (ddbh >= 2 && prevdbh < 2) {
      treevolinc <- Vol_vec[Age + 1] / N_vec[Age + 1] - initialvol
      k <- treevolinc / (Age^2.7)
      for (j in seq(Age - 1, 0, by = -1)) {
        TT <- max(j, 0)
        Vol_vec[j + 1] <- (initialvol + k * TT^2.7) * N_vec[j + 1]
      }
    }
  }
  return(Vol_vec)
}

# ===========================
# Volume under bark functions (VUB)
# ===========================

#' Redwood volume under bark above height H
redwood_vub <- function(DBH, HT, H) {
  a0 <- 0.702; a1 <- 0.5646; a2 <- -0.6188
  b1 <- 2.6295; b2 <- 0.1406; b3 <- 0.1455; b4 <- -0.1275; b5 <- 22.7873
  bh <- 1.4

  if (DBH == 0 || HT == 0) return(0)

  if (H > HT) H <- HT
  H_len <- HT - H
  if (H_len / HT < 0.0001) H_len <- 0

  hp2 <- HT^b2
  dhp3 <- (DBH * HT)^b4
  beta1 <- (1 - (b3 / dhp3) * (1 - bh / HT)^b5) / ((1 - bh / HT)^(b1 / hp2))
  gohp2 <- b1 / hp2

  p1 <- (a0 * beta1 / (HT^gohp2 * (gohp2 + 1))) * H_len^(gohp2 + 1)
  p2 <- (a0 * b3 / (dhp3 * HT^b5 * (b5 + 1))) * H_len^(b5 + 1)
  p3 <- (a1 * beta1 / (HT^(gohp2 + 1) * (gohp2 + 2))) * H_len^(gohp2 + 2)
  p4 <- (a1 * b3 / (dhp3 * HT^(b5 + 1) * (b5 + 2))) * H_len^(b5 + 2)
  p5 <- (a2 * beta1 / (HT^(gohp2 + 2) * (gohp2 + 3))) * H_len^(gohp2 + 3)
  p6 <- (a2 * b3 / (dhp3 * HT^(b5 + 2) * (b5 + 3))) * H_len^(b5 + 3)

  return((pi / 40000) * DBH * DBH * (p1 + p2 + p3 + p4 + p5 + p6))
}

#' Douglas-fir volume under bark above height H
Dfir_Vub <- function(DBH, HT, H) {
  a1 <- 0.319071; a2 <- 0; a3 <- 23.9972; a4 <- -47.47884; a5 <- 26.02156
  v1 <- 1.8281198; v2 <- 1.102592; v3 <- -10.19719
  V <- DBH^v1 * (HT^2 / (HT - 1.4))^v2 * exp(v3)
  c1 <- a1 / 2; c2 <- a2 / 3; c3 <- a3 / 4; c4 <- a4 / 5; c5 <- a5 / 6
  return(V * (c1 * ((HT - H) / HT)^2 + c2 * ((HT - H) / HT)^3 +
              c3 * ((HT - H) / HT)^4 + c4 * ((HT - H) / HT)^5 +
              c5 * ((HT - H) / HT)^6))
}

#' Radiata pine partial tree volume (Gordon and Budianto)
Radiata_VUB <- function(DBH, HT, H, SPH, MnDBH, MTH_val, PRHT) {
  a0 <- 0.4242; a01 <- -0.002822; a10 <- 0.6067; a12 <- 0.06129
  a2 <- -0.207; a31 <- 0.3208
  bf0 <- 0.945; bf1 <- -0.387; bf2 <- 0.000686; bf3 <- -0.267; bf4 <- 0.00357
  b30 <- 0.7768; B31 <- -0.1347
  g10 <- 1.018; g11 <- 0.2967; g2 <- 12.68; g31 <- 1.047

  rspace <- 100 / (sqrt(SPH) * MTH_val)
  sd <- MnDBH^2 / sqrt(rspace)
  FQ <- bf0 + bf1 * exp(-bf2 * sd) + bf3 * exp(-(HT / MTH_val)^2) + bf4 * PRHT
  D6 <- FQ * DBH
  l <- HT - H
  z <- l / HT
  zb <- 1 - 1.4 / HT
  zu <- 1 - 6 / HT
  g1 <- g10 + g11 * D6 / (HT - 6)
  g3 <- g31 * HT * D6 / DBH
  b3 <- b30 + B31 * (DBH - D6) / (6 - 1.4)
  b1 <- (1 - (zb^g2 / zu^g2) * (D6^2 / DBH^2 - b3 * zu^g3) - b3 * zb^g3) /
        (zb^g1 - (zb^g2 * zu^g1) / zu^g2)
  b2 <- (D6^2 / DBH^2 - b1 * zu^g1 - b3 * zu^g3) / zu^g2

  K1 <- (a0 + a01 * HT)
  K2 <- 1 + 0.5 / exp(a12 * HT)
  result <- (pi * DBH^2 * HT / 40000) *
    ((l / HT)^g1 * ((b1 * K1 * l) / ((1 + g1) * HT) +
                     (a10 * b1 * (l / HT)^K2) / (K2 + g1)) +
     (l / HT)^g2 * ((b2 * K1 * l) / ((1 + g2) * HT) +
                     (a10 * b2 * (l / HT)^K2) / (K2 + g2)) +
     (l / HT)^g3 * ((b3 * K1 * l) / ((1 + g3) * HT) +
                     (a10 * b3 * (l / HT)^K2) / (K2 + g3)) +
     (l / HT)^(a31 * HT) *
       ((a2 * b1 * (l / HT)^(1 + g1)) / (1 + g1 + a31 * HT) +
        (a2 * b2 * (l / HT)^(1 + g2)) / (1 + g2 + a31 * HT) +
        (a2 * b3 * (l / HT)^(1 + g3)) / (1 + g3 + a31 * HT)))
  return(result)
}

#' Volume under bark dispatcher (currently uses redwood for all species,
#' matching VBA Module 2 which has species-specific routing commented out)
vub <- function(Species_name, DBH, HT, H, N, MnDBH, MTH_val, PRUNEHT_val) {
  return(redwood_vub(DBH, HT, H))
}

#' Volume under bark between two levels on stem
vol_ub <- function(Species_name, DBH, HT, lower, upper, MnDBH, MTH_val, N, PRUNEHT_val) {
  return(vub(Species_name, DBH, HT, lower, N, MnDBH, MTH_val, PRUNEHT_val) -
         vub(Species_name, DBH, HT, upper, N, MnDBH, MTH_val, PRUNEHT_val))
}

# ===========================
# Taper (diameter under bark) functions
# ===========================

#' Redwood diameter under bark at height level
Redwood_dub <- function(DBH, HT, level) {
  a0 <- 0.702; a1 <- 0.5646; a2 <- -0.6188
  b1 <- 2.6295; b2 <- 0.1406; b3 <- 0.1455; b4 <- -0.1275; b5 <- 22.7873
  bh <- 1.4

  z <- 1 - level / HT
  if (level < 0) z <- 1
  if (z < 0.0001) z <- 0
  beta_c <- (1 - b3 / ((DBH * HT)^b4) * ((1 - bh / HT)^b5)) /
            ((1 - bh / HT)^(b1 / (HT^b2)))
  dob <- sqrt(DBH * DBH * (beta_c * z^(b1 / (HT^b2)) + (b3 / ((DBH * HT)^b4)) * (z^b5)))
  return(dob * sqrt(a0 + a1 * z + a2 * z^2))
}

#' Douglas-fir diameter under bark (T136 taper)
Dfir_Dub <- function(DBH, HT, H) {
  a1 <- 0.319071; a2 <- 0; a3 <- 23.9972; a4 <- -47.47884; a5 <- 26.02156
  v1 <- 1.8281198; v2 <- 1.102592; v3 <- -10.19719
  V <- DBH^v1 * (HT^2 / (HT - 1.4))^v2 * exp(v3)
  D2 <- 40000 * V / (pi * HT) * (a1 * ((HT - H) / HT) + a2 * ((HT - H) / HT)^2 +
        a3 * ((HT - H) / HT)^3 + a4 * ((HT - H) / HT)^4 + a5 * ((HT - H) / HT)^5)
  return(sqrt(D2))
}

#' Radiata pine diameter under bark (three-point taper, Gordon & Budianto)
Radiata_Dub <- function(DBH, HT, H, SPH, MnDBH, MTH_val, PRHT) {
  a0 <- 0.4242; a01 <- -0.002822; a10 <- 0.6067; a12 <- 0.06129
  a2 <- -0.207; a31 <- 0.3208
  bf0 <- 0.945; bf1 <- -0.387; bf2 <- 0.000686; bf3 <- -0.267; bf4 <- 0.00357
  b30 <- 0.7768; B31 <- -0.1347
  g10 <- 1.018; g11 <- 0.2967; g2 <- 12.68; g31 <- 1.047

  rspace <- 100 / (sqrt(SPH) * MTH_val)
  sd <- MnDBH^2 / sqrt(rspace)
  FQ <- bf0 + bf1 * exp(-bf2 * sd) + bf3 * exp(-(HT / MTH_val)^2) + bf4 * PRHT
  D6 <- FQ * DBH
  l <- HT - H
  z <- l / HT
  zb <- 1 - 1.4 / HT
  zu <- 1 - 6 / HT
  g1 <- g10 + g11 * D6 / (HT - 6)
  g3 <- g31 * HT * D6 / DBH
  b3_val <- b30 + B31 * (DBH - D6) / (6 - 1.4)
  b1 <- (1 - (zb^g2 / zu^g2) * (D6^2 / DBH^2 - b3_val * zu^g3) - b3_val * zb^g3) /
        (zb^g1 - (zb^g2 * zu^g1) / zu^g2)
  b2 <- (D6^2 / DBH^2 - b1 * zu^g1 - b3_val * zu^g3) / zu^g2

  if (b2 > -2) {
    dob <- (DBH^2 * (b1 * z^g1 + b2 * z^g2 + b3_val * z^g3))^0.5
  } else {
    dob <- 0
  }
  dub_val <- dob * (a0 + a01 * HT + a10 * z^(exp(-a12 * HT) / 2) + a2 * z^(a31 * HT))^0.5
  return(dub_val)
}

#' Diameter under bark dispatcher (currently uses redwood for all species,
#' matching VBA Module 2)
dub <- function(Species_name, DBH, HT, H, MnDBH, MTH_val, N, PRUNEHT_val) {
  return(Redwood_dub(DBH, HT, H))
}

# ===========================
# Yield table functions
# ===========================

#' Generic yield table for non-radiata/non-dfir species
Yield_Table <- function() {
  N_vec <- numeric(rotlength + 1)
  MTH_vec <- numeric(rotlength + 1)
  DBH_vec <- numeric(rotlength + 1)
  BA_vec <- numeric(rotlength + 1)
  Vol_vec <- numeric(rotlength + 1)
  WoodDen_vec <- numeric(rotlength + 1)
  log_vol_vec <- numeric(rotlength + 1)
  T_adj_vec <- numeric(rotlength + 1)

  N_vec[1] <- Stock_hist_N[1]  # index 1 = age 0

  MTH_vec[1] <- 0.3
  DBH_vec[1] <- 0
  BA_vec[1] <- 0
  Vol_vec[1] <- 0
  WoodDen_vec[1] <- Den_a * WoodDensity_Adjustment
  log_vol_vec[1] <- 0
  Total_thin_age_shift <- 0
  Adj_T <- 0

  Stemlist <- matrix(0, nrow = 10000, ncol = 6)
  Logs_mat <- matrix(0, nrow = 10000, ncol = 10)
  stemno_local <- 1
  logno_local <- 1
  Harvest_volume_local <- 0

  for (T in 1:rotlength) {
    idx <- T + 1  # 1-indexed
    Adj_T <- Adj_T + 1
    N_vec[idx] <- N_Mort(N_vec[idx - 1], DBH_vec[idx - 1], 1, MORT_k, MORT_m, MORT_n)
    MTH_vec[idx] <- MTH_mod(T, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    log_vol_vec[idx] <- 0
    Adj_T <- T - Total_thin_age_shift

    for (thin in 1:4) {
      if (T >= Stock_hist_T[thin + 1] && Stock_hist_T[thin + 1] != 0 && T < Stock_hist_T[thin + 1] + 1) {
        Pre_thin_N[thin] <- N_vec[idx]
        Pre_thin_dbh[thin] <- DBH_mod(Adj_T, D300_30_est, H30, TBH, N_vec[idx],
                                       DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                                       DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                                       MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
        Treelist <<- Scale_tree_list(Treelist, nstems, Pre_thin_N[thin], Pre_thin_dbh[thin])
        MTDia_val <- MTD(Treelist, nstems, Pre_thin_N[thin])
        petB <<- 1.98
        petA_val <- (MTH_vec[idx] - 1.4)^(-1 / 2.5) - petB / MTDia_val
        if (petA_val < 0) {
          petA_val <- 0
          petB <<- MnHt_from_MTH(MTH_vec[idx], N_vec[idx], MTH_MnHt_a, MTH_MnHt_b)
        }
        petA <<- petA_val
        Treelist <<- Predict_height(Treelist, nstems, N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT)
        Pre_thin_vol[thin] <- Vol_stand(MTH_vec[idx], Pre_thin_dbh[thin], Pre_thin_N[thin],
                                         VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)

        Initial_thin_age_shift <- Thin_age_shift(Adj_T, N_vec[idx], Stock_hist_N[thin + 1],
                                                  Stock_hist_thin_coeff[thin + 1], TBH, D300_30_est, H30)
        Total_thin_age_shift <- Total_thin_age_shift + Initial_thin_age_shift
        Adj_T <- T - Total_thin_age_shift

        if (N_vec[idx] > Stock_hist_N[thin + 1]) N_vec[idx] <- Stock_hist_N[thin + 1]
        if (Adj_T - TBH <= 0) {
          DBH_vec[idx] <- 0
        } else {
          DBH_vec[idx] <- DBH_mod(Adj_T, D300_30_est, H30, TBH, N_vec[idx],
                                   DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                                   DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                                   MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
        }
        Vol_vec[idx] <- Vol_stand(MTH_vec[idx], DBH_vec[idx], N_vec[idx],
                                   VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
        if (Pre_thin_N[thin] - N_vec[idx] <= 0 ||
            (Pre_thin_N[thin] * Pre_thin_dbh[thin]^2 - N_vec[idx] * DBH_vec[idx]^2) <= 0) {
          Thin_dbh[thin] <- Pre_thin_dbh[thin]
        } else {
          Thin_dbh[thin] <- sqrt((Pre_thin_N[thin] * Pre_thin_dbh[thin]^2 -
                                  N_vec[idx] * DBH_vec[idx]^2) / (Pre_thin_N[thin] - N_vec[idx]))
        }
        Thin_vol[thin] <- Pre_thin_vol[thin] - Vol_vec[idx]

        if (!Cali && !Minimal_run) {
          fs_result <- Felled_stems(Species, Treelist, nstems, T,
                                    Pre_thin_N[thin] - N_vec[idx], Thin_dbh[thin], Thin_vol[thin],
                                    N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                    Stemlist, stemno_local)
          Stemlist <- fs_result$Stemlist
          stemno_local <- fs_result$stemno
        }
        if (!Cali && !Minimal_run && Stock_hist_Type[thin + 1] == 2) {
          ml_result <- Make_Logs(Species, Treelist, nstems, T,
                                Pre_thin_N[thin] - N_vec[idx], Thin_dbh[thin], Thin_vol[thin],
                                N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
          Logs_mat <- ml_result$Logs
          logno_local <- ml_result$logno
          log_vol_vec[idx] <- ml_result$Harvest_volume
        }

        Additional_shift <- Initial_thin_age_shift * 0.5
        if (Additional_shift > 0.25) Additional_shift <- 0.25
        if (Initial_thin_age_shift < 0) Additional_shift <- 0
      }
      if (Stock_hist_T[thin + 1] != 0 && T >= Stock_hist_T[thin + 1] + 1 && T < Stock_hist_T[thin + 1] + 2) {
        Total_thin_age_shift <- Total_thin_age_shift + Additional_shift
      }
    }

    Adj_T <- T - Total_thin_age_shift
    if (Adj_T - TBH <= 0) {
      DBH_vec[idx] <- 0
    } else {
      DBH_vec[idx] <- DBH_mod(Adj_T, D300_30_est, H30, TBH, N_vec[idx],
                               DBH_model, DBH_form, DBH_a, DBH_b, DBH_c,
                               DBH_d, DBH_f, DBH_g, DBH_h, DBH_k,
                               MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
    }

    T_adj_vec[idx] <- Adj_T
    if (DBH_vec[idx] <= 0) {
      DBH_vec[idx] <- 0
      BA_vec[idx] <- 0
      Vol_vec[idx] <- 0
    } else {
      BA_vec[idx] <- N_vec[idx] * pi * (DBH_vec[idx] / 200)^2
      Vol_vec[idx] <- Vol_stand(MTH_vec[idx], DBH_vec[idx], N_vec[idx],
                                 VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
      if (!Minimal_run) {
        Treelist <<- Scale_tree_list(Treelist, nstems, N_vec[idx], DBH_vec[idx])
        MTDia_val <- MTD(Treelist, nstems, N_vec[idx])
        petB <<- 1.98
        petA_val <- (MTH_vec[idx] - 1.4)^(-1 / 2.5) - petB / MTDia_val
        if (petA_val < 0) petA_val <- 0
        petA <<- petA_val
        Treelist <<- Predict_height(Treelist, nstems, N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT)
      }
    }
    WoodDen_vec[idx] <- (Den_a + Den_b * log(T)) * WoodDensity_Adjustment
  }

  # Final harvest at rotation end
  if (!Cali && !Minimal_run) {
    fs_result <- Felled_stems(Species, Treelist, nstems, rotlength,
                              N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                              N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                              Stemlist, stemno_local)
    Stemlist <- fs_result$Stemlist
    stemno_local <- fs_result$stemno
  }
  if (!Cali && !Minimal_run) {
    ml_result <- Make_Logs(Species, Treelist, nstems, rotlength,
                           N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                           N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                           Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
    Logs_mat <- ml_result$Logs
    logno_local <- ml_result$logno
    log_vol_vec[rotlength + 1] <- ml_result$Harvest_volume
  }

  Vol_vec <- ms_earlyield(N_vec, DBH_vec, Vol_vec, rotlength)

  # Store results in parent environment for use by downstream functions
  assign("N", N_vec, envir = parent.frame())
  assign("MTH", MTH_vec, envir = parent.frame())
  assign("DBH", DBH_vec, envir = parent.frame())
  assign("BA", BA_vec, envir = parent.frame())
  assign("Vol", Vol_vec, envir = parent.frame())
  assign("WoodDen", WoodDen_vec, envir = parent.frame())
  assign("log_volume", log_vol_vec, envir = parent.frame())
  assign("T_adj", T_adj_vec, envir = parent.frame())
  assign("Stemlist", Stemlist, envir = parent.frame())
  assign("stemno", stemno_local, envir = parent.frame())
  assign("Logs", Logs_mat, envir = parent.frame())
  assign("logno", logno_local, envir = parent.frame())
  assign("Pre_thin_N", Pre_thin_N, envir = parent.frame())
  assign("Pre_thin_dbh", Pre_thin_dbh, envir = parent.frame())
  assign("Pre_thin_vol", Pre_thin_vol, envir = parent.frame())
  assign("Thin_dbh", Thin_dbh, envir = parent.frame())
  assign("Thin_vol", Thin_vol, envir = parent.frame())
}

#' Radiata pine yield table - reads from 300 Index worksheet output
Yield_Table_radiata <- function() {
  # For radiata, the growth is driven by the 300 Index model (OutputGrowth)
  # Results are already in the growth_result from OutputGrowth()
  growth_result <- OutputGrowth()

  N_vec <- numeric(rotlength + 1)
  MTH_vec <- numeric(rotlength + 1)
  DBH_vec <- numeric(rotlength + 1)
  BA_vec <- numeric(rotlength + 1)
  Vol_vec <- numeric(rotlength + 1)
  WoodDen_vec <- numeric(rotlength + 1)
  log_vol_vec <- numeric(rotlength + 1)

  Stemlist <- matrix(0, nrow = 10000, ncol = 6)
  Logs_mat <- matrix(0, nrow = 10000, ncol = 10)
  stemno_local <- 1
  logno_local <- 1
  Harvest_volume_local <- 0

  gdf <- growth_result$growth_df
  thin_count <- 0

  for (T in 0:rotlength) {
    idx <- T + 1
    row <- idx
    if (row <= nrow(gdf)) {
      N_vec[idx] <- as.numeric(gdf[row, "N"])
      MTH_vec[idx] <- as.numeric(gdf[row, "MTH"])
      Vol_vec[idx] <- as.numeric(gdf[row, "Vol"])
      BA_vec[idx] <- as.numeric(gdf[row, "BA"])
      DBH_vec[idx] <- as.numeric(gdf[row, "DBH"])
      WoodDen_vec[idx] <- as.numeric(gdf[row, "WoodDensity"]) * 1000 * WoodDensity_Adjustment

      # Check for thinning
      if ("N_post_thin" %in% names(gdf) && !is.na(gdf[row, "N_post_thin"]) && gdf[row, "N_post_thin"] != 0) {
        thin_count <- thin_count + 1
        Pre_thin_N[thin_count] <- N_vec[idx]
        Pre_thin_dbh[thin_count] <- DBH_vec[idx]
        Pre_thin_vol[thin_count] <- Vol_vec[idx]
        N_vec[idx] <- as.numeric(gdf[row, "N_post_thin"])
        if ("Vol_post_thin" %in% names(gdf)) Vol_vec[idx] <- as.numeric(gdf[row, "Vol_post_thin"])
        if ("BA_post_thin" %in% names(gdf)) BA_vec[idx] <- as.numeric(gdf[row, "BA_post_thin"])
        if ("DBH_post_thin" %in% names(gdf)) DBH_vec[idx] <- as.numeric(gdf[row, "DBH_post_thin"])

        if (Pre_thin_N[thin_count] - N_vec[idx] <= 0 ||
            (Pre_thin_N[thin_count] * Pre_thin_dbh[thin_count]^2 - N_vec[idx] * DBH_vec[idx]^2) <= 0) {
          Thin_dbh[thin_count] <- Pre_thin_dbh[thin_count]
        } else {
          Thin_dbh[thin_count] <- sqrt((Pre_thin_N[thin_count] * Pre_thin_dbh[thin_count]^2 -
                                        N_vec[idx] * DBH_vec[idx]^2) / (Pre_thin_N[thin_count] - N_vec[idx]))
        }
        Thin_vol[thin_count] <- Pre_thin_vol[thin_count] - Vol_vec[idx]

        Treelist <<- Scale_tree_list(Treelist, nstems, Pre_thin_N[thin_count], Pre_thin_dbh[thin_count])
        MTDia_val <- MTD(Treelist, nstems, Pre_thin_N[thin_count])
        petB <<- 1.98
        petA_val <- (MTH_vec[idx] - 1.4)^(-1 / 2.5) - petB / MTDia_val
        if (petA_val < 0) {
          petA_val <- 0
          petB <<- MnHt_from_MTH(MTH_vec[idx], N_vec[idx], MTH_MnHt_a, MTH_MnHt_b)
        }
        petA <<- petA_val
        Treelist <<- Predict_height(Treelist, nstems, N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT)

        if (!Cali) {
          fs_result <- Felled_stems(Species, Treelist, nstems, T,
                                    Pre_thin_N[thin_count] - N_vec[idx], Thin_dbh[thin_count],
                                    Thin_vol[thin_count],
                                    N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                    Stemlist, stemno_local)
          Stemlist <- fs_result$Stemlist
          stemno_local <- fs_result$stemno
        }
        if (!Cali && Stock_hist_Type[thin_count + 1] == 2) {
          ml_result <- Make_Logs(Species, Treelist, nstems, T,
                                Pre_thin_N[thin_count] - N_vec[idx], Thin_dbh[thin_count],
                                Thin_vol[thin_count],
                                N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
          Logs_mat <- ml_result$Logs
          logno_local <- ml_result$logno
          log_vol_vec[idx] <- ml_result$Harvest_volume
        }
      }
    }
  }

  # Final harvest
  Treelist <<- Scale_tree_list(Treelist, nstems, N_vec[rotlength + 1], DBH_vec[rotlength + 1])
  MTDia_val <- MTD(Treelist, nstems, N_vec[rotlength + 1])
  petB <<- 1.98
  petA_val <- (MTH_vec[rotlength + 1] - 1.4)^(-1 / 2.5) - petB / MTDia_val
  if (petA_val < 0) {
    petA_val <- 0
    petB <<- MnHt_from_MTH(MTH_vec[rotlength + 1], N_vec[rotlength + 1], MTH_MnHt_a, MTH_MnHt_b)
  }
  petA <<- petA_val
  Treelist <<- Predict_height(Treelist, nstems, N_vec[rotlength + 1], DBH_vec[rotlength + 1],
                              MTH_vec[rotlength + 1], PRUNEHT)
  fs_result <- Felled_stems(Species, Treelist, nstems, rotlength,
                            N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                            N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                            Stemlist, stemno_local)
  Stemlist <- fs_result$Stemlist
  stemno_local <- fs_result$stemno
  ml_result <- Make_Logs(Species, Treelist, nstems, rotlength,
                         N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                         N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                         Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
  Logs_mat <- ml_result$Logs
  logno_local <- ml_result$logno
  log_vol_vec[rotlength + 1] <- ml_result$Harvest_volume

  Vol_vec <- ms_earlyield(N_vec, DBH_vec, Vol_vec, rotlength)

  assign("N", N_vec, envir = parent.frame())
  assign("MTH", MTH_vec, envir = parent.frame())
  assign("DBH", DBH_vec, envir = parent.frame())
  assign("BA", BA_vec, envir = parent.frame())
  assign("Vol", Vol_vec, envir = parent.frame())
  assign("WoodDen", WoodDen_vec, envir = parent.frame())
  assign("log_volume", log_vol_vec, envir = parent.frame())
  assign("Stemlist", Stemlist, envir = parent.frame())
  assign("stemno", stemno_local, envir = parent.frame())
  assign("Logs", Logs_mat, envir = parent.frame())
  assign("logno", logno_local, envir = parent.frame())
  assign("Pre_thin_N", Pre_thin_N, envir = parent.frame())
  assign("Pre_thin_dbh", Pre_thin_dbh, envir = parent.frame())
  assign("Pre_thin_vol", Pre_thin_vol, envir = parent.frame())
  assign("Thin_dbh", Thin_dbh, envir = parent.frame())
  assign("Thin_vol", Thin_vol, envir = parent.frame())
}

#' Douglas-fir yield table - uses Dfir_Yield from DouglasFir_500Index.R
Yield_Table_dfir <- function() {
  # Run the Douglas-fir growth model
  dfir_result <- dfir_yield()

  N_vec <- numeric(rotlength + 1)
  MTH_vec <- numeric(rotlength + 1)
  DBH_vec <- numeric(rotlength + 1)
  BA_vec <- numeric(rotlength + 1)
  Vol_vec <- numeric(rotlength + 1)
  WoodDen_vec <- numeric(rotlength + 1)
  log_vol_vec <- numeric(rotlength + 1)

  Stemlist <- matrix(0, nrow = 10000, ncol = 6)
  Logs_mat <- matrix(0, nrow = 10000, ncol = 10)
  stemno_local <- 1
  logno_local <- 1

  # Extract results from dfir model
  if (!is.null(dfir_result)) {
    for (T in 0:rotlength) {
      idx <- T + 1
      if (idx <= length(dfir_result$N)) {
        N_vec[idx] <- dfir_result$N[idx]
        MTH_vec[idx] <- dfir_result$MTH[idx]
        DBH_vec[idx] <- dfir_result$DBH[idx]
        BA_vec[idx] <- dfir_result$BA[idx]
        Vol_vec[idx] <- dfir_result$Vol[idx]
        WoodDen_vec[idx] <- dfir_result$WoodDen[idx] * WoodDensity_Adjustment
      }
    }
  }

  # Process thinnings
  for (T in 1:rotlength) {
    idx <- T + 1
    log_vol_vec[idx] <- 0
    for (thin in 1:4) {
      if (Stock_hist_T[thin + 1] != 0 && T >= Stock_hist_T[thin + 1] && T < Stock_hist_T[thin + 1] + 1) {
        Treelist <<- Scale_tree_list(Treelist, nstems, Pre_thin_N[thin], Pre_thin_dbh[thin])
        MTDia_val <- MTD(Treelist, nstems, Pre_thin_N[thin])
        petB <<- 1.98
        petA_val <- (MTH_vec[idx] - 1.4)^(-1 / 2.5) - petB / MTDia_val
        if (petA_val < 0) {
          petA_val <- 0
          petB <<- MnHt_from_MTH(MTH_vec[idx], N_vec[idx], MTH_MnHt_a, MTH_MnHt_b)
        }
        petA <<- petA_val
        Treelist <<- Predict_height(Treelist, nstems, N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT)
        N_vec[idx] <- Stock_hist_N[thin + 1]

        if (Pre_thin_N[thin] - N_vec[idx] <= 0 ||
            (Pre_thin_N[thin] * Pre_thin_dbh[thin]^2 - N_vec[idx] * DBH_vec[idx]^2) <= 0) {
          Thin_dbh[thin] <- Pre_thin_dbh[thin]
        } else {
          Thin_dbh[thin] <- sqrt((Pre_thin_N[thin] * Pre_thin_dbh[thin]^2 -
                                  N_vec[idx] * DBH_vec[idx]^2) / (Pre_thin_N[thin] - N_vec[idx]))
        }
        Thin_vol[thin] <- Pre_thin_vol[thin] - Vol_vec[idx]

        if (!Cali) {
          fs_result <- Felled_stems(Species, Treelist, nstems, T,
                                    Pre_thin_N[thin] - N_vec[idx], Thin_dbh[thin], Thin_vol[thin],
                                    N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                    Stemlist, stemno_local)
          Stemlist <- fs_result$Stemlist
          stemno_local <- fs_result$stemno
        }
        if (!Cali && Stock_hist_Type[thin + 1] == 2) {
          ml_result <- Make_Logs(Species, Treelist, nstems, T,
                                Pre_thin_N[thin] - N_vec[idx], Thin_dbh[thin], Thin_vol[thin],
                                N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT,
                                Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
          Logs_mat <- ml_result$Logs
          logno_local <- ml_result$logno
          log_vol_vec[idx] <- ml_result$Harvest_volume
        }
      }
    }

    if (DBH_vec[idx] > 0) {
      Treelist <<- Scale_tree_list(Treelist, nstems, N_vec[idx], DBH_vec[idx])
      MTDia_val <- MTD(Treelist, nstems, N_vec[idx])
      petB <<- 1.98
      petA_val <- (MTH_vec[idx] - 1.4)^(-1 / 2.5) - petB / MTDia_val
      if (petA_val < 0) petA_val <- 0
      petA <<- petA_val
      Treelist <<- Predict_height(Treelist, nstems, N_vec[idx], DBH_vec[idx], MTH_vec[idx], PRUNEHT)
    }
    WoodDen_vec[idx] <- WoodDen_vec[idx] * WoodDensity_Adjustment
  }

  # Final harvest
  if (!Cali) {
    fs_result <- Felled_stems(Species, Treelist, nstems, rotlength,
                              N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                              N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                              Stemlist, stemno_local)
    Stemlist <- fs_result$Stemlist
    stemno_local <- fs_result$stemno
  }
  if (!Cali) {
    ml_result <- Make_Logs(Species, Treelist, nstems, rotlength,
                           N_vec[rotlength + 1], DBH_vec[rotlength + 1], Vol_vec[rotlength + 1],
                           N_vec[rotlength + 1], DBH_vec[rotlength + 1], MTH_vec[rotlength + 1], PRUNEHT,
                           Logs_mat, logno_local, log_length, min_SED, break_height, log_losses)
    Logs_mat <- ml_result$Logs
    logno_local <- ml_result$logno
    log_vol_vec[rotlength + 1] <- ml_result$Harvest_volume
  }

  assign("N", N_vec, envir = parent.frame())
  assign("MTH", MTH_vec, envir = parent.frame())
  assign("DBH", DBH_vec, envir = parent.frame())
  assign("BA", BA_vec, envir = parent.frame())
  assign("Vol", Vol_vec, envir = parent.frame())
  assign("WoodDen", WoodDen_vec, envir = parent.frame())
  assign("log_volume", log_vol_vec, envir = parent.frame())
  assign("Stemlist", Stemlist, envir = parent.frame())
  assign("stemno", stemno_local, envir = parent.frame())
  assign("Logs", Logs_mat, envir = parent.frame())
  assign("logno", logno_local, envir = parent.frame())
  assign("Pre_thin_N", Pre_thin_N, envir = parent.frame())
  assign("Pre_thin_dbh", Pre_thin_dbh, envir = parent.frame())
  assign("Pre_thin_vol", Pre_thin_vol, envir = parent.frame())
  assign("Thin_dbh", Thin_dbh, envir = parent.frame())
  assign("Thin_vol", Thin_vol, envir = parent.frame())
}

# ===========================
# Calibration functions
# ===========================

#' Generic calibration: estimate D300_30 from measured DBH using bisection
Calibrate <- function() {
  Cali <<- TRUE

  if (Species == "Coast redwood" && T2 != 0) {
    MTH_b <<- MTHmodel_b(T1, H1, T2, H2)
  }

  H30 <<- SI_eqn(T1, H1, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
  TBH <<- AgeBH(30, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)

  # Make stocking history consistent with calibration measurement
  for (i in 1:Nthins) {
    if (as.integer(T1) < Stock_hist_T[i + 1]) {
      for (j in seq(Nthins, i, by = -1)) {
        Stock_hist_T[j + 2] <<- Stock_hist_T[j + 1]
        Stock_hist_N[j + 2] <<- Stock_hist_N[j + 1]
        Stock_hist_thin_coeff[j + 2] <<- Stock_hist_thin_coeff[j + 1]
      }
      Stock_hist_T[i + 1] <<- as.integer(T1)
      Stock_hist_N[i + 1] <<- N1
      Stock_hist_thin_coeff[i + 1] <<- 1
      break
    }
    if (as.integer(T1) == Stock_hist_T[i + 1]) {
      Stock_hist_T[i + 1] <<- as.integer(T1)
      Stock_hist_N[i + 1] <<- N1
      Stock_hist_thin_coeff[i + 1] <<- 1
      break
    }
  }
  if (as.integer(T1) > Stock_hist_T[Nthins + 1]) {
    Stock_hist_T[Nthins + 2] <<- as.integer(T1)
    Stock_hist_N[Nthins + 2] <<- N1
    Stock_hist_thin_coeff[Nthins + 2] <<- 1
  }

  # Bisection to estimate D300_30
  D300_30lo <- 10
  D300_30up <- 120
  if (DBH_model == "Korf" && DBH_form == "CA") D300_30up <- 80
  if (DBH_model == "Richards" && DBH_form == "Anamorphic") D300_30up <- 90

  Rotlength_temp <- rotlength
  rotlength <<- 200

  for (i in 1:16) {
    D300_30mid <- (D300_30lo + D300_30up) / 2

    D300_30_est <<- D300_30lo
    Yield_Table()
    Pred_D <- DBH[as.integer(T1) + 1] +
              (DBH[as.integer(T1 + 1) + 1] - DBH[as.integer(T1) + 1]) * (T1 - as.integer(T1))
    flo <- Pred_D - D1

    D300_30_est <<- D300_30up
    Yield_Table()
    Pred_D <- DBH[as.integer(T1) + 1] +
              (DBH[as.integer(T1 + 1) + 1] - DBH[as.integer(T1) + 1]) * (T1 - as.integer(T1))
    fup <- Pred_D - D1

    D300_30_est <<- D300_30mid
    Yield_Table()
    Pred_D <- DBH[as.integer(T1) + 1] +
              (DBH[as.integer(T1 + 1) + 1] - DBH[as.integer(T1) + 1]) * (T1 - as.integer(T1))
    fmid <- Pred_D - D1

    if (flo * fmid < 0) D300_30up <- D300_30mid else D300_30lo <- D300_30mid
  }

  D300_30_est <<- D300_30mid
  I300 <<- I300_from_SI_D300_30(H30, D300_30_est, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
  rotlength <<- Rotlength_temp
  Cali <<- FALSE
}

#' Douglas-fir calibration via CombineSolver (from DouglasFir_500Index.R)
Calibrate_dfir <- function() {
  # Uses the Douglas-fir solver from DouglasFir_500Index.R
  dfir_result <- dfir_si_solver(T1, H1, N1, D1, Stock_hist_N, Stock_hist_T)
  I300 <<- dfir_result$I500

  H30 <<- dfir_result$SI
}

# ===========================
# Carbon / CChange integration
# ===========================

#' Initialize CChange model inputs from multi-species yield table
Initialise_C_Change <- function() {
  # Determine species code for CChange
  if (Species == "Radiata pine") {
    sp_code <- "PRAD"
  } else if (Species %in% c("Blackwood", "Eucalyptus regnans", "Eucalyptus fastigata",
                             "Eucalyptus nitens (N.I.)", "Eucalyptus nitens (S.I.)",
                             "Eucalyptus delegatensis", "Eucalyptus saligna")) {
    sp_code <- "EUC"
  } else if (Species == "Coast redwood") {
    sp_code <- "RED"
  } else if (Species %in% c("Cupressus macrocarpa (N.I.)", "Cupressus macrocarpa (S.I.)",
                             "Cupressus lusitanica (N.I.)", "Cupressus lusitanica (S.I.)")) {
    sp_code <- "CLUS"
  } else if (Species == "Douglas-fir") {
    sp_code <- "PMEN"
  } else {
    sp_code <- "PRAD"
  }

  # Determine density type
  if (Species %in% c("Radiata pine", "Douglas-fir")) {
    density_type <- "S"  # Sheath Density
  } else {
    density_type <- "T"  # Whole Tree Density
  }

  # Build CChange input data frame
  cchange_input <- data.frame(
    Age = 0:rotlength,
    N = N[1:(rotlength + 1)],
    MTH = MTH[1:(rotlength + 1)],
    Vol = Vol[1:(rotlength + 1)],
    WoodDen = WoodDen[1:(rotlength + 1)],
    DBH = DBH[1:(rotlength + 1)],
    log_volume = log_volume[1:(rotlength + 1)]
  )

  # Store for downstream use
  assign("cchange_input", cchange_input, envir = parent.frame())
  assign("sp_code", sp_code, envir = parent.frame())
  assign("density_type", density_type, envir = parent.frame())

  # Add thinning info
  thin_info <- list()
  for (thin in 1:4) {
    if (Stock_hist_T[thin + 1] != 0) {
      thin_info[[thin]] <- list(
        type = if (Stock_hist_Type[thin + 1] == 1) "W" else "P",
        age = Stock_hist_T[thin + 1],
        post_thin_N = Stock_hist_N[thin + 1]
      )
    }
  }
  assign("thin_info", thin_info, envir = parent.frame())
}

#' Estimate mortality volume (dead wood volume from natural mortality)
Mortality_Volume <- function() {
  dead_vol <- numeric(rotlength + 1)
  for (T in 1:rotlength) {
    idx <- T + 1
    prev_idx <- T
    if (N[prev_idx] > 0 && Vol[prev_idx] > 0) {
      dead_stems <- N[prev_idx] - N[idx]
      if (dead_stems > 0) {
        dead_vol[idx] <- Vol[prev_idx] * dead_stems / N[prev_idx]
      }
    }
  }
  assign("dead_vol", dead_vol, envir = parent.frame())
}

#' Run CChange carbon model (wrapper around CChange_model.R functions)
Run_C_Change <- function() {
  if (exists("run_cchange", mode = "function")) {
    tryCatch({
      cchange_result <- run_cchange(
        sp_code = sp_code,
        rotlength = rotlength,
        N = N,
        MTH = MTH,
        Vol = Vol,
        WoodDen = WoodDen,
        DBH = DBH,
        dead_vol = dead_vol,
        Stock_hist_N = Stock_hist_N,
        Stock_hist_T = Stock_hist_T,
        Stock_hist_Type = Stock_hist_Type,
        latitude = latitude,
        elevation = elevation,
        Soil_C = Soil_C,
        Soil_N = Soil_N,
        MAT = MAT
      )
      assign("cchange_result", cchange_result, envir = parent.frame())
    }, error = function(e) {
      message("CChange model failed: ", e$message)
      assign("cchange_result", NULL, envir = parent.frame())
    })
  }
}

#' Kizha-Han redwood carbon allometric model
Kizha_Han <- function() {
  CFraction_wood <- 0.53
  CFraction_bark <- 0.519
  CFraction_roots <- 0.53
  CFraction_needles <- 0.495
  CFraction_branches <- 0.495
  CFraction_DWL <- 0.5
  bark_wood_ratio <- 0.18
  bark_density <- 437
  root_shoot_ratio <- 0.23

  carbon_results <- data.frame(
    Age = 0:rotlength,
    total_C = numeric(rotlength + 1),
    AGB_C = numeric(rotlength + 1),
    BGB_C = numeric(rotlength + 1),
    DWL_C = numeric(rotlength + 1),
    litter_C = numeric(rotlength + 1),
    stem_wood_C = numeric(rotlength + 1),
    bark_C = numeric(rotlength + 1),
    live_branch_C = numeric(rotlength + 1),
    dead_branch_C = numeric(rotlength + 1),
    foliage_C = numeric(rotlength + 1)
  )

  AGCWD <- 0
  BGCWD <- 0

  for (i in 0:rotlength) {
    idx <- i + 1
    Age <- i
    stocking <- N[idx]
    Vol_val <- Vol[idx]
    DBH_val <- DBH[idx]
    MTH_val <- MTH[idx]
    wooddensity <- WoodDen[idx]
    dead_vol_val <- if (exists("dead_vol")) dead_vol[idx] else 0
    post_thin_stocking <- 0
    post_thin_vol <- 0
    post_thin_dbh <- 0
    harvest_vol <- log_volume[idx]

    # Check for thinning at this age
    for (thin in 1:4) {
      if (Stock_hist_T[thin + 1] != 0 && i == Stock_hist_T[thin + 1]) {
        post_thin_stocking <- Stock_hist_N[thin + 1]
        if (exists("Pre_thin_vol") && thin <= length(Pre_thin_vol)) {
          post_thin_vol <- Vol[idx]
          post_thin_dbh <- DBH[idx]
        }
      }
    }

    if (DBH_val == 0) {
      biomass_stem_wood <- wooddensity * Vol_val / 1000
      biomass_bark <- 0
      biomass_live_branches <- 0
      biomass_dead_branches <- 0
      biomass_foliage <- 0
      BGB <- root_shoot_ratio * biomass_stem_wood
      AGCWD <- 0
      BGCWD <- 0
    } else {
      Treelist_local <- Scale_tree_list(Treelist, nstems, stocking, DBH_val)
      biomass_stem_wood <- wooddensity * Vol_val / 1000
      biomass_bark <- Vol_val * bark_wood_ratio * bark_density / 1000
      biomass_harvested <- harvest_vol * wooddensity / 1000 + harvest_vol * bark_wood_ratio * bark_density / 1000
      biomass_live_branches <- 0
      biomass_dead_branches <- 0
      biomass_foliage <- 0
      for (tree in 1:nstems) {
        biomass_live_branches <- biomass_live_branches + (0.01475 * Treelist_local[tree, 3]^2.0382) / 1000
        biomass_dead_branches <- biomass_dead_branches + (0.00038117 * Treelist_local[tree, 3]^2.3257) / 1000
        biomass_foliage <- biomass_foliage + (0.05064 * Treelist_local[tree, 3]^1.5819) / 1000
      }
      biomass_live_branches <- biomass_live_branches / nstems * stocking
      biomass_dead_branches <- biomass_dead_branches / nstems * stocking
      biomass_foliage <- biomass_foliage / nstems * stocking
      AGB <- biomass_stem_wood + biomass_bark + biomass_live_branches + biomass_dead_branches + biomass_foliage
      BGB <- AGB * root_shoot_ratio

      AGCWD <- AGCWD * 0.5^(1 / AGCWD_half_life)
      BGCWD <- BGCWD * 0.5^(1 / BGCWD_half_life)
      if (Vol_val > 0) {
        AGCWD <- AGCWD + (biomass_stem_wood + biomass_bark) * dead_vol_val / Vol_val
        BGCWD <- BGCWD + BGB * dead_vol_val / Vol_val
      }

      if (post_thin_stocking != 0) {
        pre_thin_biomass_stem_wood <- biomass_stem_wood
        pre_thin_biomass_bark <- biomass_bark
        pre_thin_BGB <- BGB
        stocking <- post_thin_stocking
        Vol_val <- post_thin_vol
        DBH_val <- post_thin_dbh
        Treelist_local <- Scale_tree_list(Treelist, nstems, stocking, DBH_val)
        biomass_stem_wood <- wooddensity * Vol_val / 1000
        biomass_bark <- Vol_val * bark_wood_ratio * bark_density / 1000
        biomass_live_branches <- 0
        biomass_dead_branches <- 0
        biomass_foliage <- 0
        for (tree in 1:nstems) {
          biomass_live_branches <- biomass_live_branches + (0.01475 * Treelist_local[tree, 3]^2.0382) / 1000
          biomass_dead_branches <- biomass_dead_branches + (0.0003817 * Treelist_local[tree, 3]^2.3257) / 1000
          biomass_foliage <- biomass_foliage + (0.05064 * Treelist_local[tree, 3]^1.5819) / 1000
        }
        biomass_live_branches <- biomass_live_branches / nstems * stocking
        biomass_dead_branches <- biomass_dead_branches / nstems * stocking
        biomass_foliage <- biomass_foliage / nstems * stocking
        AGB <- biomass_stem_wood + biomass_bark + biomass_live_branches + biomass_dead_branches + biomass_foliage
        BGB <- AGB * root_shoot_ratio
        AGCWD <- AGCWD + pre_thin_biomass_stem_wood - biomass_stem_wood + pre_thin_biomass_bark - biomass_bark - biomass_harvested
        BGCWD <- BGCWD + pre_thin_BGB - BGB
      }
    }

    carbon_results$litter_C[idx] <- 0
    carbon_results$DWL_C[idx] <- (AGCWD + BGCWD) * CFraction_DWL
    carbon_results$BGB_C[idx] <- BGB * CFraction_roots
    carbon_results$AGB_C[idx] <- biomass_stem_wood * CFraction_wood +
                                  biomass_bark * CFraction_bark +
                                  biomass_live_branches * CFraction_branches +
                                  biomass_dead_branches * CFraction_DWL +
                                  biomass_foliage * CFraction_needles
    carbon_results$total_C[idx] <- carbon_results$litter_C[idx] + carbon_results$DWL_C[idx] +
                                    carbon_results$BGB_C[idx] + carbon_results$AGB_C[idx]
    carbon_results$stem_wood_C[idx] <- biomass_stem_wood * CFraction_wood
    carbon_results$bark_C[idx] <- biomass_bark * CFraction_bark
    carbon_results$live_branch_C[idx] <- biomass_live_branches * CFraction_branches
    carbon_results$dead_branch_C[idx] <- biomass_dead_branches * CFraction_branches
    carbon_results$foliage_C[idx] <- biomass_foliage * CFraction_needles
  }

  assign("carbon_results", carbon_results, envir = parent.frame())
}

#' Output table: build results data frame (replaces Excel output)
Output_table <- function() {
  results <- data.frame(
    Age = 0:rotlength,
    N = N[1:(rotlength + 1)],
    BA = BA[1:(rotlength + 1)],
    DBH = DBH[1:(rotlength + 1)],
    Vol = Vol[1:(rotlength + 1)],
    log_volume = log_volume[1:(rotlength + 1)],
    MTH = MTH[1:(rotlength + 1)],
    MnHt = sapply(1:(rotlength + 1), function(i) {
      MnHt_from_MTH(MTH[i], N[i], MTH_MnHt_a, MTH_MnHt_b)
    }),
    WoodDen = WoodDen[1:(rotlength + 1)]
  )

  if (Species == "Coast redwood") {
    results$CrownHt <- sapply(1:(rotlength + 1), function(i) {
      CrownHeight(MnHt_from_MTH(MTH[i], N[i], MTH_MnHt_a, MTH_MnHt_b), N[i])
    })
  }

  # Add carbon columns if available
  if (exists("carbon_results", envir = parent.frame())) {
    cr <- get("carbon_results", envir = parent.frame())
    results$total_C <- cr$total_C
    results$AGB_C <- cr$AGB_C
    results$BGB_C <- cr$BGB_C
    results$DWL_C <- cr$DWL_C
    results$litter_C <- cr$litter_C
  } else if (exists("cchange_result", envir = parent.frame())) {
    ccr <- get("cchange_result", envir = parent.frame())
    if (!is.null(ccr) && is.data.frame(ccr)) {
      for (col in names(ccr)) {
        if (nrow(ccr) == nrow(results)) {
          results[[col]] <- ccr[[col]]
        }
      }
    }
  }

  # Build felled stems output
  if (exists("Stemlist") && exists("stemno")) {
    if (stemno > 1) {
      felled_stems_df <- as.data.frame(Stemlist[1:(stemno - 1), , drop = FALSE])
      names(felled_stems_df) <- c("Age", "StemNo", "SPH", "DBH", "Height", "Volume")
      assign("felled_stems_df", felled_stems_df, envir = parent.frame())
    }
  }

  # Build harvest logs output
  if (exists("Logs") && exists("logno")) {
    if (logno > 1) {
      logs_df <- as.data.frame(Logs[1:(logno - 1), 1:8, drop = FALSE])
      names(logs_df) <- c("Age", "SPH", "StemNo", "LogNo", "SED_mm", "LED_mm", "Length", "Volume")
      assign("logs_df", logs_df, envir = parent.frame())
    }
  }

  # Build harvest summary
  if (exists("Logs") && exists("logno")) {
    harvest_sum <- harvest_summary_fn(Logs, logno)
    assign("harvest_sum", harvest_sum, envir = parent.frame())
  }

  assign("yield_table", results, envir = parent.frame())
  return(results)
}

# ===========================
# Batch functions
# ===========================

#' Batch estimation of site indices for PSP data
Batch_estimate_indices <- function(batch_data) {
  results <- data.frame(
    row = integer(0),
    I300 = numeric(0),
    SI = numeric(0)
  )

  Check_errors <<- FALSE
  Minimal_run <<- TRUE

  for (row in 1:nrow(batch_data)) {
    tryCatch({
      sp <- batch_data[row, "Species"]
      # Set species and inputs
      Species <<- sp
      Input_parameters()

      # Set calibration inputs from batch data
      Stock_hist_N[1] <<- as.numeric(batch_data[row, "InitialStocking"])
      T1 <<- as.numeric(batch_data[row, "Age"])
      N1 <<- as.numeric(batch_data[row, "Stocking"])
      H1 <<- as.numeric(batch_data[row, "MTH"])
      D1 <<- as.numeric(batch_data[row, "BA"])

      run_model()

      results <- rbind(results, data.frame(
        row = row,
        I300 = I300,
        SI = H30
      ))
    }, error = function(e) {
      message("Batch row ", row, " failed: ", e$message)
    })
  }

  Check_errors <<- TRUE
  Minimal_run <<- FALSE
  return(results)
}

#' Batch carbon estimation at ages 10, 20, 30, 40, 50, 60
Batch_carbon <- function(batch_data) {
  results <- list()

  for (i in 1:nrow(batch_data)) {
    tryCatch({
      # Set inputs from batch data
      I300 <<- as.numeric(batch_data[i, "I300"])
      H30 <<- as.numeric(batch_data[i, "SI"])
      Stock_hist_N[1] <<- as.numeric(batch_data[i, "InitialStocking"])
      if ("ThinAge" %in% names(batch_data)) {
        Stock_hist_T[2] <<- as.numeric(batch_data[i, "ThinAge"])
      }
      if ("ThinN" %in% names(batch_data)) {
        Stock_hist_N[2] <<- as.numeric(batch_data[i, "ThinN"])
      }

      run_model()

      row_result <- list(row = i)
      for (Age in seq(10, 60, by = 10)) {
        idx <- Age + 1
        if (idx <= length(Vol)) {
          row_result[[paste0("Vol_", Age)]] <- Vol[idx]
          if (exists("carbon_results")) {
            row_result[[paste0("total_C_", Age)]] <- carbon_results$total_C[idx]
            row_result[[paste0("AGB_C_", Age)]] <- carbon_results$AGB_C[idx]
            row_result[[paste0("BGB_C_", Age)]] <- carbon_results$BGB_C[idx]
            row_result[[paste0("DWL_C_", Age)]] <- carbon_results$DWL_C[idx]
            row_result[[paste0("litter_C_", Age)]] <- carbon_results$litter_C[idx]
          }
        }
      }
      results[[i]] <- as.data.frame(row_result)
    }, error = function(e) {
      message("Batch carbon row ", i, " failed: ", e$message)
    })
  }

  return(do.call(rbind, results))
}
