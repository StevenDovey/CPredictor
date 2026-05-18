#next need to split this into separate modules for data prep, and processing

library(readxl)
library(writexl)

if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}

age300 <- 30.0
DBHcalage <- 28.7
voltable <- 1L
driftO <- 0.0
bias_old <- FALSE
bias_young <- FALSE
bias_SI <- FALSE

# Normalise parameter names for lookup (VBA / Excel sheets vary in spelling and case).
norm_param_key <- function(x) {
  x <- tolower(trimws(as.character(x)))
  gsub("[^a-z0-9]", "", x)
}

param_pick_row <- function(param_names, aliases) {
  nk <- vapply(param_names, norm_param_key, "")
  for (a in aliases) {
    w <- which(nk == norm_param_key(a))
    if (length(w)) return(w[1L])
  }
  NA_integer_
}

# Excel often uses "x" for bias toggles; as.numeric() alone would lose that.
param_as_logical <- function(chr, num) {
  if (is.finite(num) && num != 0) return(TRUE)
  s <- tolower(trimws(as.character(chr)))
  if (!nzchar(s) || is.na(s)) return(FALSE)
  s %in% c("x", "1", "true", "yes", "y", "t")
}

# VBA Module1 Sub voltab() — coefficients are hardcoded there; Excel only marks which row (1–11) via "x".
voltab_matrix_module1 <- function() {
  m <- matrix(NA_real_, nrow = 11L, ncol = 8L)
  m[1L, 1:3] <- c(0.942, -1.161, 0.317)
  m[2L, 1:3] <- c(0.989, -1.2752, 0.3191)
  m[3L, ] <- c(1.492912924, -0.999113309, 1.250753941, -0.397037159, 0.027218164, -0.063166205, 0.064609459, -0.030665365)
  m[4L, ] <- c(1.633105986, -1.039327204, 1.212696953, -0.359131176, 0.026454943, -0.067457458, 0.066992488, -0.030528278)
  m[5L, ] <- c(0.730448717, -0.617440226, 1.095616037, -0.222220223, 0.013858949, -0.11022445, 0.059157535, -0.016942593)
  m[6L, ] <- c(1.09857999, -0.883862258, 1.165375013, -0.28047221, 0.022081234, -0.059261776, 0.053187392, -0.025226521)
  m[7L, ] <- c(1.403009551, -0.96392392, 1.221046594, -0.358337009, 0.024975712, -0.061374804, 0.061895757, -0.028672533)
  m[8L, ] <- c(2.834246614, -1.856804825, 1.152097786, -0.201346156, -0.000721117, 0.081503044, 0.024428222, 0.001938887)
  m[9L, ] <- c(2.7023, -2.1301, 1.3901, -0.5056, 0.0548, 0.0991, 0.1478, -0.088)
  m[10L, 1:3] <- c(6.2733, 0.1284, -0.00097)
  m[11L, 1:3] <- c(2.1819, 0.2504, -0.00081)
  m
}

load_model_inputs <- function(input_file = "input.xlsx", voltab_source = c("module1", "input")) {
  voltab_source <- match.arg(voltab_source)
  parameters <- read_excel(input_file, sheet = "parameters", .name_repair = "minimal")[, 1:4]
  nm <- trimws(as.character(parameters$name))
  val_chr <- trimws(as.character(parameters$value))
  val_num <- suppressWarnings(as.numeric(val_chr))
  coefs <- as.list(setNames(val_num, nm))
  if (voltab_source == "module1") {
    v <- voltab_matrix_module1()
  } else {
    voltab <- read_excel(input_file, sheet = "VolTab", range = "A1:K9", col_names = TRUE, .name_repair = "minimal")
    v <- t(data.matrix(voltab))
  }
  # SI and 300 Index are always computed in R (solve_si300). No "300 Index" or other index sheet is read.
  # Config must mirror VBA Inputparms: voltable, heightmodel, bias_*, drift (not hard-coded defaults).
  ir <- function(aliases) param_pick_row(nm, aliases)
  voltable <- {
    j <- ir(c("voltable", "vol_table", "VolTab", "densitymodel", "DensityModel"))
    if (is.na(j)) 1L else {
      k <- as.integer(val_num[j])
      if (is.finite(k) && k >= 1L && k <= 11L) k else 1L
    }
  }
  height_model <- {
    j <- ir(c("height_model", "heightmodel", "Height_model", "HeightModel"))
    if (is.na(j)) 3L else {
      k <- as.integer(val_num[j])
      if (is.finite(k) && k %in% 1L:3L) k else 3L
    }
  }
  jbo <- ir(c("bias_old", "Bias_old", "BIAS_OLD"))
  bias_old_c <- if (is.na(jbo)) FALSE else param_as_logical(val_chr[jbo], val_num[jbo])
  jby <- ir(c("bias_young", "Bias_young", "BIAS_YOUNG"))
  bias_young_c <- if (is.na(jby)) FALSE else param_as_logical(val_chr[jby], val_num[jby])
  jbs <- ir(c("bias_SI", "bias_si", "Bias_SI", "BIAS_SI"))
  bias_SI_c <- if (is.na(jbs)) FALSE else param_as_logical(val_chr[jbs], val_num[jbs])
  drift <- {
    j <- ir(c("drift", "driftO", "Drift", "DRIFT"))
    if (is.na(j)) 0 else {
      x <- val_num[j]
      if (is.finite(x)) x else 0
    }
  }
  config <- list(
    voltable = voltable,
    height_model = height_model,
    bias_old = bias_old_c,
    bias_young = bias_young_c,
    bias_SI = bias_SI_c,
    drift = drift,
    voltab_source = voltab_source
  )
  list(coefs = coefs, v = v, config = config)
}

## Debug report helpers removed.

calcheightcoeff_model <- function(SI, latitude, elevation, height_model, pars) {
  if (height_model == 1L) {
    ha <- exp(pars$hNSWa)
    hb <- 1 / (pars$hNSWb + pars$hNSWp * SI)
  } else if (height_model == 2L) {
    ha <- exp(pars$ha0 + pars$ha1 * SI)
    hb <- 1 / (pars$hb0 + pars$hb1 * SI)
  } else {
    ha <- exp(pars$hae0 + pars$hae1 * latitude + pars$hae2 * elevation)
    hb <- 1 / (pars$hbe0 + pars$hbe1 * SI)
  }
  list(ha = ha, hb = hb)
}

CalcMTH_model <- function(SI, Age, latitude, elevation, height_model, pars) {
  hc <- calcheightcoeff_model(SI, latitude, elevation, height_model, pars)
  0.25 + (SI - 0.25) * ((1 - exp(-hc$ha * Age)) / (1 - exp(-hc$ha * 20)))^hc$hb
}

# VBA Module1 heightmod(): model 3 (environmental NZ) only if lat and elev exist and lat in [30, 48] DD.
# Otherwise use simple NZ (2). NSW (1) from config is unchanged. This avoids dropping SI/300 when elev
# or lat is missing but the workbook would still run with height model 2.
resolve_height_model_for_stand <- function(latitude, elevation, cfg_height_model) {
  hm <- suppressWarnings(as.integer(cfg_height_model))
  if (!is.finite(hm) || hm < 1L || hm > 3L) hm <- 3L
  if (hm %in% 1L:2L) return(hm)
  latp <- latitude_for_vba_model(latitude)
  if (is.finite(latp) && is.finite(elevation) && latp >= 30 && latp <= 48) return(3L)
  2L
}

CalcDBHfromBA <- function(BA, N) sqrt(1.273 * BA / N) * 100

calcBAfromVol <- function(MTH, Vol, N, voltable, vmat) {
  if (Vol <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2)) {
    return(Vol / (MTH * (vmat[voltable, 1] * (MTH - 1.4)^vmat[voltable, 2] + vmat[voltable, 3])))
  }
  if (voltable %in% c(10, 11)) {
    return(Vol / (vmat[voltable, 1] + vmat[voltable, 2] * MTH + vmat[voltable, 3] * N))
  }
  exp(vmat[voltable, 1] + vmat[voltable, 2] * log(MTH) + vmat[voltable, 3] * log(Vol) +
        vmat[voltable, 4] * log(N) + vmat[voltable, 5] * log(N)^2 + vmat[voltable, 6] * log(MTH)^2 +
        vmat[voltable, 7] * log(MTH) * log(N) + vmat[voltable, 8] * log(Vol) * log(N))
}

CalcVol <- function(MTH, BA, N, voltable, vmat) {
  if (BA <= 0 || MTH <= 1.6 || N <= 0) return(0)
  if (voltable %in% c(1, 2)) {
    return(MTH * BA * (vmat[voltable, 1] * (MTH - 1.4)^vmat[voltable, 2] + vmat[voltable, 3]))
  }
  if (voltable %in% c(10, 11)) {
    return(BA * (vmat[voltable, 1] + vmat[voltable, 2] * MTH + vmat[voltable, 3] * N))
  }
  exp(-(vmat[voltable, 1] + vmat[voltable, 2] * log(MTH) +
          vmat[voltable, 4] * log(N) + vmat[voltable, 5] * log(N)^2 +
          vmat[voltable, 6] * log(MTH)^2 + vmat[voltable, 7] * log(MTH) * log(N) -
          log(BA)) / (vmat[voltable, 3] + vmat[voltable, 8] * log(N)))
}

OldAgeCorrection <- function(Age, agez, B) {
  T <- (Age - agez) - 25
  if (T < 0) T <- 0
  1 + 4.350585474 * (1 - exp(-0.001473784 * T))^0.973636099
}

Calcagezero_model <- function(SI, latitude, elevation, height_model, pars) {
  hc <- calcheightcoeff_model(SI, latitude, elevation, height_model, pars)
  -log(-(1 - exp(-hc$ha * 20)) * ((1.4 - 0.25) / (SI - 0.25))^(1 / hc$hb) + 1) / hc$ha
}

approxDBH <- function(D200, P, q, pars) D200 - q * pars$Ds * (D200 - P)

dBA_dN <- function(D200, P, q, N, pars) {
  dp_dN <- pars$dm
  dq_dN <- q * pars$dr2 / N / (log(N) - log(200))
  dD_dN <- -pars$Ds * D200 * dq_dN + pars$Ds * P * dq_dN + pars$Ds * q * dp_dN
  D <- approxDBH(D200, P, q, pars)
  if (D < 0) return(0)
  D * (D + 2 * N * dD_dN)
}

MaxBAStocking <- function(D200, site_effect, SI, Nstart, pars) {
  f <- function(N) {
    q <- pars$dr * (1 + pars$drsi * (SI - 28)) * sign(N - 200) * (abs(log(N) - log(200)))^pars$dr2
    P <- pars$dl + pars$dm * N + pars$dn * site_effect
    dBA_dN(D200, P, q, N, pars)
  }
  f_start <- f(Nstart)
  if (!is.finite(f_start) || f_start >= 0) return(NA_real_)
  A <- 250
  B <- max(Nstart, 260)
  fA <- f(A)
  fB <- f(B)
  cap <- 5000
  while (is.finite(fB) && fA * fB > 0 && B < cap) {
    B <- min(cap, B * 1.5)
    fB <- f(B)
  }
  if (!is.finite(fA) || !is.finite(fB) || fA * fB > 0) return(NA_real_)
  for (j in 1:20) {
    C <- 0.5 * (A + B)
    fC <- f(C)
    if (!is.finite(fC)) break
    if (fA * fC <= 0) {
      B <- C
      fB <- fC
    } else {
      A <- C
      fA <- fC
    }
  }
  0.5 * (A + B)
}

DBHmodel_raw <- function(A200, SI, Age, N, latitude, elevation, height_model, pars) {
  agezero <- Calcagezero_model(SI, latitude, elevation, height_model, pars)
  site_effect <- A200 / pars$da1 - 1
  A <- pars$da1 * (1 + site_effect)
  B <- pars$db2 * (pars$db1 + pars$dbSI * (SI - 28) + pars$dbdia * site_effect + pars$dbsidia * (SI - 28) * site_effect)
  B <- min(B, -0.05)
  if (Age < agezero) return(0)
  D200 <- OldAgeCorrection(Age, agezero, B) * A * ((1 - exp(B * (Age - agezero))) / (1 - exp(B * (30 - agezero))))^pars$dc
  qq <- if (N > 220) (log(N) - log(200))^pars$dr2 else 2 * (log(220) - log(200))^pars$dr2 - (log(242) - log(N))^pars$dr2
  q <- pars$dr * (1 + pars$drsi * (SI - 28)) * qq
  P <- pars$dl + pars$dm * N + pars$dn * site_effect
  D <- D200 - q * log(1 + exp(pars$Ds * (D200 - P)))
  if (N > 250 && dBA_dN(D200, P, q, N, pars) <= 0) {
    Nmax <- MaxBAStocking(D200, site_effect, SI, N, pars)
    q2 <- pars$dr * (1 + pars$drsi * (SI - 28)) * sign(Nmax - 200) * (abs(log(Nmax) - log(200)))^pars$dr2
    P2 <- pars$dl + pars$dm * Nmax + pars$dn * site_effect
    D2 <- D200 - q2 * log(1 + exp(pars$Ds * (D200 - P2)))
    D <- D2 * sqrt(Nmax / N)
  }
  max(D, 0)
}

# VBA Index300 sets fn = DBH300 - DBH after Growth(False): full step loop, mortality, initial stocking.
# R inverts I300 with CalcDBH_cubic → DBHmodel_raw at measured Age and SPH_live (no Growth loop).
# BA300_30 in calcBAfromVol uses VolTab + voltable — if VolTab ≠ FCP, I300 shifts even when MTH/SI match.
CalcA200start <- function(Age, I300, SI, latitude, elevation, height_model, vmat, pars) {
  adjI300 <- I300
  if (bias_young && Age < 6.77) {
    i300adj <- 180.5 * adjI300^(-3.256) * (Age - 6.77)^2
    adjI300 <- adjI300 + min(i300adj, 5)
  }
  if (bias_SI) {
    if (SI < 25 && SI >= 15) adjI300 <- adjI300 * (30 - 0.02 * (25 - SI) * (Age - 28.6)) / 30
    if (SI < 15) adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
    if (SI > 35 && SI <= 45) adjI300 <- adjI300 * (30 - 0.02 * (SI - 35) * (Age - 28.6)) / 30
    if (SI > 45) adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
  }
  if (Age < age300) adjI300 <- adjI300 * (age300 + driftO * (Age - 28.6)) / age300
  MTH30 <- CalcMTH_model(SI, age300, latitude, elevation, height_model, pars)
  BA300_30 <- calcBAfromVol(MTH30, adjI300 * age300, 300, voltable, vmat)
  DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
  f <- function(A200) DBHmodel_raw(A200, SI, DBHcalage, 300, latitude, elevation, height_model, pars) - DBH300_30
  a <- 10
  b <- 150
  fa <- f(a)
  # Match Module1 CalcA200 = Bisection(10, 150, 20, 4, SI, Age, stock, DBH) — 20 bisection steps.
  for (i in seq_len(20L)) {
    m <- 0.5 * (a + b)
    fm <- f(m)
    if (fa * fm <= 0) {
      b <- m
    } else {
      a <- m
      fa <- fm
    }
  }
  0.5 * (a + b)
}

CalcDBH_cubic <- function(I300, SI, Age, N, latitude, elevation, height_model, vmat, pars) {
  if (Age <= 20 || Age >= 40) {
    A200 <- CalcA200start(Age, I300, SI, latitude, elevation, height_model, vmat, pars)
    return(DBHmodel_raw(A200, SI, Age, N, latitude, elevation, height_model, pars))
  }
  DBH1 <- DBHmodel_raw(CalcA200start(19.5, I300, SI, latitude, elevation, height_model, vmat, pars), SI, 19.5, N, latitude, elevation, height_model, pars)
  DBH2 <- DBHmodel_raw(CalcA200start(20.5, I300, SI, latitude, elevation, height_model, vmat, pars), SI, 20.5, N, latitude, elevation, height_model, pars)
  DBH3 <- DBHmodel_raw(CalcA200start(39.5, I300, SI, latitude, elevation, height_model, vmat, pars), SI, 39.5, N, latitude, elevation, height_model, pars)
  DBH4 <- DBHmodel_raw(CalcA200start(40.5, I300, SI, latitude, elevation, height_model, vmat, pars), SI, 40.5, N, latitude, elevation, height_model, pars)
  Y0 <- (DBH1 + DBH2) / 2
  Y1 <- (DBH3 + DBH4) / 2
  Y0p <- (DBH2 - DBH1)
  Y1p <- (DBH4 - DBH3)
  A <- Y0
  B <- Y0p
  D <- (2 * (Y0 + Y0p * 20 - Y1) + 20 * (Y1p - Y0p)) / (20^3)
  C <- (Y1p - Y0p - 3 * D * 20^2) / (2 * 20)
  A + B * (Age - 20) + C * (Age - 20)^2 + D * (Age - 20)^3
}

bisection <- function(f, lo, hi, niterations) {
  xA <- lo
  FA <- f(xA)
  xB <- hi
  xC <- NA_real_
  for (j in seq_len(niterations)) {
    xC <- (xA + xB) / 2
    FC <- f(xC)
    if (FA * FC < 0) {
      xB <- xC
    } else {
      xA <- xC
      FA <- FC
    }
  }
  xC
}

solve_SI_from_MTH <- function(MTH_obs, Age, latitude, elevation, height_model, pars) {
  f <- function(SI) MTH_obs - CalcMTH_model(SI, Age, latitude, elevation, height_model, pars)
  bisection(f, 5, 60, 15L)
}

solve_I300_from_DBH <- function(DBH_obs_cm, SI, T, N_at_T, latitude, elevation, height_model, vmat, pars) {
  g <- function(I300) CalcDBH_cubic(I300, SI, Age = T, N = N_at_T, latitude, elevation, height_model, vmat, pars) - DBH_obs_cm
  bisection(g, 1.328, 60, 14L)
}

parse_excel_date <- function(x) {
  if (inherits(x, "Date")) return(x)
  if (is.numeric(x)) return(as.Date(x, origin = "1899-12-30"))
  as.Date(x)
}

# Per-plot site file (standalone workbook; first sheet; one row per Plot). Not read from input.xlsx.
PLOT_SITE_SCHEMA <- c(
  "Plot",
  "Species",
  "Year_planted",
  "Latitude_(decimal_degrees)",
  "Elevation_above_sea_level_(m)",
  "Needle_retention_score",
  "Soil_%C",
  "Soil_%N",
  "Soil_Organic_P_(mg/kg)",
  "Early_survival_(%)",
  "Mean_Temperature_(°C)"
)

normalize_plot_site_header <- function(x) {
  x <- as.character(x)
  x <- gsub("\u00b0", "deg", x, fixed = TRUE)
  x <- tolower(trimws(gsub("\\s+", " ", x)))
  x <- gsub("%", "pct", x, fixed = TRUE)
  gsub("[^a-z0-9]+", "", x)
}

match_plot_site_column <- function(required_label, have_names) {
  req_exact <- trimws(required_label)
  w <- which(tolower(have_names) == tolower(req_exact))
  if (length(w)) return(w[1])
  nr <- normalize_plot_site_header(required_label)
  nh <- vapply(have_names, normalize_plot_site_header, "")
  w <- which(nh == nr)
  if (length(w)) w[1] else NA_integer_
}

# First tree column whose header matches any of the labels (for pulling plot-level attributes from the ind-tree file).
find_tree_column <- function(have_names, labels) {
  for (lb in labels) {
    j <- match_plot_site_column(lb, have_names)
    if (!is.na(j)) return(have_names[j])
  }
  NA_character_
}

first_finite_numeric_in_rows <- function(rows_df, colnm) {
  if (is.na(colnm) || !colnm %in% names(rows_df)) return(NA_real_)
  x <- suppressWarnings(as.numeric(rows_df[[colnm]]))
  x <- x[is.finite(x)]
  if (length(x)) x[[1]] else NA_real_
}

first_non_empty_char_in_rows <- function(rows_df, colnm) {
  if (is.na(colnm) || !colnm %in% names(rows_df)) return(NA_character_)
  x <- trimws(as.character(rows_df[[colnm]]))
  x <- x[!is.na(x) & nzchar(x)]
  if (length(x)) x[[1]] else NA_character_
}

plot_site_cell_is_blank <- function(x) {
  if (length(x) != 1L) return(TRUE)
  v <- x[[1]]
  if (is.na(v)) return(TRUE)
  if (is.numeric(v) && !is.finite(v)) return(TRUE)
  if (is.character(v)) {
    s <- trimws(as.character(v))
    return(!nzchar(s) || tolower(s) %in% c("na", "n/a", ".", "-"))
  }
  FALSE
}

# Excel often stores NZTM Easting/Northing as text with thousands separators; as.numeric() alone is NA.
parse_nztm_number <- function(x) {
  if (length(x) == 0L) return(numeric(0))
  if (is.numeric(x)) return(suppressWarnings(as.numeric(x)))
  s <- trimws(gsub(",", "", gsub("\u00a0", "", as.character(x))))
  suppressWarnings(as.numeric(s))
}

# NOTE: FCP / Module1 height sheets use positive southern latitude (~30–48°), not WGS84 −φ (see Cells(3,6), heightmod).
# When reading lat from Excel, sf, or joins, pass through latitude_for_vba_model() before height / SI / 300 Index.
latitude_for_vba_model <- function(lat_dd) {
  x <- suppressWarnings(as.numeric(lat_dd))
  ifelse(is.finite(x), abs(x), NA_real_)
}

# Site workbook is plot-level: one row per Plot. Extra rows (e.g. repeated meas keys) are dropped (first kept).
dedupe_plot_site_rows <- function(raw, plot_col_idx) {
  pid <- trimws(as.character(raw[[plot_col_idx]]))
  d <- duplicated(pid, fromLast = FALSE)
  if (!any(d)) return(raw)
  warning(
    "Plot site data: removed ", sum(d), " duplicate row(s) for the same Plot (kept first row per plot). ",
    "This table is plot-level, not one row per measurement."
  )
  raw[!d, , drop = FALSE]
}

read_tree_workbook_for_plot_enrich <- function(tree_file, tree_sheet = 1L) {
  dat <- as.data.frame(read_excel(tree_file, sheet = tree_sheet, .name_repair = "minimal"), stringsAsFactors = FALSE)
  nms <- trimws(names(dat))
  nms <- gsub("\\s+", " ", nms)
  names(dat) <- nms
  names(dat)[names(dat) == "Meas Date"] <- "MeasDate"
  if (!"Plot_Id" %in% names(dat)) stop("Tree file must contain column Plot_Id")
  if (!"Planted" %in% names(dat)) stop("Tree file must contain column Planted")
  dat$MeasDate <- parse_excel_date(dat$MeasDate)
  dat$PlantedDate <- parse_excel_date(dat$Planted)
  dat
}

# Read operator-maintained Plot_site_data.xlsx (do not overwrite). Fill only blank cells from
# tree measurements and from E/N + DEM (terra) where lat/elev still missing. Returns completed data.frame.
complete_plot_site_data <- function(
  plot_site_path,
  tree_file,
  tree_sheet = 1L,
  dem_path = "",
  dem_xy_crs = "EPSG:2193",
  sheet = 1L,
  site_rasters = list(soil_c = "", soil_n = "", soil_p = "", mean_temp = "")
) {
  raw <- as.data.frame(read_excel(plot_site_path, sheet = sheet, .name_repair = "minimal"), stringsAsFactors = FALSE)
  names(raw) <- trimws(names(raw))
  idx <- vapply(PLOT_SITE_SCHEMA, function(rq) match_plot_site_column(rq, names(raw)), integer(1))
  if (any(is.na(idx))) {
    miss <- PLOT_SITE_SCHEMA[is.na(idx)]
    stop("Plot site file missing required columns: ", paste(miss, collapse = "; "), "\nFile: ", plot_site_path)
  }
  raw <- dedupe_plot_site_rows(raw, idx[1])
  plots_chk <- trimws(as.character(raw[[idx[1]]]))
  if (any(plots_chk == "", na.rm = TRUE) || any(is.na(plots_chk))) {
    stop("Plot site data: blank Plot in file: ", plot_site_path)
  }
  en_east <- match_plot_site_column("Easting", names(raw))
  en_north <- match_plot_site_column("Northing", names(raw))

  dat <- read_tree_workbook_for_plot_enrich(tree_file, tree_sheet)
  sp_col <- find_tree_column(names(dat), c("Species", "species", "SP", "Sp"))
  lat_col <- find_tree_column(
    names(dat),
    c(
      "Latitude_(decimal_degrees)", "Latitude (decimal degrees)", "Latitude", "Lat", "GPSLat",
      "Latitude_dd", "Y_lat"
    )
  )
  elev_col <- find_tree_column(
    names(dat),
    c(
      "Elevation_above_sea_level_(m)", "Elevation above sea level (m)", "Altitude", "Elevation", "Elev", "RL",
      "Elevation_m"
    )
  )
  east_col <- find_tree_column(names(dat), c("Easting", "East", "NZTM_E", "X_Easting"))
  north_col <- find_tree_column(names(dat), c("Northing", "North", "NZTM_N", "Y_Northing"))
  needle_col <- find_tree_column(names(dat), c("Needle_retention_score", "Needle retention score", "Needle retention", "NRS"))
  soilc_col <- find_tree_column(names(dat), c("Soil_%C", "Soil %C", "Soil C", "SoilC", "Soil_pct_C"))
  soiln_col <- find_tree_column(names(dat), c("Soil_%N", "Soil %N", "Soil N", "SoilN", "Soil_pct_N"))
  soilp_col <- find_tree_column(
    names(dat),
    c("Soil_Organic_P_(mg/kg)", "Soil Organic P (mg/kg)", "Soil P", "Soil_Organic_P", "Organic P")
  )
  surv_col <- find_tree_column(names(dat), c("Early_survival_(%)", "Early survival (%)", "Early survival", "Survival"))
  temp_col <- find_tree_column(names(dat), c("Mean_Temperature_(°C)", "Mean Temperature (°C)", "Mean temp", "MAT", "Temperature"))

  nm <- function(j) names(raw)[idx[j]]

  for (k in seq_len(nrow(raw))) {
    plot_id <- trimws(as.character(raw[[idx[1]]][k]))
    sub <- dat[trimws(as.character(dat$Plot_Id)) == plot_id, , drop = FALSE]
    if (nrow(sub) < 1) next

    if (plot_site_cell_is_blank(raw[[idx[2]]][k])) {
      v <- first_non_empty_char_in_rows(sub, sp_col)
      if (nzchar(v) && !is.na(v)) raw[[nm(2)]][k] <- v
    }
    if (plot_site_cell_is_blank(raw[[idx[3]]][k])) {
      yp <- sub$PlantedDate[1]
      ypv <- if (inherits(yp, "Date")) as.POSIXlt(yp)$year + 1900L else suppressWarnings(as.numeric(yp))
      if (is.finite(ypv)) raw[[nm(3)]][k] <- ypv
    }
    fill_num <- function(j, tree_col) {
      if (plot_site_cell_is_blank(raw[[idx[j]]][k])) {
        v <- first_finite_numeric_in_rows(sub, tree_col)
        if (is.finite(v)) raw[[idx[j]]][k] <- v
      }
    }
    fill_num(4L, lat_col)
    fill_num(5L, elev_col)
    fill_num(6L, needle_col)
    fill_num(7L, soilc_col)
    fill_num(8L, soiln_col)
    fill_num(9L, soilp_col)
    fill_num(10L, surv_col)
    fill_num(11L, temp_col)

    if (!is.na(en_east) && plot_site_cell_is_blank(raw[[en_east]][k])) {
      v <- first_finite_numeric_in_rows(sub, east_col)
      if (is.finite(v)) raw[[names(raw)[en_east]]][k] <- v
    }
    if (!is.na(en_north) && plot_site_cell_is_blank(raw[[en_north]][k])) {
      v <- first_finite_numeric_in_rows(sub, north_col)
      if (is.finite(v)) raw[[names(raw)[en_north]]][k] <- v
    }
  }

  lat_j <- idx[4]
  elev_j <- idx[5]
  tmp <- data.frame(
    Easting = if (!is.na(en_east)) parse_nztm_number(raw[[en_east]]) else rep(NA_real_, nrow(raw)),
    Northing = if (!is.na(en_north)) parse_nztm_number(raw[[en_north]]) else rep(NA_real_, nrow(raw)),
    # NOTE: raw lat may be WGS84 negative; fill_lat_elev + latitude_for_vba_model normalize to VBA convention.
    Latitude_dd = suppressWarnings(as.numeric(raw[[lat_j]])),
    Elevation_m = suppressWarnings(as.numeric(raw[[elev_j]])),
    Longitude_dd = NA_real_,
    stringsAsFactors = FALSE
  )
  tmp <- fill_lat_elev_from_en_dem(tmp, dem_path = dem_path, xy_crs = dem_xy_crs)
  raw[[names(raw)[lat_j]]] <- latitude_for_vba_model(tmp$Latitude_dd)
  raw[[names(raw)[elev_j]]] <- tmp$Elevation_m

  raw <- fill_plot_site_numeric_blanks_from_rasters(
    raw, idx, en_east, en_north, dem_xy_crs, site_rasters
  )

  raw
}

site_bundle_from_dataframe <- function(raw_df) {
  raw <- as.data.frame(raw_df, stringsAsFactors = FALSE)
  names(raw) <- trimws(names(raw))
  idx <- vapply(PLOT_SITE_SCHEMA, function(rq) match_plot_site_column(rq, names(raw)), integer(1))
  if (any(is.na(idx))) {
    stop("Completed plot site data.frame missing required columns: ", paste(PLOT_SITE_SCHEMA[is.na(idx)], collapse = "; "))
  }
  raw <- dedupe_plot_site_rows(raw, idx[1])
  plots <- trimws(as.character(raw[[idx[1]]]))
  if (any(plots == "", na.rm = TRUE) || any(is.na(plots))) stop("Blank Plot in completed site data")
  en_east <- match_plot_site_column("Easting", names(raw))
  en_north <- match_plot_site_column("Northing", names(raw))
  list(raw = raw, col_idx = idx, en_east = en_east, en_north = en_north)
}

read_plot_site_table <- function(path, sheet = 1L) {
  if (!file.exists(path)) {
    stop("Plot site file not found: ", path)
  }
  raw <- as.data.frame(read_excel(path, sheet = sheet, .name_repair = "minimal"), stringsAsFactors = FALSE)
  names(raw) <- trimws(names(raw))
  idx <- vapply(PLOT_SITE_SCHEMA, function(rq) match_plot_site_column(rq, names(raw)), integer(1))
  if (any(is.na(idx))) {
    miss <- PLOT_SITE_SCHEMA[is.na(idx)]
    stop(
      "Plot site file must contain these columns (headers may differ slightly): ",
      paste(miss, collapse = "; "),
      "\nFile: ", path
    )
  }
  raw <- dedupe_plot_site_rows(raw, idx[1])
  plots <- trimws(as.character(raw[[idx[1]]]))
  if (any(plots == "", na.rm = TRUE) || any(is.na(plots))) {
    stop("Plot site data: blank Plot in file: ", path)
  }
  en_east <- match_plot_site_column("Easting", names(raw))
  en_north <- match_plot_site_column("Northing", names(raw))
  list(raw = raw, col_idx = idx, en_east = en_east, en_north = en_north)
}

# Run once in RStudio to create the standalone plot-site workbook (same columns the script expects).
write_plot_site_data_template <- function(path) {
  nms <- c(PLOT_SITE_SCHEMA, "Easting", "Northing")
  tpl <- stats::setNames(as.data.frame(matrix(ncol = length(nms), nrow = 0)), nms)
  write_xlsx(list(Plot_site_data = tpl), path = path)
  message("Wrote template: ", path)
  invisible(path)
}

ensure_site_data_columns <- function(df) {
  if (!"Longitude_dd" %in% names(df)) df$Longitude_dd <- NA_real_
  if (!"Latitude_dd" %in% names(df)) df$Latitude_dd <- NA_real_
  if (!"Elevation_m" %in% names(df)) df$Elevation_m <- NA_real_
  if (!"Easting" %in% names(df)) df$Easting <- NA_real_
  if (!"Northing" %in% names(df)) df$Northing <- NA_real_
  for (nm in c(
    "site_Species", "site_Year_planted", "site_Needle_retention_score",
    "site_Soil_pct_C", "site_Soil_pct_N", "site_Soil_Organic_P_mg_kg",
    "site_Early_survival_pct", "site_Mean_Temperature_C"
  )) {
    if (!nm %in% names(df)) df[[nm]] <- if (nm == "site_Species") NA_character_ else NA_real_
  }
  df
}

# WGS84 latitude/longitude (decimal degrees) and elevation (m) from NZTM-style Easting/Northing + DEM GeoTIFF.
fill_lat_elev_from_en_dem <- function(df, dem_path = "", xy_crs = "EPSG:2193") {
  df <- ensure_site_data_columns(df)
  df$Easting <- parse_nztm_number(df$Easting)
  df$Northing <- parse_nztm_number(df$Northing)
  en_ok <- is.finite(df$Easting) & is.finite(df$Northing)
  need_lat <- !is.finite(df$Latitude_dd) & en_ok
  if (any(need_lat) && requireNamespace("sf", quietly = TRUE)) {
    ix <- which(need_lat)
    pts <- sf::st_as_sf(
      data.frame(E = df$Easting[ix], N = df$Northing[ix]),
      coords = c("E", "N"),
      crs = xy_crs
    )
    ll <- sf::st_transform(pts, 4326)
    crd <- sf::st_coordinates(ll)
    df$Longitude_dd[ix] <- crd[, 1]
    df$Latitude_dd[ix] <- latitude_for_vba_model(crd[, 2])
  }
  need_z <- !is.finite(df$Elevation_m) & en_ok
  if (any(need_z) && nzchar(dem_path) && requireNamespace("terra", quietly = TRUE)) {
    ix <- which(need_z)
    r <- terra::rast(dem_path)
    xy <- data.frame(Easting = df$Easting[ix], Northing = df$Northing[ix])
    pts <- terra::vect(xy, geom = c("Easting", "Northing"), crs = xy_crs)
    pts <- tryCatch(terra::project(pts, terra::crs(r)), error = function(e) NULL)
    if (!is.null(pts)) {
      ex <- terra::extract(r, pts, fun = mean, na.rm = TRUE)
      band_nm <- setdiff(names(ex), "ID")
      vals <- if (length(band_nm)) ex[[band_nm[1]]] else ex[[2]]
      df$Elevation_m[ix] <- suppressWarnings(as.numeric(vals))
    }
  }
  if ("Latitude_dd" %in% names(df)) {
    df$Latitude_dd <- latitude_for_vba_model(df$Latitude_dd)
  }
  df
}

# Single-band (or first band) raster value at NZTM Easting/Northing; NA where coords missing or extract fails.
extract_raster_values_at_en <- function(east, north, raster_path, xy_crs = "EPSG:2193") {
  n <- length(east)
  out <- rep(NA_real_, n)
  if (length(north) != n) stop("east and north must have the same length")
  rp <- as.character(raster_path)[1]
  if (!nzchar(rp) || !requireNamespace("terra", quietly = TRUE)) return(out)
  ok <- is.finite(east) & is.finite(north)
  if (!any(ok)) return(out)
  r <- tryCatch(terra::rast(rp), error = function(e) NULL)
  if (is.null(r)) return(out)
  xy <- data.frame(Easting = east[ok], Northing = north[ok])
  pts <- tryCatch(
    terra::project(terra::vect(xy, geom = c("Easting", "Northing"), crs = xy_crs), terra::crs(r)),
    error = function(e) NULL
  )
  if (is.null(pts)) return(out)
  ex <- tryCatch(terra::extract(r, pts, fun = mean, na.rm = TRUE), error = function(e) NULL)
  if (is.null(ex)) return(out)
  band_nm <- setdiff(names(ex), "ID")
  vals <- if (length(band_nm)) ex[[band_nm[1]]] else ex[[2]]
  vals <- suppressWarnings(as.numeric(vals))
  out[which(ok)] <- vals
  out
}

# Fill blank Soil_%C, Soil_%N, Soil_Organic_P, Mean_Temperature from GeoTIFFs (same CRS workflow as DEM).
fill_plot_site_numeric_blanks_from_rasters <- function(raw, idx, en_east, en_north, xy_crs, site_rasters) {
  if (is.null(site_rasters) || !length(site_rasters)) return(raw)
  east <- if (!is.na(en_east)) parse_nztm_number(raw[[en_east]]) else rep(NA_real_, nrow(raw))
  north <- if (!is.na(en_north)) parse_nztm_number(raw[[en_north]]) else rep(NA_real_, nrow(raw))
  col_by_key <- list(soil_c = 7L, soil_n = 8L, soil_p = 9L, mean_temp = 11L)
  for (key in names(col_by_key)) {
    path <- site_rasters[[key]]
    if (is.null(path)) next
    path <- trimws(as.character(path)[1])
    if (!nzchar(path)) next
    j <- as.integer(idx[[col_by_key[[key]]]][1L])
    vals <- extract_raster_values_at_en(east, north, path, xy_crs)
    for (k in seq_len(nrow(raw))) {
      if (plot_site_cell_is_blank(raw[[j]][k]) && is.finite(vals[k])) raw[[j]][k] <- vals[k]
    }
  }
  raw
}

append_geography_validation_failures <- function(df) {
  for (i in seq_len(nrow(df))) {
    reasons <- character(0)
    if (!is.finite(df$Latitude_dd[i])) reasons <- c(reasons, "missing_latitude_after_plot_data_and_dem")
    if (!is.finite(df$Elevation_m[i])) reasons <- c(reasons, "missing_elevation_after_plot_data_and_dem")
    if (length(reasons) == 0) next
    prev <- as.character(df$calc_failure_reason[i])
    df$calc_failed[i] <- TRUE
    df$calc_failure_reason[i] <- paste(unique(c(
      if (nzchar(prev) && !is.na(prev)) strsplit(prev, ";", fixed = TRUE)[[1]] else character(0),
      reasons
    )), collapse = ";")
  }
  df
}

coerce_year_planted_site <- function(v) {
  if (inherits(v, "Date")) return(as.POSIXlt(v)$year + 1900L)
  if (inherits(v, "POSIXct")) return(as.POSIXlt(v)$year + 1900L)
  suppressWarnings(as.numeric(v))
}

# Join strict per-plot site data; marks rows where Plot is missing or any site field invalid.
attach_plot_site_data_strict <- function(summary_df, site_bundle) {
  raw <- site_bundle$raw
  idx <- site_bundle$col_idx
  plot_key <- trimws(as.character(raw[[idx[1]]]))
  n <- nrow(summary_df)
  Latitude_dd <- Elevation_m <- rep(NA_real_, n)
  site_Species <- character(n)
  site_Year_planted <- rep(NA_real_, n)
  site_Needle_retention_score <- rep(NA_real_, n)
  site_Soil_pct_C <- rep(NA_real_, n)
  site_Soil_pct_N <- rep(NA_real_, n)
  site_Soil_Organic_P <- rep(NA_real_, n)
  site_Early_survival_pct <- rep(NA_real_, n)
  site_Mean_Temperature_C <- rep(NA_real_, n)
  Easting <- Northing <- rep(NA_real_, n)

  m <- match(trimws(as.character(summary_df$Plot_Id)), plot_key)

  ok_row <- is.finite(m)
  if (any(ok_row)) {
    ii <- which(ok_row)
    j <- m[ii]
    # NOTE: same positive-latitude convention as FCP_5_2.xlsm / Module1 (see latitude_for_vba_model).
    Latitude_dd[ii] <- latitude_for_vba_model(raw[[idx[4]]][j])
    Elevation_m[ii] <- suppressWarnings(as.numeric(raw[[idx[5]]][j]))
    site_Species[ii] <- as.character(raw[[idx[2]]][j])
    site_Year_planted[ii] <- coerce_year_planted_site(raw[[idx[3]]][j])
    site_Needle_retention_score[ii] <- suppressWarnings(as.numeric(raw[[idx[6]]][j]))
    site_Soil_pct_C[ii] <- suppressWarnings(as.numeric(raw[[idx[7]]][j]))
    site_Soil_pct_N[ii] <- suppressWarnings(as.numeric(raw[[idx[8]]][j]))
    site_Soil_Organic_P[ii] <- suppressWarnings(as.numeric(raw[[idx[9]]][j]))
    site_Early_survival_pct[ii] <- suppressWarnings(as.numeric(raw[[idx[10]]][j]))
    site_Mean_Temperature_C[ii] <- suppressWarnings(as.numeric(raw[[idx[11]]][j]))
    if (!is.na(site_bundle$en_east)) {
      Easting[ii] <- suppressWarnings(as.numeric(raw[[site_bundle$en_east]][j]))
    }
    if (!is.na(site_bundle$en_north)) {
      Northing[ii] <- suppressWarnings(as.numeric(raw[[site_bundle$en_north]][j]))
    }
  }

  out <- summary_df
  out$Latitude_dd <- Latitude_dd
  out$Elevation_m <- Elevation_m
  out$Easting <- Easting
  out$Northing <- Northing
  out$site_Species <- site_Species
  out$site_Year_planted <- site_Year_planted
  out$site_Needle_retention_score <- site_Needle_retention_score
  out$site_Soil_pct_C <- site_Soil_pct_C
  out$site_Soil_pct_N <- site_Soil_pct_N
  out$site_Soil_Organic_P_mg_kg <- site_Soil_Organic_P
  out$site_Early_survival_pct <- site_Early_survival_pct
  out$site_Mean_Temperature_C <- site_Mean_Temperature_C

  for (i in seq_len(n)) {
    reasons <- character(0)
    if (is.na(m[i])) {
      reasons <- c(reasons, "plot_not_in_Plot_site_data")
    } else {
      if (!is.finite(out$site_Year_planted[i]) || out$site_Year_planted[i] < 1800 || out$site_Year_planted[i] > 2100) {
        reasons <- c(reasons, "missing_or_invalid_site_Year_planted")
      }
    }
    if (length(reasons) > 0) {
      prev <- as.character(out$calc_failure_reason[i])
      out$calc_failed[i] <- TRUE
      out$calc_failure_reason[i] <- paste(unique(c(
        if (nzchar(prev) && !is.na(prev)) strsplit(prev, ";", fixed = TRUE)[[1]] else character(0),
        reasons
      )), collapse = ";")
    }
  }
  out
}

# Optional: left-join NZFM summary metadata for plot_level export columns only (not used to compute SI / 300I).
attach_summary_metadata <- function(df, summary_file = NULL) {
  if (is.null(summary_file) || !nzchar(summary_file)) {
    return(df)
  }
  ref <- as.data.frame(read_excel(summary_file, .name_repair = "minimal"))
  names(ref) <- gsub("\\s+", " ", trimws(names(ref)))
  names(df) <- gsub("\\s+", " ", trimws(names(df)))
  ref$Plot_id <- trimws(as.character(ref$Plot_id))
  ref$Meas_date <- parse_excel_date(ref$Meas_date)
  df$Plot_Id <- trimws(as.character(df$Plot_Id))
  df$MeasDate <- parse_excel_date(df$MeasDate)
  meta_cols <- intersect(
    c("Forest", "Cpt", "Stand", "Plot Status", "Species", "Plant_date", "Plot_size", "Easting", "Northing"),
    names(ref)
  )
  meta_cols <- setdiff(meta_cols, names(df))
  key <- ref[, c("Plot_id", "Meas_date", meta_cols), drop = FALSE]
  names(key)[1:2] <- c("Plot_Id", "MeasDate")
  out <- merge(df, key, by = c("Plot_Id", "MeasDate"), all.x = TRUE, sort = FALSE)
  out <- out[match(paste(df$Plot_Id, df$MeasDate), paste(out$Plot_Id, out$MeasDate)), ]
  rownames(out) <- NULL
  out
}

# Column order and names matching NZFM summary export (plus trailing QC columns).
REFERENCE_SUMMARY_COLS <- c(
  "Plot_id", "Forest", "Cpt", "Stand", "Plot Status", "Site Index", "Index 300_500",
  "Easting", "Northing", "Species", "Plant_date", "Plot_size", "Meas_date", "Age",
  "SPH_total", "SPH_dead", "SPH_wind", "SPH_b4_thn", "SPH_thin", "SPH_live",
  "MnDBH_dead", "MnDBH_wind", "MnDBH_b4th", "MnDBH_CAI", "MnDBH_thin", "MnDBH_live",
  "Mean_ht", "Mean_CrHt", "MTH", "BA_total", "BA_dead", "BA_wind", "BA_b4_thin",
  "BA_net_CAI", "BA_grs_CAI", "BA_thin", "BA_live", "Vol_total", "Vol_dead", "Vol_wind",
  "Vol_b4_thn", "Vol_netCAI", "Vol_grsCAI", "Vol_thin", "Vol_live"
)

layout_output_like_reference <- function(df) {
  df$`Site Index` <- df$SI
  df$`Index 300_500` <- df$Index300
  df$Plot_id <- df$Plot_Id
  df$Meas_date <- df$MeasDate
  df$Age <- df$Age_years
  if (!"SPH_thin" %in% names(df)) df$SPH_thin <- 0
  n <- nrow(df)
  out <- as.data.frame(
    stats::setNames(
      replicate(length(REFERENCE_SUMMARY_COLS), rep(NA, n), simplify = FALSE),
      REFERENCE_SUMMARY_COLS
    ),
    stringsAsFactors = FALSE
  )
  for (nm in REFERENCE_SUMMARY_COLS) {
    if (nm %in% names(df)) out[[nm]] <- df[[nm]]
  }
  if ("calc_failed" %in% names(df)) out$calc_failed <- df$calc_failed
  if ("calc_failure_reason" %in% names(df)) out$calc_failure_reason <- df$calc_failure_reason
  out
}

## Diagnostic plotting helpers removed.

## FCP comparison helpers removed.

## NZFM comparison helpers removed.

calc_mth_from_mean_ht <- function(mean_ht, stocking_sph, pars) {
  pick_first <- function(nms) {
    for (nm in nms) {
      if (!is.null(pars[[nm]]) && is.finite(as.numeric(pars[[nm]]))) return(as.numeric(pars[[nm]]))
    }
    NA_real_
  }
  a <- pick_first(c("MTH_MnHt_a_rad", "MTH_MnHt_a"))
  b <- pick_first(c("MTH_MnHt_b_rad", "MTH_MnHt_b"))
  if (!is.finite(mean_ht) || !is.finite(stocking_sph) || stocking_sph <= 0) return(NA_real_)
  if (!is.finite(a) || !is.finite(b)) return(mean_ht)
  mean_ht / (1 - a * (1 - exp(b * (stocking_sph - 100))))
}

calc_mtd <- function(dbh_cm, stocking_sph) {
  nstems <- length(dbh_cm)
  if (nstems < 1 || !is.finite(stocking_sph) || stocking_sph <= 0) return(NA_real_)
  nMTD <- 100
  sumWt <- 0
  sumDBH2Wt <- 0
  sorted <- sort(dbh_cm, decreasing = FALSE)
  wt_per_tree <- stocking_sph / nstems
  for (j in nstems:1) {
    if (sumWt + wt_per_tree > nMTD) {
      wt <- nMTD - sumWt
    } else {
      wt <- wt_per_tree
    }
    sumWt <- sumWt + wt
    sumDBH2Wt <- sumDBH2Wt + wt * sorted[j]^2
    if (sumWt >= nMTD) break
  }
  sqrt(sumDBH2Wt / sumWt)
}

fit_petterson_type1 <- function(dbh_cm, ht_m) {
  valid <- is.finite(dbh_cm) & dbh_cm > 0 & is.finite(ht_m) & ht_m > 1.4
  dbh <- dbh_cm[valid]
  ht <- ht_m[valid]
  n <- length(dbh)
  if (n < 2) return(list(petA = NA_real_, petB = NA_real_))
  X <- dbh
  Y <- dbh / (ht - 1.4)^0.4
  sum_x <- sum(X); sum_y <- sum(Y); sum_x2 <- sum(X^2); sum_xy <- sum(X * Y)
  petA <- (sum_xy - sum_x * sum_y / n) / (sum_x2 - sum_x^2 / n)
  petB <- sum_y / n - petA * (sum_x / n)
  if (petB < 0) {
    petB <- 0
    petA <- sum_y / sum_x
  }
  if (petA < 0) {
    petA <- 0
    petB <- sum_y / n
  }
  list(petA = petA, petB = petB)
}

summarise_plot_year_from_trees <- function(input_file = "NZFM Fert Trial Ind Tree Data as at Nov25.xlsx", sheet = 1, model_inputs) {
  vmat <- model_inputs$v
  dat <- as.data.frame(read_excel(input_file, sheet = sheet, .name_repair = "minimal"))
  nm <- trimws(names(dat))
  nm <- gsub("\\s+", " ", nm)
  names(dat) <- nm
  to_num <- function(x) {
    x <- gsub(",", "", trimws(as.character(x)))
    suppressWarnings(as.numeric(x))
  }
  names(dat)[names(dat) == "Meas Date"] <- "MeasDate"
  names(dat)[names(dat) == "Tree Sample Ht"] <- "TreeSampleHt"
  names(dat)[names(dat) == "Status"] <- "Status"
  dat$MeasDate <- parse_excel_date(dat$MeasDate)
  dat$PlantedDate <- parse_excel_date(dat$Planted)
  dat$PlotArea_ha <- ifelse(dat$PlotArea > 5, dat$PlotArea / 10000, dat$PlotArea)
  grp <- split(dat, interaction(dat$Plot_Id, dat$MeasDate, drop = TRUE, lex.order = TRUE))
  out <- lapply(grp, function(g) {
    fail_reasons <- character(0)
    add_fail <- function(msg) fail_reasons <<- unique(c(fail_reasons, msg))

    get_mth <- function(df) {
      sampled <- toupper(trimws(as.character(df$TreeSampleHt))) == "Y"
      h <- to_num(df$TotalHt[sampled])
      h <- h[!is.na(h) & h > 0]
      if (length(h) == 0) return(NA_real_)
      mean(h, na.rm = TRUE)
    }
    status <- toupper(trimws(as.character(g$Status)))
    alive_idx <- status == "A"
    dead_idx <- status == "X"
    wind_idx <- status == "W"

    dbh_live <- to_num(g$DBH[alive_idx]); dbh_live <- dbh_live[!is.na(dbh_live) & dbh_live > 0]
    dbh_dead <- to_num(g$DBH[dead_idx]);  dbh_dead <- dbh_dead[!is.na(dbh_dead) & dbh_dead > 0]
    dbh_wind <- to_num(g$DBH[wind_idx]);  dbh_wind <- dbh_wind[!is.na(dbh_wind) & dbh_wind > 0]

    area_ha <- as.numeric(g$PlotArea_ha[1])
    age <- as.numeric(g$MeasDate[1] - g$PlantedDate[1]) / 365.25
    if (!is.finite(area_ha) || area_ha <= 0) add_fail("missing_or_invalid_plot_area")
    if (!is.finite(age) || age <= 0) add_fail("missing_or_invalid_age")
    n_alive <- sum(alive_idx, na.rm = TRUE)
    n_dead <- sum(dead_idx, na.rm = TRUE)
    n_wind <- sum(wind_idx, na.rm = TRUE)
    if (n_alive < 1) add_fail("no_alive_trees")
    SPH_live <- n_alive / area_ha
    SPH_dead <- n_dead / area_ha
    SPH_wind <- n_wind / area_ha
    SPH_total <- SPH_live + SPH_dead + SPH_wind

    MnDBH_live <- mean(dbh_live, na.rm = TRUE)
    MnDBH_dead <- mean(dbh_dead, na.rm = TRUE)
    MnDBH_wind <- mean(dbh_wind, na.rm = TRUE)
    dbh_all <- c(dbh_live, dbh_dead, dbh_wind)
    MnDBH_b4th <- if (length(dbh_all) > 0) mean(dbh_all, na.rm = TRUE) else NA_real_
    qdbh <- sqrt(mean(dbh_live^2, na.rm = TRUE))

    BA_live <- sum(pi * (dbh_live / 200)^2, na.rm = TRUE) / area_ha
    BA_dead <- sum(pi * (dbh_dead / 200)^2, na.rm = TRUE) / area_ha
    BA_wind <- sum(pi * (dbh_wind / 200)^2, na.rm = TRUE) / area_ha
    BA_total <- BA_live + BA_dead + BA_wind
    if (!is.finite(BA_live)) add_fail("missing_or_invalid_ba_live")

    Mean_ht <- NA_real_
    MTH_dead <- get_mth(g[dead_idx, , drop = FALSE])
    MTH_wind <- get_mth(g[wind_idx, , drop = FALSE])

    sampled_alive <- alive_idx & (toupper(trimws(as.character(g$TreeSampleHt))) == "Y")
    dbh_for_mth <- to_num(g$DBH[sampled_alive])
    ht_for_mth <- to_num(g$TotalHt[sampled_alive])
    pair_ok <- is.finite(dbh_for_mth) & dbh_for_mth > 0 & is.finite(ht_for_mth) & ht_for_mth > 1.4
    dbh_for_mth <- dbh_for_mth[pair_ok]
    ht_for_mth <- ht_for_mth[pair_ok]
    if (length(ht_for_mth) < 2 || length(dbh_for_mth) < 2) add_fail("insufficient_tree_sample_height_for_mth")
    MTDia <- calc_mtd(dbh_live, SPH_live)
    pet <- fit_petterson_type1(dbh_for_mth, ht_for_mth)

    # Predict missing alive-tree heights from DBH using fitted Petterson curve.
    dbh_alive_all <- to_num(g$DBH[alive_idx])
    ht_alive_all <- to_num(g$TotalHt[alive_idx])
    can_predict <- is.finite(dbh_alive_all) & dbh_alive_all > 0 & !is.finite(ht_alive_all)
    if (is.finite(pet$petA) && is.finite(pet$petB) && any(can_predict)) {
      ht_alive_all[can_predict] <- 1.4 + (pet$petA + pet$petB / dbh_alive_all[can_predict])^(-2.5)
    }
    ht_alive_valid <- ht_alive_all[is.finite(ht_alive_all) & ht_alive_all > 0]
    if (length(ht_alive_valid) > 0) Mean_ht <- mean(ht_alive_valid)
    if (!is.finite(Mean_ht)) add_fail("failed_mean_height_calculation")

    MTH <- 1.4 + (pet$petA + pet$petB / MTDia)^(-2.5)
    if (!is.finite(MTH)) add_fail("failed_mth_calculation")

    Vol_live <- if (is.finite(MTH) && is.finite(BA_live) && is.finite(SPH_live)) CalcVol(MTH, BA_live, SPH_live, voltable, vmat) else NA_real_
    mth_dead_vol <- if (is.finite(MTH_dead)) MTH_dead else MTH
    mth_wind_vol <- if (is.finite(MTH_wind)) MTH_wind else MTH
    Vol_dead <- if (is.finite(mth_dead_vol) && is.finite(BA_dead) && is.finite(SPH_dead)) CalcVol(mth_dead_vol, BA_dead, SPH_dead, voltable, vmat) else NA_real_
    Vol_wind <- if (is.finite(mth_wind_vol) && is.finite(BA_wind) && is.finite(SPH_wind)) CalcVol(mth_wind_vol, BA_wind, SPH_wind, voltable, vmat) else NA_real_
    vd <- if (is.finite(Vol_dead)) Vol_dead else 0
    vw <- if (is.finite(Vol_wind)) Vol_wind else 0
    Vol_total <- if (is.finite(Vol_live)) Vol_live + vd + vw else NA_real_
    if (!is.finite(Vol_live)) add_fail("failed_volume_calculation")

    yp <- g$PlantedDate[1]
    year_planted <- if (inherits(yp, "Date")) {
      as.POSIXlt(yp)$year + 1900L
    } else {
      suppressWarnings(as.integer(as.numeric(yp)))
    }

    data.frame(
      Plot_Id = as.character(g$Plot_Id[1]),
      MeasDate = g$MeasDate[1],
      Year_planted = year_planted,
      Age_years = age,
      PlotArea_ha = area_ha,
      n_trees = n_alive,
      n_dead = n_dead,
      n_wind = n_wind,
      SPH_total = SPH_total,
      SPH_dead = SPH_dead,
      SPH_wind = SPH_wind,
      SPH_b4_thn = SPH_total,
      SPH_thin = 0,
      SPH_live = SPH_live,
      MnDBH_dead = MnDBH_dead,
      MnDBH_wind = MnDBH_wind,
      MnDBH_b4th = MnDBH_b4th,
      MnDBH_CAI = NA_real_,
      MnDBH_thin = NA_real_,
      MnDBH_live = MnDBH_live,
      Mean_ht = Mean_ht,
      Mean_CrHt = NA_real_,
      qDBH_cm = qdbh,
      MTH = MTH,
      BA_total = BA_total,
      BA_dead = BA_dead,
      BA_wind = BA_wind,
      BA_b4_thin = BA_total,
      BA_net_CAI = NA_real_,
      BA_grs_CAI = NA_real_,
      BA_thin = 0,
      BA_live = BA_live,
      Vol_total = Vol_total,
      Vol_dead = Vol_dead,
      Vol_wind = Vol_wind,
      Vol_b4_thn = Vol_total,
      Vol_netCAI = NA_real_,
      Vol_grsCAI = NA_real_,
      Vol_thin = 0,
      Vol_live = Vol_live,
      Elevation_m = NA_real_,
      calc_failed = length(fail_reasons) > 0,
      calc_failure_reason = paste(fail_reasons, collapse = ";")
    )
  })
  summary_df <- do.call(rbind, out)
  summary_df[order(summary_df$Plot_Id, summary_df$MeasDate), ]
}

solve_si300_for_plot_year <- function(summary_df, model_inputs) {
  pars <- model_inputs$coefs
  vmat <- model_inputs$v
  cfg <- model_inputs$config
  out <- vector("list", nrow(summary_df))
  for (i in seq_len(nrow(summary_df))) {
    fail_reasons <- character(0)
    add_fail <- function(msg) fail_reasons <<- unique(c(fail_reasons, msg))
    age <- as.numeric(summary_df$Age_years[i])
    N <- as.numeric(summary_df$SPH_live[i])
    BA <- as.numeric(summary_df$BA_live[i])
    MTH <- as.numeric(summary_df$MTH[i])
    latitude <- latitude_for_vba_model(summary_df$Latitude_dd[i])
    elevation <- as.numeric(summary_df$Elevation_m[i])
    DBH <- CalcDBHfromBA(BA, N)
    hm <- resolve_height_model_for_stand(latitude, elevation, cfg$height_model)
    # Lat/elev required only for environmental height (3), same as VBA before calcheightcoeff uses them.
    if (hm == 3L) {
      if (!is.finite(latitude)) add_fail("missing_site_latitude")
      if (!is.finite(elevation)) add_fail("missing_site_elevation")
    }
    lat_m <- if (is.finite(latitude)) latitude else 0
    elev_m <- if (is.finite(elevation)) elevation else 0
    if (!is.finite(age) || age <= 0) add_fail("missing_or_invalid_age")
    if (!is.finite(N) || N <= 0) add_fail("missing_or_invalid_stems_per_ha")
    if (!is.finite(BA) || BA <= 0) add_fail("missing_or_invalid_basal_area")
    if (!is.finite(MTH) || MTH <= 0) add_fail("missing_or_invalid_mth")
    if (!is.finite(DBH) || DBH <= 0) add_fail("failed_dbh_from_ba")

    SI <- NA_real_
    I300 <- NA_real_
    if (length(fail_reasons) == 0) {
      SI <- if (isTRUE(all.equal(age, 20))) {
        MTH
      } else {
        solve_SI_from_MTH(MTH, age, lat_m, elev_m, hm, pars)
      }
      if (!is.finite(SI)) add_fail("failed_site_index_calculation")
      if (length(fail_reasons) == 0) {
        I300 <- solve_I300_from_DBH(DBH, SI, age, N, lat_m, elev_m, hm, vmat, pars)
        if (!is.finite(I300)) add_fail("failed_300_index_calculation")
      }
    }
    out[[i]] <- data.frame(
      Plot_Id = as.character(summary_df$Plot_Id[i]),
      MeasDate = summary_df$MeasDate[i],
      Age_years = age,
      PlotArea_ha = as.numeric(summary_df$PlotArea_ha[i]),
      SPH_total = as.numeric(summary_df$SPH_total[i]),
      SPH_dead = as.numeric(summary_df$SPH_dead[i]),
      SPH_wind = as.numeric(summary_df$SPH_wind[i]),
      SPH_b4_thn = as.numeric(summary_df$SPH_b4_thn[i]),
      SPH_thin = as.numeric(summary_df$SPH_thin[i]),
      SPH_live = N,
      MnDBH_dead = as.numeric(summary_df$MnDBH_dead[i]),
      MnDBH_wind = as.numeric(summary_df$MnDBH_wind[i]),
      MnDBH_b4th = as.numeric(summary_df$MnDBH_b4th[i]),
      MnDBH_CAI = as.numeric(summary_df$MnDBH_CAI[i]),
      MnDBH_thin = as.numeric(summary_df$MnDBH_thin[i]),
      MnDBH_live = as.numeric(summary_df$MnDBH_live[i]),
      Mean_ht = as.numeric(summary_df$Mean_ht[i]),
      Mean_CrHt = as.numeric(summary_df$Mean_CrHt[i]),
      qDBH_cm = as.numeric(summary_df$qDBH_cm[i]),
      MTH = MTH,
      BA_total = as.numeric(summary_df$BA_total[i]),
      BA_dead = as.numeric(summary_df$BA_dead[i]),
      BA_wind = as.numeric(summary_df$BA_wind[i]),
      BA_b4_thin = as.numeric(summary_df$BA_b4_thin[i]),
      BA_net_CAI = as.numeric(summary_df$BA_net_CAI[i]),
      BA_grs_CAI = as.numeric(summary_df$BA_grs_CAI[i]),
      BA_thin = as.numeric(summary_df$BA_thin[i]),
      BA_live = BA,
      Vol_total = as.numeric(summary_df$Vol_total[i]),
      Vol_dead = as.numeric(summary_df$Vol_dead[i]),
      Vol_wind = as.numeric(summary_df$Vol_wind[i]),
      Vol_b4_thn = as.numeric(summary_df$Vol_b4_thn[i]),
      Vol_netCAI = as.numeric(summary_df$Vol_netCAI[i]),
      Vol_grsCAI = as.numeric(summary_df$Vol_grsCAI[i]),
      Vol_thin = as.numeric(summary_df$Vol_thin[i]),
      Vol_live = as.numeric(summary_df$Vol_live[i]),
      SI = SI,
      Index300 = I300,
      calc_failed = as.logical(summary_df$calc_failed[i]) || (length(fail_reasons) > 0),
      calc_failure_reason = paste(
        unique(c(
          fail_reasons,
          {
            z <- as.character(summary_df$calc_failure_reason[i])
            if (!is.na(z) && nzchar(z)) z else character(0)
          }
        )),
        collapse = ";"
      )
    )
  }
  out_df <- do.call(rbind, out)
  ord <- order(out_df$Plot_Id, out_df$MeasDate)
  out_df <- out_df[ord, ]
  rownames(out_df) <- NULL
  sp <- split(seq_len(nrow(out_df)), out_df$Plot_Id)
  for (idx in sp) {
    if (length(idx) < 2) next
    for (k in 2:length(idx)) {
      i <- idx[k]
      j <- idx[k - 1]
      dA <- out_df$Age_years[i] - out_df$Age_years[j]
      if (is.finite(dA) && dA > 0) {
        out_df$MnDBH_CAI[i] <- (out_df$MnDBH_live[i] - out_df$MnDBH_live[j]) / dA
        out_df$BA_net_CAI[i] <- (out_df$BA_live[i] - out_df$BA_live[j]) / dA
        out_df$Vol_netCAI[i] <- (out_df$Vol_live[i] - out_df$Vol_live[j]) / dA
        bd <- if (is.finite(out_df$BA_dead[i])) out_df$BA_dead[i] else 0
        bw <- if (is.finite(out_df$BA_wind[i])) out_df$BA_wind[i] else 0
        out_df$BA_grs_CAI[i] <- out_df$BA_net_CAI[i] + (bd + bw) / dA
        vd <- if (is.finite(out_df$Vol_dead[i])) out_df$Vol_dead[i] else 0
        vw <- if (is.finite(out_df$Vol_wind[i])) out_df$Vol_wind[i] else 0
        out_df$Vol_grsCAI[i] <- out_df$Vol_netCAI[i] + (vd + vw) / dA
      }
    }
  }
  out_df
}

merge_site_context_into_out <- function(out_df, summary_df) {
  extras <- unique(c(
    "Year_planted", "Latitude_dd", "Longitude_dd", "Elevation_m", "Easting", "Northing",
    grep("^site_", names(summary_df), value = TRUE)
  ))
  extras <- intersect(extras, names(summary_df))
  if (length(extras) == 0) return(out_df)
  mk <- summary_df[, c("Plot_Id", "MeasDate", extras), drop = FALSE]
  out2 <- merge(out_df, mk, by = c("Plot_Id", "MeasDate"), all.x = TRUE, sort = FALSE)
  out2 <- out2[match(paste(out_df$Plot_Id, out_df$MeasDate), paste(out2$Plot_Id, out2$MeasDate)), ]
  rownames(out2) <- NULL
  out2
}

# One row per Plot: latest measurement carries SI / 300 Index; site covariates from merged context.
build_plot_description <- function(out_df) {
  ord <- order(out_df$Plot_Id, out_df$MeasDate)
  x <- out_df[ord, ]
  pick <- !duplicated(x$Plot_Id, fromLast = TRUE)
  last <- x[pick, ]
  n <- nrow(last)
  species_val <- if ("Species" %in% names(last)) as.character(last$Species) else rep(NA_character_, n)
  if ("site_Species" %in% names(last)) {
    ss <- trimws(as.character(last$site_Species))
    species_val <- ifelse(nzchar(ss) & !is.na(ss), ss, species_val)
  }
  yp <- last$Year_planted
  if ("site_Year_planted" %in% names(last)) {
    yp <- ifelse(is.finite(last$site_Year_planted), last$site_Year_planted, yp)
  }
  data.frame(
    check.names = FALSE,
    Plot = last$Plot_Id,
    Species = species_val,
    `Year planted` = yp,
    `Latitude (decimal degrees)` = last$Latitude_dd,
    `Elevation above sea level (m)` = last$Elevation_m,
    `Needle retention score` = if ("site_Needle_retention_score" %in% names(last)) last$site_Needle_retention_score else NA_real_,
    `Soil %C` = if ("site_Soil_pct_C" %in% names(last)) last$site_Soil_pct_C else NA_real_,
    `Soil %N` = if ("site_Soil_pct_N" %in% names(last)) last$site_Soil_pct_N else NA_real_,
    `Soil Organic P (mg/kg)` = if ("site_Soil_Organic_P_mg_kg" %in% names(last)) last$site_Soil_Organic_P_mg_kg else NA_real_,
    `Early survival (%)` = if ("site_Early_survival_pct" %in% names(last)) last$site_Early_survival_pct else NA_real_,
    `Mean Temperature (°C)` = if ("site_Mean_Temperature_C" %in% names(last)) last$site_Mean_Temperature_C else NA_real_,
    `Mean Outerwood Density (kg/m3)` = NA_real_,
    `Outerwood density assessment age (years)` = NA_real_,
    `Inner ring` = NA_real_,
    `Outer ring` = NA_real_,
    `Site Index` = last$SI,
    `Index 300_500` = last$Index300
  )
}

# One row per plot × measurement. Type M = measurement (allowed codes also include E, TW, TP, P1–P5 when you supply them).
build_measurement_info <- function(out_df) {
  data.frame(
    check.names = FALSE,
    Plot = out_df$Plot_Id,
    Type = "M",
    `Age (years)` = out_df$Age_years,
    `Stocking (stems/ha)` = out_df$SPH_live,
    `BA (m2/ha)` = out_df$BA_live,
    `MTH (m)` = out_df$MTH,
    `Pruned stems (stems/ha)` = NA_real_,
    `Pruned height (m)` = NA_real_,
    `Meas date` = out_df$MeasDate,
    `Site Index` = out_df$SI,
    `Index 300_500` = out_df$Index300
  )
}

# Rounding applied only for Excel/report output (full precision retained in returned data.frame).
round_plot_summary <- function(df) {
  round_map <- list(
    Age = 2,
    Easting = 2,
    Northing = 2,
    SPH_total = 0,
    SPH_dead = 0,
    SPH_wind = 0,
    SPH_b4_thn = 0,
    SPH_thin = 0,
    SPH_live = 0,
    MnDBH_dead = 1,
    MnDBH_wind = 1,
    MnDBH_b4th = 1,
    MnDBH_CAI = 2,
    MnDBH_thin = 1,
    MnDBH_live = 1,
    Mean_ht = 1,
    Mean_CrHt = 1,
    MTH = 1,
    BA_total = 2,
    BA_dead = 2,
    BA_wind = 2,
    BA_b4_thin = 2,
    BA_net_CAI = 2,
    BA_grs_CAI = 2,
    BA_thin = 2,
    BA_live = 2,
    Vol_total = 1,
    Vol_dead = 1,
    Vol_wind = 1,
    Vol_b4_thn = 1,
    Vol_netCAI = 1,
    Vol_grsCAI = 1,
    Vol_thin = 1,
    Vol_live = 1,
    `Site Index` = 1,
    `Index 300_500` = 2
  )

  for (nm in names(round_map)) {
    if (nm %in% names(df)) {
      df[[nm]] <- round(as.numeric(df[[nm]]), round_map[[nm]])
    }
  }
  df
}

# ==============================================================================
# OPERATOR RUN — discrete steps (Source this script in RStudio, or run to a line)
# Math helpers are functions above; orchestration is linear so errors show the step.
# ==============================================================================

# ------------------------------------------------------------------------------
# CONFIG — edit paths and switches (all relative to setwd() above = script directory)
# ------------------------------------------------------------------------------
tree_file <- "NZFM Fert Trial Ind Tree Data as at Nov25.xlsx"
tree_sheet <- 1L
input_file <- "input.xlsx"
plot_site_file <- "Plot_site_data.xlsx"
plot_site_complete_file <- "Plot_site_data_complete.xlsx"
plot_site_sheet <- 1L
rasters_dir <- "rasters"
dem_path <- file.path(rasters_dir, "NZ_DEM.tif")
dem_xy_crs <- "EPSG:2193"
site_raster_paths <- list(
  soil_c = file.path(rasters_dir, "C.tif"),
  soil_n = file.path(rasters_dir, "N.tif"),
  soil_p = file.path(rasters_dir, "P.tif"),
  mean_temp = file.path(rasters_dir, "MAT.tif")
)
output_file <- "Plot_Summary.xlsx"
# "module1" = same V(1:11,1:8) as VBA Module1 Sub voltab(); "input" = VolTab sheet.
voltab_source <- "module1"

# Collect step text; one message() at end so the console shows the full run (RStudio scrollback hides early lines).
pipeline_run_log <- character(0)
pipeline_log <- function(...) {
  pipeline_run_log <<- c(pipeline_run_log, paste0(...))
}

# ------------------------------------------------------------------------------
# STEP 1 — Load model coefficients and volume table (input.xlsx: parameters, VolTab)
# ------------------------------------------------------------------------------
model_inputs <- load_model_inputs(input_file, voltab_source = voltab_source)
voltable <<- as.integer(model_inputs$config$voltable)
driftO <<- as.numeric(model_inputs$config$drift)
bias_old <<- isTRUE(model_inputs$config$bias_old)
bias_young <<- isTRUE(model_inputs$config$bias_young)
bias_SI <<- isTRUE(model_inputs$config$bias_SI)
pipeline_log(
  "STEP 1 OK: parameters loaded; voltab_source=", model_inputs$config$voltab_source,
  ". Config: voltable=", model_inputs$config$voltable,
  ", height_model=", model_inputs$config$height_model,
  ", bias_old=", model_inputs$config$bias_old,
  ", bias_young=", model_inputs$config$bias_young,
  ", bias_SI=", model_inputs$config$bias_SI,
  ", drift=", model_inputs$config$drift
)

# ------------------------------------------------------------------------------
# STEP 2 — Read Plot_site_data.xlsx only (operator master; never overwritten here).
#         Fill blank cells from tree measurements; lat/elev from E/N + DEM; soil C/N/P and mean temp from GeoTIFFs (site_raster_paths) when still blank.
#         Write Plot_site_data_complete.xlsx for review / client; modeling uses this output.
# ------------------------------------------------------------------------------
plot_site_complete_df <- complete_plot_site_data(
  plot_site_path = plot_site_file,
  tree_file = tree_file,
  tree_sheet = tree_sheet,
  dem_path = dem_path,
  dem_xy_crs = dem_xy_crs,
  sheet = plot_site_sheet,
  site_rasters = site_raster_paths
)
write_xlsx(list(Plot_site_data_complete = plot_site_complete_df), path = plot_site_complete_file)
pipeline_log("STEP 2 OK: read ", plot_site_file, "; wrote ", plot_site_complete_file, " (", nrow(plot_site_complete_df), " plots).")

# ------------------------------------------------------------------------------
# STEP 3 — Tree-level → plot×measurement: Petterson + height infill, MTH, BA, SPH,
#         volumes, year planted (from trees only)
# ------------------------------------------------------------------------------
summary_df <- summarise_plot_year_from_trees(tree_file, tree_sheet, model_inputs)
pipeline_log("STEP 3 OK: summary_df (plot×MeasDate) nrow = ", nrow(summary_df), ".")

# ------------------------------------------------------------------------------
# STEP 4 — Use completed plot-site table (in memory); join onto each plot×measurement row
# ------------------------------------------------------------------------------
site_bundle <- site_bundle_from_dataframe(plot_site_complete_df)
summary_df <- attach_plot_site_data_strict(summary_df, site_bundle)
pipeline_log("STEP 4 OK: site covariates from completed plot-site data attached to summary_df.")

# ------------------------------------------------------------------------------
# STEP 5 — Fill missing latitude (WGS84 °) from E/N (sf); elevation (m) from DEM (terra);
#         flag rows still missing lat or elev (calc_failed / reasons)
# ------------------------------------------------------------------------------
summary_df <- fill_lat_elev_from_en_dem(summary_df, dem_path = dem_path, xy_crs = dem_xy_crs)
summary_df <- append_geography_validation_failures(summary_df)
summary_df$Latitude_dd <- latitude_for_vba_model(summary_df$Latitude_dd)
pipeline_log("STEP 5 OK: lat/elev pass complete; rows with calc_failed = ", sum(summary_df$calc_failed, na.rm = TRUE), ".")

# ------------------------------------------------------------------------------
# STEP 6 — Site Index (SI), then 300 Index (same row; uses lat, elev, MTH, BA, age, N)
# ------------------------------------------------------------------------------
summary_enriched <- summary_df
out_df <- solve_si300_for_plot_year(summary_enriched, model_inputs)
out_df <- merge_site_context_into_out(out_df, summary_enriched)
pipeline_log(
  "STEP 6 OK: nrow = ", nrow(out_df),
  "; rows with finite SI = ", sum(is.finite(out_df$SI)),
  "; finite Index300 = ", sum(is.finite(out_df$Index300)),
  ". If zero, inspect calc_failure_reason and input.xlsx coefficients (ha0, hae0, …)."
)

# ------------------------------------------------------------------------------
# STEP 7 — Operator tables: plot_description (one row per plot) and measurement_info
# ------------------------------------------------------------------------------
plot_description <- build_plot_description(out_df)
measurement_info <- build_measurement_info(out_df)
pipeline_log("STEP 7 OK: plot_description nrow = ", nrow(plot_description), "; measurement_info nrow = ", nrow(measurement_info), ".")

# ------------------------------------------------------------------------------
# STEP 8 — Reference-layout plot_level sheet + write Plot_Summary.xlsx (3 sheets)
# ------------------------------------------------------------------------------
out_export <- layout_output_like_reference(out_df)
write_xlsx(
  list(
    plot_description = plot_description,
    measurement_info = measurement_info,
    plot_level = round_plot_summary(out_export)
  ),
  path = output_file
)
pipeline_log("STEP 8 OK: wrote ", output_file, ".")

plot_summary_result <- list(
  plot_site_complete = plot_site_complete_df,
  plot_description = plot_description,
  measurement_info = measurement_info,
  plot_level = layout_output_like_reference(out_df),
  out_df = out_df,
  summary_df = summary_enriched
)
pipeline_log("All steps finished.")
message(paste(pipeline_run_log, collapse = "\n"))
