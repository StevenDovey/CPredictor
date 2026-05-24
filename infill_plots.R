# =============================================================================
# infill_plots.R
#
# Standalone plot-metric infill script. Computes plot-level descriptors
# (Site Index, 300 Index, drift stats) from psp_summary measurements and
# writes plots_infilled.csv.
#
# Designed to be run independently of the yield/carbon batch. The infilled
# CSV becomes the input for downstream analyses.
#
# Behaviour:
#   - For each plot, checks Site Index and 300 Index columns
#   - If BOTH present -> skip (no calculation, status = "INPUT_VALUES_USED")
#   - If missing      -> solve from M records, set status = "INFILLED" or
#                        an error code if validation fails
#   - For plots with multiple M records: solves per measurement, reports
#     drift stats (sd, range, slope), uses LAST measurement for main value
#   - Latitude is sign-agnostic: stored internally as negative (NZ convention)
#
# Errors flag the plot but do not halt the batch. The output CSV gains:
#   infill_status, infill_error  -- so users can see and fix issues
#
# Run standalone:
#   source("infill_plots.R")
#   infill_plots("examples/csv_inputs/plots.csv",
#                "examples/csv_inputs/psp_summary.csv",
#                "batch_output")
#
# Or via run_batch.R with the run_infill flag.
# =============================================================================

# =============================================================================
# Standalone helpers (don't require full Inputparms() / MODEL_ENV setup)
# =============================================================================

# Populate the volume-table matrix in global env, sized correctly for the
# VBA-style indexing used by the _env functions: v[voltable, coef].
# Requires voltab() from TreeLevel_Input.R to be available.
.infill_init_v <- function(voltable = 2L) {
  # voltab() writes V[coef, voltable] into MODEL_ENV. The _env functions
  # in 300index2025V1.2.R expect v[voltable, coef], so we transpose.
  voltab()
  V_env <- as.matrix(get("V", envir = MODEL_ENV))
  assign("v",        t(V_env), envir = .GlobalEnv)
  assign("voltable", voltable, envir = .GlobalEnv)
  invisible(t(V_env))
}

# Solve I300 from a measurement (age, MTH or BA, stocking).
# Uses solve_I300_from_DBH from 300index2025V1.2.R.
# Returns NA on failure rather than throwing.
.solve_i300_standalone <- function(SI, Age, MTH, BA, N, latitude, elevation) {
  if (is.na(BA) || BA <= 0) {
    # Without BA we cannot derive DBH directly; would need volume too.
    return(NA_real_)
  }
  DBH_obs <- sqrt(1.273 * BA / N) * 100   # cm
  tryCatch(
    solve_I300_from_DBH(DBH_obs, SI, Age, N, latitude, elevation),
    error = function(e) NA_real_
  )
}

# =============================================================================
infill_plots <- function(plots_csv, psp_csv, output_dir,
                         pars = NULL,
                         verbose = TRUE) {

  # Initialise voltable matrix (default = Kimberley 2006 for radiata)
  .infill_init_v(voltable = 2L)

  plots_df <- read.csv(plots_csv, check.names = FALSE, stringsAsFactors = FALSE)
  psp_df   <- read.csv(psp_csv,   check.names = FALSE, stringsAsFactors = FALSE)

  plots_df <- plots_df[nzchar(plots_df$Plot), , drop = FALSE]

  col_lat  <- "Latitude (decimal degrees)"
  col_elev <- "Elevation above sea level (m)"
  col_i300 <- "300 Index"
  col_si   <- "Site Index"

  stopifnot(all(c(col_lat, col_elev, col_i300, col_si) %in% names(plots_df)))

  # Normalise latitude to negative (NZ convention) -- sign-agnostic input
  plots_df[[col_lat]] <- -abs(as.numeric(plots_df[[col_lat]]))

  # Drift / status columns
  plots_df$SI_sd            <- NA_real_
  plots_df$SI_range         <- NA_real_
  plots_df$I300_sd          <- NA_real_
  plots_df$I300_range       <- NA_real_
  plots_df$I300_drift_slope <- NA_real_
  plots_df$n_measurements   <- 0L
  plots_df$infill_status    <- ""
  plots_df$infill_error     <- ""

  n_infilled_si    <- 0
  n_infilled_i300  <- 0
  n_already_filled <- 0
  n_flagged        <- 0

  flag <- function(i, code, msg) {
    plots_df$infill_status[i] <<- code
    plots_df$infill_error[i]  <<- msg
    n_flagged <<- n_flagged + 1
    if (verbose) message(sprintf("  %s: [%s] %s", plots_df$Plot[i], code, msg))
  }

  for (i in seq_len(nrow(plots_df))) {
    pid  <- plots_df$Plot[i]
    lat  <- as.numeric(plots_df[[col_lat]][i])
    elev <- as.numeric(plots_df[[col_elev]][i])
    i300 <- suppressWarnings(as.numeric(plots_df[[col_i300]][i]))
    si   <- suppressWarnings(as.numeric(plots_df[[col_si]][i]))

    # ---- Validate location ----
    if (is.na(lat) || lat < -48 || lat > -30) {
      flag(i, "BAD_LATITUDE",
           sprintf("latitude %.2f outside NZ range [-48, -30]", lat))
      next
    }
    if (is.na(elev) || elev < 0 || elev > 2000) {
      flag(i, "BAD_ELEVATION",
           sprintf("elevation %.0f outside valid range [0, 2000]", elev))
      next
    }

    # ---- Validate existing SI / I300 if present ----
    if (!is.na(si) && (si < 5 || si > 60)) {
      flag(i, "BAD_SI",
           sprintf("Site Index %.2f outside valid range [5, 60]", si))
      next
    }
    if (!is.na(i300) && (i300 < 1 || i300 > 60)) {
      flag(i, "BAD_I300",
           sprintf("300 Index %.2f outside valid range [1, 60]", i300))
      next
    }

    # ---- Skip if both already present ----
    if (!is.na(si) && !is.na(i300)) {
      plots_df$infill_status[i] <- "INPUT_VALUES_USED"
      n_already_filled <- n_already_filled + 1
      next
    }

    # ---- Collect M records ----
    plot_psp <- psp_df[psp_df$Plot == pid & psp_df$Type == "M", ]
    plot_psp <- plot_psp[order(plot_psp[["Age (years)"]]), ]
    n_meas   <- nrow(plot_psp)
    plots_df$n_measurements[i] <- n_meas

    if (n_meas == 0) {
      flag(i, "NO_MEASUREMENTS",
           "SI or 300 Index missing and no M record in psp_summary")
      next
    }

    # ---- Solve SI and I300 per measurement ----
    si_vec   <- rep(NA_real_, n_meas)
    i300_vec <- rep(NA_real_, n_meas)
    skip_plot <- FALSE
    for (m in seq_len(n_meas)) {
      m_age <- as.numeric(plot_psp[["Age (years)"]][m])
      m_mth <- as.numeric(plot_psp[["MTH (m)"]][m])
      m_sph <- as.numeric(plot_psp[["Stocking (stems/ha)"]][m])
      m_ba  <- as.numeric(plot_psp[["BA (m2/ha)"]][m])

      if (is.na(m_age) || m_age < 0.1 || m_age > 100) {
        flag(i, "BAD_AGE",
             sprintf("meas %d age %.2f outside valid range [0.1, 100]", m, m_age))
        skip_plot <- TRUE; break
      }
      if (is.na(m_mth) || m_mth < 0.1 || m_mth > 100) {
        flag(i, "BAD_MTH",
             sprintf("meas %d MTH %.2f outside valid range [0.1, 100]", m, m_mth))
        skip_plot <- TRUE; break
      }
      if (is.na(m_sph) || m_sph < 10 || m_sph > 15000) {
        flag(i, "BAD_STOCKING",
             sprintf("meas %d stocking %.0f outside valid range [10, 15000]",
                     m, m_sph))
        skip_plot <- TRUE; break
      }

      # SI from MTH
      si_val <- if (isTRUE(all.equal(m_age, 20))) m_mth
                else solve_SI_from_MTH_env(m_mth, m_age, lat, elev)
      if (is.na(si_val) || si_val < 5 || si_val > 60) {
        flag(i, "BAD_SI_SOLVED",
             sprintf("meas %d solved SI=%.2f outside valid range [5, 60]",
                     m, si_val))
        skip_plot <- TRUE; break
      }
      si_vec[m] <- si_val

      # I300 from BA (needs SI first)
      i300_val <- .solve_i300_standalone(si_val, m_age, m_mth, m_ba, m_sph, lat, elev)
      if (!is.na(i300_val) && (i300_val < 1 || i300_val > 60)) {
        i300_val <- NA_real_  # don't fail, just leave NA
      }
      i300_vec[m] <- i300_val
    }
    if (skip_plot) next

    # SI drift stats
    if (sum(!is.na(si_vec)) >= 2) {
      plots_df$SI_sd[i]    <- round(sd(si_vec, na.rm = TRUE), 3)
      plots_df$SI_range[i] <- round(max(si_vec, na.rm = TRUE) -
                                    min(si_vec, na.rm = TRUE), 3)
    }

    # I300 drift stats
    valid_i300 <- !is.na(i300_vec)
    if (sum(valid_i300) >= 2) {
      i300_valid_vals <- i300_vec[valid_i300]
      ages_valid     <- as.numeric(plot_psp[["Age (years)"]])[valid_i300]
      plots_df$I300_sd[i]    <- round(sd(i300_valid_vals), 3)
      plots_df$I300_range[i] <- round(max(i300_valid_vals) - min(i300_valid_vals), 3)
      # VBA-style slope: (last - first) / (last_age - first_age)
      plots_df$I300_drift_slope[i] <- round(
        (tail(i300_valid_vals, 1) - head(i300_valid_vals, 1)) /
        (tail(ages_valid, 1) - head(ages_valid, 1)), 4)
    }

    # Infill SI from last measurement
    if (is.na(si)) {
      last_si <- tail(si_vec[!is.na(si_vec)], 1)
      plots_df[[col_si]][i] <- round(last_si, 2)
      n_infilled_si <- n_infilled_si + 1
      si <- last_si
    }

    # Infill I300 from last measurement
    if (is.na(i300)) {
      last_i300 <- tail(i300_vec[!is.na(i300_vec)], 1)
      if (length(last_i300) == 0) {
        plots_df$infill_status[i] <- "SI_INFILLED_I300_FAILED"
      } else {
        plots_df[[col_i300]][i] <- round(last_i300, 2)
        n_infilled_i300 <- n_infilled_i300 + 1
        plots_df$infill_status[i] <- "INFILLED"
        i300 <- last_i300
      }
    } else {
      plots_df$infill_status[i] <- "SI_INFILLED"
    }

    if (verbose) {
      msg <- sprintf("  %s: SI=%.2f, I300=%s (last of %d meas)",
                     pid, si,
                     if (is.na(i300)) "NA" else sprintf("%.2f", i300),
                     n_meas)
      if (n_meas > 1) {
        msg <- paste0(msg, sprintf("  SI sd=%.2f range=%.2f",
                                   plots_df$SI_sd[i], plots_df$SI_range[i]))
        if (!is.na(plots_df$I300_sd[i])) {
          msg <- paste0(msg, sprintf("; I300 sd=%.2f range=%.2f slope=%.3f",
                                     plots_df$I300_sd[i],
                                     plots_df$I300_range[i],
                                     plots_df$I300_drift_slope[i]))
        }
      }
      message(msg)
    }
  }

  if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
  out_file <- file.path(output_dir, "plots_infilled.csv")
  write.csv(plots_df, out_file, row.names = FALSE)

  if (verbose) {
    message(sprintf(
      "\nInfill complete: %d SI infilled, %d I300 infilled, %d already had values, %d flagged",
      n_infilled_si, n_infilled_i300, n_already_filled, n_flagged))
    message(sprintf("Wrote %s", out_file))
    if (n_flagged > 0) {
      message("Flagged plots remain in output with infill_status / infill_error set")
    }
  }

  invisible(plots_df)
}


# =============================================================================
# populate_i300_drift()
#
# Post-batch step. Reads per-measurement I300 values (which the batch may
# export) and writes drift stats back into plots_infilled.csv.
#
# Expected per-measurement CSV columns: Plot, Age, I300
# =============================================================================
populate_i300_drift <- function(infilled_csv, per_measurement_i300_csv) {
  plots_df <- read.csv(infilled_csv, check.names = FALSE, stringsAsFactors = FALSE)
  pm_df    <- read.csv(per_measurement_i300_csv, check.names = FALSE,
                       stringsAsFactors = FALSE)

  for (pid in unique(pm_df$Plot)) {
    rows <- pm_df[pm_df$Plot == pid, ]
    rows <- rows[order(rows$Age), ]
    if (nrow(rows) < 2) next

    i300_vals <- rows$I300
    ages      <- rows$Age

    idx <- which(plots_df$Plot == pid)
    if (length(idx) == 0) next

    plots_df$I300_sd[idx]    <- round(sd(i300_vals, na.rm = TRUE), 3)
    plots_df$I300_range[idx] <- round(max(i300_vals, na.rm = TRUE) -
                                      min(i300_vals, na.rm = TRUE), 3)
    plots_df$I300_drift_slope[idx] <- round(
      (tail(i300_vals, 1) - head(i300_vals, 1)) /
      (tail(ages, 1) - head(ages, 1)), 4)
  }

  write.csv(plots_df, infilled_csv, row.names = FALSE)
  invisible(plots_df)
}
