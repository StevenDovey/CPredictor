setwd(dirname(rstudioapi::getSourceEditorContext()$path))

if (!exists("read_data", mode = "function")) source("io_utils.R")
source("TreeLevel_Input.R")
source("300index2025V1.2.R")
source("CChange_model.R")
source("DouglasFir_500Index.R")
source("MultiSpecies_Growth.R")

# ---------------------------------------------------------------------------
# PSP batch runner — implements the FCP_5_2.xlsm batch workflow
# Reads 3 user-input sheets:
#   C_Change control  — run settings (row range, rotation, Y/N toggles)
#   PSP Summary       — per-plot records (measurements, thinnings, prunings)
#   Plots             — site info per plot
# ---------------------------------------------------------------------------
run_batch_psp <- function(input_source,
                          parameters_csv = NULL,
                          output_dir = "batch_output",
                          output_format = "csv") {

  # Load species model parameters from CSV
  if (is.null(parameters_csv)) {
    if (dir.exists(input_source)) {
      parameters_csv <- file.path(input_source, "parameters.csv")
    }
  }
  if (!is.null(parameters_csv) && file.exists(parameters_csv)) {
    message("Loading model parameters from: ", parameters_csv)
  } else {
    stop("parameters.csv not found. Provide path via parameters_csv argument or place in input directory.")
  }

  # ---- 1. Read C_Change control settings ----------------------------------
  # Read with col_names = TRUE so CSV header row is consumed; for Excel the
  # first row becomes the header automatically.  We then look up settings by
  # matching the label text in column 1, which avoids hard-coded row indices
  # and works identically for both .xlsm and .csv inputs.
  ctrl <- read_sheet(input_source, "C_Change control", col_names = FALSE)
  ctrl_lookup <- function(pattern) {
    idx <- grep(pattern, ctrl[[1]], ignore.case = TRUE)
    if (length(idx) == 0) return(NA)
    ctrl[idx[1], 2]
  }
  rotlth1    <- as.integer(ctrl_lookup("1st Rotation"))
  rotlth2    <- as.integer(ctrl_lookup("2nd Rotation"))
  run_cc     <- toupper(trimws(as.character(ctrl_lookup("Run C-Change")))) == "Y"
  detail_cc  <- toupper(trimws(as.character(ctrl_lookup("detailed")))) == "Y"
  est_drift  <- toupper(trimws(as.character(ctrl_lookup("drift")))) == "Y"
  last_meas  <- toupper(trimws(as.character(ctrl_lookup("last measurement")))) == "Y"
  run_nubalm <- toupper(trimws(as.character(ctrl_lookup("NuBalM")))) == "Y"

  message(sprintf("PSP batch: rot1=%d, rot2=%d, C-Change=%s",
                  rotlth1, rotlth2, run_cc))

  # ---- 2. Read PSP Summary ------------------------------------------------
  psp <- read_sheet(input_source, "PSP Summary", col_names = TRUE)
  # Standardise column names (handle slight variations)
  psp_names <- c("Plot", "Type", "Age", "Stocking", "BA", "MTH",
                 "Pruned_stems", "Pruned_height")
  if (ncol(psp) >= 8) names(psp)[1:8] <- psp_names
  # Convert numeric columns
  for (col in c("Age", "Stocking", "BA", "MTH", "Pruned_stems", "Pruned_height")) {
    if (col %in% names(psp)) psp[[col]] <- suppressWarnings(as.numeric(psp[[col]]))
  }

  # ---- 3. Read Plots -------------------------------------------------------
  plots_df <- read_sheet(input_source, "Plots", col_names = TRUE)
  # Standardise column names to match VBA column indices
  plots_std <- c("Plot", "Species", "Year_planted", "Latitude", "Elevation",
                 "Needle_retention", "Soil_C", "Soil_N", "Soil_Organic_P",
                 "Early_survival", "Mean_Temp",
                 "Outerwood_Density", "Outerwood_density_age",
                 "Inner_ring", "Outer_ring",
                 "I300", "Site_Index", "Drift",
                 "Mort_add", "Mort_mult")
  nc <- min(ncol(plots_df), length(plots_std))
  names(plots_df)[1:nc] <- plots_std[1:nc]
  # If there are initial litter columns (21-35), name them
  litter_names <- c(
    "Init_DM_needle", "Init_DM_branch", "Init_DM_stem",
    "Init_DM_coarse_root", "Init_DM_fine_root",
    "Init_N_needle", "Init_N_branch", "Init_N_stem",
    "Init_N_coarse_root", "Init_N_fine_root",
    "Init_P_needle", "Init_P_branch", "Init_P_stem",
    "Init_P_coarse_root", "Init_P_fine_root"
  )
  if (ncol(plots_df) >= 35) {
    names(plots_df)[21:35] <- litter_names
  }

  if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

  # ---- 4. Process plots ----------------------------------------------------
  # Process all rows in PSP Summary
  plot_ids <- unique(psp$Plot)
  message(sprintf("Found %d unique plots in PSP Summary", length(plot_ids)))

  all_yield      <- list()
  all_carbon     <- list()
  all_cc_detail  <- list()
  plots_processed <- data.frame()

  for (pid in plot_ids) {
    message(sprintf("  Processing plot: %s", pid))
    reset_model_env()
    load_parameters_from_csv(parameters_csv)

    # Look up plot-level site info from Plots sheet
    plot_row <- plots_df[plots_df$Plot == pid, , drop = FALSE]
    if (nrow(plot_row) == 0) {
      message(sprintf("    WARNING: Plot '%s' not found in Plots sheet, skipping", pid))
      next
    }
    plot_row <- plot_row[1, ]  # take first match

    species   <- as.character(plot_row$Species)
    latitude  <- as.numeric(plot_row$Latitude)
    altitude  <- as.numeric(plot_row$Elevation)
    nr_val    <- as.numeric(plot_row$Needle_retention)
    soil_c    <- as.numeric(plot_row$Soil_C)
    soil_n    <- as.numeric(plot_row$Soil_N)
    soil_p    <- as.numeric(plot_row$Soil_Organic_P)
    early_surv <- as.numeric(plot_row$Early_survival)
    temp_val  <- as.numeric(plot_row$Mean_Temp)
    core_dens <- if ("Outerwood_Density" %in% names(plot_row)) as.numeric(plot_row$Outerwood_Density) else NA
    core_age  <- if ("Outerwood_density_age" %in% names(plot_row)) as.numeric(plot_row$Outerwood_density_age) else NA
    inner_ring <- if ("Inner_ring" %in% names(plot_row)) as.numeric(plot_row$Inner_ring) else NA
    outer_ring <- if ("Outer_ring" %in% names(plot_row)) as.numeric(plot_row$Outer_ring) else NA
    i300_plot <- if ("I300" %in% names(plot_row)) as.numeric(plot_row$I300) else NA
    si_plot   <- if ("Site_Index" %in% names(plot_row)) as.numeric(plot_row$Site_Index) else NA
    drift_val <- if ("Drift" %in% names(plot_row)) as.numeric(plot_row$Drift) else NA
    mort_add  <- if ("Mort_add" %in% names(plot_row)) as.numeric(plot_row$Mort_add) else NA
    mort_mult <- if ("Mort_mult" %in% names(plot_row)) as.numeric(plot_row$Mort_mult) else NA

    # Get PSP records for this plot
    plot_psp <- psp[psp$Plot == pid, , drop = FALSE]

    # Filter to last measurement only if requested
    if (last_meas) {
      m_rows <- plot_psp[plot_psp$Type == "M", , drop = FALSE]
      non_m_rows <- plot_psp[plot_psp$Type != "M", , drop = FALSE]
      if (nrow(m_rows) > 1) {
        m_rows <- m_rows[nrow(m_rows), , drop = FALSE]  # keep only last M row
      }
      plot_psp <- rbind(non_m_rows, m_rows)
    }

    # Parse PSP records into regime components (matching VBA Module 7 logic)
    thins <- list()
    prunes <- list()
    measurement <- list(Age = NA, Stocking = NA, BA = NA, MTH = NA)
    initial_stocking <- NA

    for (r in seq_len(nrow(plot_psp))) {
      rec <- plot_psp[r, ]
      rtype <- trimws(as.character(rec$Type))

      if (rtype == "E") {
        # Establishment record — initial stocking
        initial_stocking <- rec$Stocking
      } else if (rtype == "M") {
        # Measurement record
        measurement$Age      <- rec$Age
        measurement$Stocking <- rec$Stocking
        measurement$BA       <- rec$BA
        measurement$MTH      <- rec$MTH
      } else if (rtype == "TW") {
        # Waste thinning
        thins[[length(thins) + 1]] <- list(type = "W", age = rec$Age, stocking_after = rec$Stocking)
      } else if (rtype == "TP") {
        # Production thinning
        thins[[length(thins) + 1]] <- list(type = "P", age = rec$Age, stocking_after = rec$Stocking)
      } else if (grepl("^P[1-5]$", rtype)) {
        # Pruning lift
        lift_num <- as.integer(sub("P", "", rtype))
        prunes[[length(prunes) + 1]] <- list(
          lift = lift_num, age = rec$Age,
          pruned_stems = rec$Pruned_stems, pruned_height = rec$Pruned_height
        )
      }
    }

    # Map species codes to full names for run_model()
    species_map <- c(
      "PRAD" = "Radiata pine", "PSME" = "Douglas-fir",
      "CUST" = "Cypress (lusitanica)", "CUMC" = "Cypress (macrocarpa)",
      "EUCL" = "Eucalyptus",
      "ACME" = "Blackwood", "SEQU" = "Coast redwood"
    )
    species_full <- if (species %in% names(species_map)) species_map[[species]] else species

    # ---- Build C Change parameters in MODEL_ENV ----------------------------
    assign("Species", species_full, envir = MODEL_ENV)
    assign("rotlength", rotlth1, envir = MODEL_ENV)
    assign("rotlength2", rotlth2, envir = MODEL_ENV)

    if (!is.na(initial_stocking) && initial_stocking != 0) {
      assign("N0", initial_stocking, envir = MODEL_ENV)
    }
    if (!is.na(latitude)) assign("latitude", latitude, envir = MODEL_ENV)
    if (!is.na(altitude)) assign("altitude", altitude, envir = MODEL_ENV)
    if (!is.na(soil_c))     assign("SoilC", soil_c, envir = MODEL_ENV)
    if (!is.na(soil_n))     assign("SoilN", soil_n, envir = MODEL_ENV)
    if (!is.na(soil_p))     assign("SoilOrganicP", soil_p, envir = MODEL_ENV)
    if (!is.na(early_surv)) assign("Early_survival", early_surv, envir = MODEL_ENV)
    if (!is.na(temp_val))   assign("MATEMP", temp_val, envir = MODEL_ENV)
    if (!is.na(nr_val))     assign("NR", nr_val, envir = MODEL_ENV)
    if (!is.na(core_dens))  assign("CoreDens", core_dens, envir = MODEL_ENV)
    if (!is.na(core_age))   assign("CoreAge", core_age, envir = MODEL_ENV)
    if (!is.na(inner_ring)) assign("InnerRing", inner_ring, envir = MODEL_ENV)
    if (!is.na(outer_ring)) assign("OuterRing", outer_ring, envir = MODEL_ENV)
    if (!is.na(i300_plot))  assign("I300", i300_plot, envir = MODEL_ENV)
    if (!is.na(si_plot))    assign("SI", si_plot, envir = MODEL_ENV)
    if (!is.na(drift_val))  assign("drift", drift_val, envir = MODEL_ENV)
    if (!is.na(mort_add))   assign("mortadd", mort_add, envir = MODEL_ENV)
    if (!is.na(mort_mult))  assign("mortmult", mort_mult, envir = MODEL_ENV)
    assign("RunNuBalM", run_nubalm, envir = MODEL_ENV)
    assign("estimate_drift", est_drift, envir = MODEL_ENV)
    assign("detailed_cchange", detail_cc, envir = MODEL_ENV)

    # Measurement values
    if (!is.na(measurement$Age)) {
      assign("HMTH", measurement$MTH, envir = MODEL_ENV)
      assign("HAge", measurement$Age, envir = MODEL_ENV)
      assign("Hstock", measurement$Stocking, envir = MODEL_ENV)
      assign("HBA", measurement$BA, envir = MODEL_ENV)
    }

    # Stocking history (thinnings)
    stock_hist_N <- rep(0, 10)
    stock_hist_T <- rep(0, 10)
    stock_hist_Type <- rep(0, 10)
    if (!is.na(initial_stocking)) stock_hist_N[1] <- initial_stocking
    for (ti in seq_along(thins)) {
      th <- thins[[ti]]
      stock_hist_T[ti + 1] <- th$age
      stock_hist_N[ti + 1] <- th$stocking_after
      stock_hist_Type[ti + 1] <- if (th$type == "P") 2 else 1
    }
    assign("Stock_hist_N", stock_hist_N, envir = MODEL_ENV)
    assign("Stock_hist_T", stock_hist_T, envir = MODEL_ENV)
    assign("Stock_hist_Type", stock_hist_Type, envir = MODEL_ENV)

    # Pruning history
    for (pi_idx in seq_along(prunes)) {
      pr <- prunes[[pi_idx]]
      assign(paste0("PruneAge", pr$lift), pr$age, envir = MODEL_ENV)
      if (!is.na(pr$pruned_stems)) assign(paste0("PruneN", pr$lift), pr$pruned_stems, envir = MODEL_ENV)
      if (!is.na(pr$pruned_height)) assign(paste0("PRUNEHT", pr$lift), pr$pruned_height, envir = MODEL_ENV)
    }

    # Initial litter stocks
    if (ncol(plot_row) >= 35) {
      for (li in 1:5) {
        dm <- as.numeric(plot_row[[20 + li]])
        nv <- as.numeric(plot_row[[25 + li]])
        pv <- as.numeric(plot_row[[30 + li]])
        if (!is.na(dm)) assign(paste0("Initial_DryMat_", li), dm, envir = MODEL_ENV)
        if (!is.na(nv)) assign(paste0("Initial_N_", li), nv, envir = MODEL_ENV)
        if (!is.na(pv)) assign(paste0("Initial_P_", li), pv, envir = MODEL_ENV)
      }
    }

    # ---- Build synthetic data_300_index and input_data matrices -------------
    # These matrices replicate the Excel "300 Index" and "Inputs" sheet layouts
    # that Inputparms() and Input_parameters() expect to read from MODEL_ENV.
    data_300_index <- matrix(NA, nrow = 72, ncol = 6)
    if (!is.na(si_plot))            data_300_index[4, 3]  <- si_plot
    if (!is.na(measurement$Age))    data_300_index[7, 3]  <- measurement$Age
    if (!is.na(measurement$Stocking)) data_300_index[8, 3] <- measurement$Stocking
    if (!is.na(measurement$BA))     data_300_index[9, 3]  <- measurement$BA
    if (!is.na(measurement$MTH))    data_300_index[10, 3] <- measurement$MTH
    if (!is.na(measurement$Age))    data_300_index[14, 3] <- measurement$Age
    if (!is.na(measurement$MTH))    data_300_index[15, 3] <- measurement$MTH
    if (!is.na(initial_stocking))   data_300_index[19, 3] <- initial_stocking
    data_300_index[8, 6] <- 2  # implementation mode = 2 (Offset mode for PSP)
    if (!is.na(drift_val))          data_300_index[64, 6] <- drift_val
    else                            data_300_index[64, 6] <- 0
    assign("data_300_index", as.data.frame(data_300_index), envir = MODEL_ENV)

    # Synthetic data_300_indexX (text flags — voltable selection etc.)
    data_300_indexX <- data.frame(matrix("", nrow = 22, ncol = 3), stringsAsFactors = FALSE)
    data_300_indexX[2, 1] <- "x"  # voltable = 2 (Kimberley 2006)
    assign("data_300_indexX", data_300_indexX, envir = MODEL_ENV)

    # Synthetic input_data matrix (Inputs sheet layout)
    input_data <- matrix(0, nrow = 80, ncol = 10)
    input_data[2, 4] <- match(species_full, c("Radiata pine", "Douglas-fir",
      "Cypress (lusitanica)", "Cypress (macrocarpa)", "Eucalyptus",
      "Blackwood", "Coast redwood",
      "E. regnans", "E. fastigata", "E. nitens",
      "E. delegatensis", "E. saligna"))
    if (is.na(input_data[2, 4])) input_data[2, 4] <- 0
    if (!is.na(latitude))         input_data[3, 4]  <- latitude
    if (!is.na(altitude))         input_data[4, 4]  <- altitude
    if (!is.na(nr_val))           input_data[5, 4]  <- nr_val
    if (!is.na(soil_c))           input_data[6, 4]  <- soil_c
    if (!is.na(soil_n))           input_data[7, 4]  <- soil_n
    if (!is.na(soil_p))           input_data[8, 4]  <- soil_p
    if (!is.na(early_surv))       input_data[9, 4]  <- early_surv
    if (!is.na(temp_val))         input_data[10, 4] <- temp_val
    # Stocking history rows (matching VBA layout)
    for (si in seq_along(stock_hist_N)) {
      input_data[19 + si, 2] <- stock_hist_T[si]
      input_data[19 + si, 3] <- stock_hist_N[si]
      input_data[19 + si, 4] <- stock_hist_Type[si]
    }
    # Pruning history rows
    for (pi_idx in seq_along(prunes)) {
      pr <- prunes[[pi_idx]]
      input_data[29 + pi_idx, 2] <- pr$age
      if (!is.na(pr$pruned_stems))  input_data[29 + pi_idx, 3] <- pr$pruned_stems
      if (!is.na(pr$pruned_height)) input_data[29 + pi_idx, 4] <- pr$pruned_height
    }
    assign("input_data", as.data.frame(input_data), envir = MODEL_ENV)

    # ---- Run model for this plot -------------------------------------------
    tryCatch({
      run_model()

      # --- Plots Processed row (matches FCP_5_2 "Plots Processed" sheet) ------
      # Columns: Plot, Age, 300 Index, Site Index, Total C, AGL C, BGL C,
      #          DWL C, FL C, Species
      env_get <- function(var, default = NA) {
        if (exists(var, envir = MODEL_ENV)) get(var, envir = MODEL_ENV) else default
      }
      model_I300  <- env_get("I300")
      model_SI    <- env_get("SI")
      total_c     <- env_get("TotalC", 0)
      agl_c       <- env_get("AGL_C", 0)
      bgl_c       <- env_get("BGL_C", 0)
      dwl_c       <- env_get("DWL_C", 0)
      fl_c        <- env_get("FL_C", 0)

      proc_row <- data.frame(
        Plot           = pid,
        Age            = measurement$Age,
        `300 Index`    = ifelse(is.na(model_I300), 0, model_I300),
        `Site Index`   = ifelse(is.na(model_SI), 0, model_SI),
        `Total C`      = total_c,
        `AGL C`        = agl_c,
        `BGL C`        = bgl_c,
        `DWL C`        = dwl_c,
        `FL C`         = fl_c,
        Species        = species,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
      plots_processed <- rbind(plots_processed, proc_row)

      # --- Yield Tables (matches FCP_5_2 "Yield Tables" sheet) ----------------
      # Columns: Plot, Index age, Age, Stocking b4 thin, Stocking aft thin,
      #          MTH, Crown Lth, Volume b4 thin, Volume aft thin,
      #          BA b4 thin, BA aft thin, DBH b4 thin, DBH aft thin, Mean Height
      if (exists("yield_table", envir = MODEL_ENV)) {
        yt_raw <- get("yield_table", envir = MODEL_ENV)
        # Build output with FCP_5_2 column names — pull from whatever
        # columns the model produced and map to the standard names
        n_yt <- nrow(yt_raw)
        col_or_zero <- function(df, col) {
          if (col %in% names(df)) df[[col]] else rep(0, nrow(df))
        }
        yt_out <- data.frame(
          Plot                = rep(pid, n_yt),
          `Index age`         = col_or_zero(yt_raw, "Index_age"),
          Age                 = col_or_zero(yt_raw, "Age"),
          `Stocking b4 thin`  = col_or_zero(yt_raw, "Stocking_b4_thin"),
          `Stocking aft thin` = col_or_zero(yt_raw, "Stocking_aft_thin"),
          MTH                 = col_or_zero(yt_raw, "MTH"),
          `Crown Lth`         = col_or_zero(yt_raw, "Crown_length"),
          `Volume b4 thin`    = col_or_zero(yt_raw, "Volume_b4_thin"),
          `Volume aft thin`   = col_or_zero(yt_raw, "Volume_aft_thin"),
          `BA b4 thin`        = col_or_zero(yt_raw, "BA_b4_thin"),
          `BA aft thin`       = col_or_zero(yt_raw, "BA_aft_thin"),
          `DBH b4 thin`       = col_or_zero(yt_raw, "DBH_b4_thin"),
          `DBH aft thin`      = col_or_zero(yt_raw, "DBH_aft_thin"),
          `Mean Height`       = col_or_zero(yt_raw, "Mean_Height"),
          stringsAsFactors = FALSE,
          check.names = FALSE
        )
        all_yield[[length(all_yield) + 1]] <- yt_out
      }

      # --- C_Change Predictions (matches FCP_5_2 "C_Change Predictions" sheet) -
      # Columns: Plot, Index age, Age (years), Stocking b4 thin,
      #   Stocking aft thin, Height (m), Volume net (m3/ha),
      #   Volume aft thin (m3/ha), Volume dead (m3/ha),
      #   Density sheath (kg/m3), <blank>,
      #   Rot1 Total (tC/ha), Rot1 AGL, Rot1 BGL, Rot1 DWL, Rot1 FL,
      #   Rot2 Total (tC/ha), Rot2 AGL, Rot2 BGL, Rot2 DWL, Rot2 FL,
      #   <blank>, Shrub Rot1 (tC/ha), Shrub Rot2 (tC/ha)
      if (run_cc && exists("carbon_results", envir = MODEL_ENV)) {
        cr_raw <- get("carbon_results", envir = MODEL_ENV)
        n_cr <- nrow(cr_raw)
        cc_col <- function(col) {
          if (col %in% names(cr_raw)) cr_raw[[col]] else rep(0, n_cr)
        }
        cc_out <- data.frame(
          Plot                        = rep(pid, n_cr),
          `Index age`                 = cc_col("Index_age"),
          `Age (years)`               = cc_col("Age"),
          `Stocking b4 thin (sph)`    = cc_col("Stocking_b4_thin"),
          `Stocking aft thin`         = cc_col("Stocking_aft_thin"),
          `Height (m)`                = cc_col("Height"),
          `Volume net (m3/ha)`        = cc_col("Volume_net"),
          `Volume aft thin (m3/ha)`   = cc_col("Volume_aft_thin"),
          `Volume dead (m3/ha)`       = cc_col("Volume_dead"),
          `Density sheath (kg/m3)`    = cc_col("Density_sheath"),
          blank1                      = rep(NA, n_cr),
          `Rot1 Total (tC/ha)`        = cc_col("Rot1_Total"),
          `Rot1 AGL (tC/ha)`          = cc_col("Rot1_AGL"),
          `Rot1 BGL (tC/ha)`          = cc_col("Rot1_BGL"),
          `Rot1 DWL (tC/ha)`          = cc_col("Rot1_DWL"),
          `Rot1 FL (tC/ha)`           = cc_col("Rot1_FL"),
          `Rot2 Total (tC/ha)`        = cc_col("Rot2_Total"),
          `Rot2 AGL (tC/ha)`          = cc_col("Rot2_AGL"),
          `Rot2 BGL (tC/ha)`          = cc_col("Rot2_BGL"),
          `Rot2 DWL (tC/ha)`          = cc_col("Rot2_DWL"),
          `Rot2 FL (tC/ha)`           = cc_col("Rot2_FL"),
          blank2                      = rep(NA, n_cr),
          `Shrub Rot1 (tC/ha)`        = cc_col("Shrub_Rot1"),
          `Shrub Rot2 (tC/ha)`        = cc_col("Shrub_Rot2"),
          stringsAsFactors = FALSE,
          check.names = FALSE
        )
        all_carbon[[length(all_carbon) + 1]] <- cc_out
      }

      # --- C_Change Output (detailed, matches LP1OUT structure) ----------------
      if (run_cc && detail_cc && exists("cchange_detail", envir = MODEL_ENV)) {
        cd_raw <- get("cchange_detail", envir = MODEL_ENV)
        cd_raw$Plot <- pid
        all_cc_detail[[length(all_cc_detail) + 1]] <- cd_raw
      }

      message(sprintf("    OK: I300=%.2f, SI=%.2f",
                       ifelse(is.na(model_I300), 0, model_I300),
                       ifelse(is.na(model_SI), 0, model_SI)))
    }, error = function(e) {
      message(sprintf("    ERROR: %s", conditionMessage(e)))
    })
  }

  # ---- 5. Write outputs ----------------------------------------------------
  # Output files mirror the 4 FCP_5_2.xlsm output sheets:
  #   Plots Processed, Yield Tables, C_Change Predictions, C_Change Output
  ext <- ".csv"
  output_files <- list()

  if (nrow(plots_processed) > 0) {
    out_file <- file.path(output_dir, paste0("plots_processed", ext))
    write_output(plots_processed, out_file)
    output_files$plots_processed <- out_file
  }

  if (length(all_yield) > 0) {
    yield_df <- do.call(rbind, all_yield)
    out_file <- file.path(output_dir, paste0("yield_tables", ext))
    write_output(yield_df, out_file)
    output_files$yield_tables <- out_file
  }

  if (length(all_carbon) > 0) {
    carbon_df <- do.call(rbind, all_carbon)
    out_file <- file.path(output_dir, paste0("c_change_predictions", ext))
    write_output(carbon_df, out_file)
    output_files$c_change_predictions <- out_file
  }

  if (length(all_cc_detail) > 0) {
    detail_df <- do.call(rbind, all_cc_detail)
    out_file <- file.path(output_dir, paste0("c_change_output", ext))
    write_output(detail_df, out_file)
    output_files$c_change_output <- out_file
  }

  message(sprintf("\nPSP batch complete: %d plots processed, output in %s/",
                  nrow(plots_processed), output_dir))
  message(sprintf("Output files: %s", paste(basename(unlist(output_files)), collapse = ", ")))
  invisible(list(
    plots_processed     = plots_processed,
    yield_tables        = if (length(all_yield) > 0) do.call(rbind, all_yield) else NULL,
    c_change_predictions = if (length(all_carbon) > 0) do.call(rbind, all_carbon) else NULL,
    c_change_output     = if (length(all_cc_detail) > 0) do.call(rbind, all_cc_detail) else NULL,
    output_files        = output_files
  ))
}


