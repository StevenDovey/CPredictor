setwd(dirname(rstudioapi::getSourceEditorContext()$path))

if (!exists("read_data", mode = "function")) source("io_utils.R")
source("TreeLevel_Input.R")
source("300index2025V1.2.R")
source("CChange_model.R")
source("DouglasFir_500Index.R")
source("MultiSpecies_Growth.R")

# Capture initial environment state so reset_model_env() preserves constants
snapshot_initial_env()

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

    # ---- Validate required data -----------------------------------------------
    if (is.na(initial_stocking) || initial_stocking < 1) stop(sprintf("Plot '%s': missing or invalid initial stocking — no E record in PSP Summary", pid))
    if (is.na(measurement$Age))     stop(sprintf("Plot '%s': missing measurement Age in PSP Summary (no M record)", pid))
    if (is.na(measurement$Stocking)) stop(sprintf("Plot '%s': missing measurement Stocking in PSP Summary", pid))

    # ---- Calculate SI from measurement if not provided ----------------------
    # Mirrors VBA: if SI is blank, solve from Age/MTH via height model bisection
    if (is.na(si_plot) || si_plot == 0) {
      if (is.na(measurement$MTH) || measurement$MTH <= 0) {
        stop(sprintf("Plot '%s': Site_Index is blank and MTH is missing — cannot calculate SI", pid))
      }
      if (is.na(latitude) || is.na(altitude)) {
        stop(sprintf("Plot '%s': Site_Index is blank and Latitude/Elevation missing — cannot calculate SI", pid))
      }
      si_plot <- solve_SI_from_MTH_env(measurement$MTH, measurement$Age, latitude, altitude)
      if (is.na(si_plot) || si_plot <= 0) {
        stop(sprintf("Plot '%s': SI calculation failed (MTH=%.2f, Age=%.2f)", pid, measurement$MTH, measurement$Age))
      }
      message(sprintf("    Calculated SI=%.2f from Age=%.2f, MTH=%.2f", si_plot, measurement$Age, measurement$MTH))
      assign("SI", si_plot, envir = MODEL_ENV)
    }

    # ---- Build synthetic data_300_index matrix --------------------------------
    # Replicates Excel "300 Index" sheet for Inputparms().
    # Cross-reference -- Inputparms() reads:
    #   [4,3]=SI  [7,3]=Age  [8,3]=Stocking  [9,3]=BA  [10,3]=MTH
    #   [14,3]=HAge  [15,3]=HMTH  [19,3]=initialstocking
    #   [20:23, 2:6]=Stocking_history  [40:44, 2:5]=Pruning_history
    #   [47,3]=maxage  [48,3]=steplength
    #   [8,6]=implementation  [64,6]=drift
    # Input_parameters() writes back for Radiata:
    #   [3,3]=I300  [3,6]=latitude  [4,6]=elevation
    #   [75,4]=Soil_C  [76,4]=Soil_N  [77,4]=MAT
    data_300_index <- matrix(NA, nrow = 80, ncol = 6)

    # Site indices
    if (!is.na(i300_plot))          data_300_index[3, 3]  <- i300_plot
    if (!is.na(si_plot))            data_300_index[4, 3]  <- si_plot

    # Measurement data
    data_300_index[7, 3]  <- measurement$Age
    data_300_index[8, 3]  <- measurement$Stocking
    data_300_index[9, 3]  <- 0  # DBH not directly measured in PSP; model derives from BA
    if (!is.na(measurement$BA))     data_300_index[10, 3] <- measurement$BA
    data_300_index[14, 3] <- measurement$Age       # HAge
    if (!is.na(measurement$MTH))    data_300_index[15, 3] <- measurement$MTH  # HMTH

    # Initial stocking
    if (!is.na(initial_stocking))   data_300_index[19, 3] <- initial_stocking

    # Stocking history -- rows 20:23, cols 2:6
    # Inputparms() reads: col2=age, col3=N_before, col4=N_after, col5=thincoeff, col6=thinratio
    for (ti in seq_along(thins)) {
      if (ti > 4) break
      th <- thins[[ti]]
      data_300_index[19 + ti, 2] <- th$age
      data_300_index[19 + ti, 4] <- th$stocking_after
      data_300_index[19 + ti, 5] <- 0  # thincoeff 0 = use default
    }

    # Pruning history -- rows 40:44, cols 2:5
    # Inputparms() reads: col2=age, col3=height, col4=sph, col5=prunecoeff
    for (pi_idx in seq_along(prunes)) {
      if (pi_idx > 5) break
      pr <- prunes[[pi_idx]]
      data_300_index[39 + pi_idx, 2] <- pr$age
      if (!is.na(pr$pruned_height)) data_300_index[39 + pi_idx, 3] <- pr$pruned_height
      if (!is.na(pr$pruned_stems))  data_300_index[39 + pi_idx, 4] <- pr$pruned_stems
    }

    # Growth model control
    data_300_index[47, 3] <- rotlth1       # maxage = rotation length
    data_300_index[48, 3] <- 1.0           # steplength
    data_300_index[8, 6]  <- 2             # implementation = Offset mode for PSP
    data_300_index[64, 6] <- if (!is.na(drift_val)) drift_val else 0

    # Spatial/environmental for 300 Index
    if (!is.na(latitude))   data_300_index[3, 6]  <- latitude
    if (!is.na(altitude))   data_300_index[4, 6]  <- altitude
    if (!is.na(soil_c))     data_300_index[75, 4] <- soil_c
    if (!is.na(soil_n))     data_300_index[76, 4] <- soil_n
    if (!is.na(temp_val))   data_300_index[77, 4] <- temp_val

    assign("data_300_index", as.data.frame(data_300_index), envir = MODEL_ENV)

    # Synthetic data_300_indexX (text flags -- voltable, bias, mortality model)
    # Needs 66 rows x 6 cols for check_input_htfn [64:65,4] and voltable [1:11,1]
    data_300_indexX <- data.frame(matrix("", nrow = 66, ncol = 6), stringsAsFactors = FALSE)
    data_300_indexX[2, 1] <- "x"  # voltable = 2 (Kimberley 2006)
    assign("data_300_indexX", data_300_indexX, envir = MODEL_ENV)

    # ---- Build synthetic input_data matrix ------------------------------------
    # Replicates Excel "Inputs" sheet for Input_parameters().
    # Cross-reference -- Input_parameters() reads:
    #   [2,4]=species  [3,4]=I300  [4,4]=H30  [5,4]=initial_stocking  [6,4]=rotlength
    #   [9:13, 4:7]=Thinning_schedule  [9:11, 11:14]=Pruning_schedule
    #   [16,3]=mode  [16,4]=DiaDist
    #   [20,4]=T1  [21,4]=N1  [22,4]=H1  [22,5]=H1_type  [23,4]=D1  [23,5]=D1_type
    #   [24,4]=T2  [25,4]=H2
    #   [26:34,4]=model params (-999=use default)
    #   [35,4]=latitude  [36,4]=elevation  [37,4]=Soil_C  [38,4]=Soil_N
    #   [39,4]=MAT  [40,4]=drift
    input_data <- matrix(-999, nrow = 80, ncol = 14)

    # Species index
    input_data[2, 4] <- match(species_full, c("Radiata pine", "Douglas-fir",
      "Cypress (lusitanica)", "Cypress (macrocarpa)", "Eucalyptus",
      "Blackwood", "Coast redwood",
      "E. regnans", "E. fastigata", "E. nitens",
      "E. delegatensis", "E. saligna"))
    if (is.na(input_data[2, 4])) stop(sprintf("Plot '%s': species '%s' not recognised", pid, species_full))

    # Site indices (0 = calibrate from stand metrics)
    input_data[3, 4] <- if (!is.na(i300_plot)) i300_plot else 0   # I300
    input_data[4, 4] <- if (!is.na(si_plot)) si_plot else 0       # H30 / Site Index

    # Initial stocking and rotation length
    input_data[5, 4] <- if (!is.na(initial_stocking)) initial_stocking else 0
    input_data[6, 4] <- rotlth1

    # Thinning schedule -- rows 9:13, cols 4:7 (thin 1-4)
    # Row 9=type, 10=thin_age, 11=Stock_hist_T, 12=Stock_hist_N, 13=thincoeff
    for (ti in seq_along(thins)) {
      if (ti > 4) break
      th <- thins[[ti]]
      col <- 3 + ti  # cols 4,5,6,7
      input_data[9, col]  <- if (th$type == "P") 2 else 1  # 1=waste, 2=production
      input_data[10, col] <- th$age                         # thin_age
      input_data[11, col] <- th$age                         # Stock_hist_T
      input_data[12, col] <- th$stocking_after              # Stock_hist_N
      input_data[13, col] <- -999                           # thincoeff: -999 = use default
    }

    # Pruning schedule -- rows 9:11, cols 11:14 (lift 1-4)
    # Row 9=prune_age, 10=prune_N, 11=prune_height
    for (pi_idx in seq_along(prunes)) {
      if (pi_idx > 4) break
      pr <- prunes[[pi_idx]]
      col <- 10 + pi_idx  # cols 11,12,13,14
      input_data[9, col]  <- pr$age
      input_data[10, col] <- if (!is.na(pr$pruned_stems)) pr$pruned_stems else 0
      input_data[11, col] <- if (!is.na(pr$pruned_height)) pr$pruned_height else 0
    }

    # Calibration mode
    input_data[16, 3] <- 2    # mode = 2: Calibrate from stand metrics
    input_data[16, 4] <- 1    # DiaDist = 1: Weibull

    # Calibration data from measurement record
    input_data[20, 4] <- measurement$Age                        # T1
    input_data[21, 4] <- measurement$Stocking                   # N1
    input_data[22, 4] <- if (!is.na(measurement$MTH)) measurement$MTH else -999  # H1
    input_data[22, 5] <- 1    # H1 is already MTH (no conversion needed)
    input_data[23, 4] <- if (!is.na(measurement$BA)) measurement$BA else -999    # D1 (as BA)
    input_data[23, 5] <- 1    # D1 flag: 1 = BA (convert to qDBH in Input_parameters)
    input_data[24, 4] <- 0    # T2 (no secondary calibration)
    input_data[25, 4] <- 0    # H2

    # Model parameters -- rows 26:34 already -999 from matrix init (= use defaults)

    # Environmental / spatial data
    input_data[35, 4] <- if (!is.na(latitude)) latitude else -999
    input_data[36, 4] <- if (!is.na(altitude)) altitude else -999
    input_data[37, 4] <- if (!is.na(soil_c)) soil_c else -999
    input_data[38, 4] <- if (!is.na(soil_n)) soil_n else -999
    input_data[39, 4] <- if (!is.na(temp_val)) temp_val else -999
    input_data[40, 4] <- if (!is.na(drift_val)) drift_val else -999

    assign("input_data", as.data.frame(input_data), envir = MODEL_ENV)

    # Control flags for run_model()
    assign("Check_errors", FALSE, envir = MODEL_ENV)  # Error_checks_1/2/3/5 not yet implemented
    assign("Minimal_run", FALSE, envir = MODEL_ENV)

    # ---- Run model for this plot -------------------------------------------
    tryCatch({
      # ---- Growth-only pathway (skip tree-level) --------------------------------
      # Instead of run_model() which needs tree-level variables (nstems, Treelist,
      # Cali, etc.), call the growth model directly.
      Inputparms()
      Input_parameters()

      # Calibration (mode 2: calibrate from stand metrics)
      siteIndex()
      Calc300Index()
      if (Species == "Radiata pine") {
        Calibrate_radiata()
      } else if (Species == "Douglas-fir") {
        Calibrate_dfir()
      } else {
        Calibrate()
      }

      # Run growth simulation
      growth_result <- OutputGrowth()
      gdf <- growth_result$growth_df

      # Build yield table directly from growth_df
      if (!is.null(gdf) && nrow(gdf) > 0) {
        yt <- data.frame(
          Plot                = rep(pid, nrow(gdf)),
          Index_age           = seq_len(nrow(gdf)),
          Age                 = gdf$Age,
          Stocking_b4_thin    = gdf$N,
          Stocking_aft_thin   = gdf$N,
          MTH                 = gdf$MTH,
          Crown_length        = rep(0, nrow(gdf)),
          Volume_b4_thin      = gdf$Vol,
          Volume_aft_thin     = gdf$Vol,
          BA_b4_thin          = gdf$BA,
          BA_aft_thin         = gdf$BA,
          DBH_b4_thin         = gdf$DBH,
          DBH_aft_thin        = gdf$DBH,
          Mean_Height         = if ("mnheight" %in% names(gdf)) gdf$mnheight else gdf$MTH,
          stringsAsFactors = FALSE
        )
        assign("yield_table", yt, envir = MODEL_ENV)
      }

      # ---- Carbon pathway (C_Change) -----------------------------------------
      # Build growth_table for run_cchange() directly from growth_df
      total_c <- 0; agl_c <- 0; bgl_c <- 0; dwl_c <- 0; fl_c <- 0
      if (run_cc && !is.null(gdf) && nrow(gdf) > 0) {
        growth_table <- data.frame(
          Age  = gdf$Age,
          SPHA = gdf$N,
          MTH  = gdf$MTH,
          BA   = gdf$BA,
          Vol  = gdf$Vol,
          GrossVol = gdf$Vol,
          WholeStemDens = if ("WoodDensity" %in% names(gdf)) gdf$WoodDensity else rep(0.42, nrow(gdf)),
          RingDens = if ("WoodDensity" %in% names(gdf)) gdf$WoodDensity else rep(0.42, nrow(gdf)),
          stringsAsFactors = FALSE
        )

        # Build disturbance schedule from thinning data
        dist_sched <- NULL
        if (length(thins) > 0) {
          dist_rows <- lapply(thins, function(th) {
            data.frame(Age = th$age, SPHA = th$stocking_after, BA = -1,
                       PruneHt = -1, StemExtract = 0, CrownExtract = 0,
                       FloorExtract = 0, stringsAsFactors = FALSE)
          })
          dist_sched <- do.call(rbind, dist_rows)
        }

        # Density inputs
        cc_soil_c <- if (!is.na(soil_c)) soil_c else 5.57
        cc_soil_n <- if (!is.na(soil_n)) soil_n else 0.296
        cc_mat    <- if (!is.na(temp_val)) temp_val else 12

        tryCatch({
          cc_result <- run_cchange(
            growth_table       = growth_table,
            disturbance_schedule = dist_sched,
            IROT               = 1,
            soil_c             = cc_soil_c,
            soil_n             = cc_soil_n,
            MATEMP             = cc_mat
          )
          if (!is.null(cc_result) && !is.null(cc_result$annual_carbon)) {
            ac <- cc_result$annual_carbon
            last <- ac[nrow(ac), ]
            total_c <- if ("CSTAND" %in% names(last)) last$CSTAND else 0
            agl_c   <- if ("CTREES" %in% names(last)) last$CTREES else 0
            bgl_c   <- if ("CROOTL" %in% names(last)) last$CROOTL else 0
            dwl_c   <- if ("C_stem_litter" %in% names(last)) last$C_stem_litter else 0
            fl_c    <- if ("C_needle_litter" %in% names(last)) last$C_needle_litter else 0

            # Store carbon_results for C_Change Predictions output
            assign("carbon_results", ac, envir = MODEL_ENV)
            assign("cchange_detail", ac, envir = MODEL_ENV)
          }
        }, error = function(e) {
          message(sprintf("    C_Change warning: %s", conditionMessage(e)))
        })
      }
      # --- Plots Processed row (matches FCP_5_2 "Plots Processed" sheet) ------
      # Columns: Plot, Age, 300 Index, Site Index, Total C, AGL C, BGL C,
      #          DWL C, FL C, Species
      env_get <- function(var, default = NA) {
        if (exists(var, envir = MODEL_ENV)) get(var, envir = MODEL_ENV) else default
      }
      model_I300  <- env_get("I300")
      model_SI    <- env_get("SI")

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
          `Age (years)`               = cc_col("Age"),
          `Stocking (sph)`            = cc_col("SPHA"),
          `Height (m)`                = cc_col("HT"),
          `Volume (m3/ha)`            = cc_col("Vol"),
          `Density (t/m3)`            = cc_col("dens"),
          `Stand C (tC/ha)`           = cc_col("CSTAND"),
          `Tree C (tC/ha)`            = cc_col("CTREES"),
          `Shrub C (tC/ha)`           = cc_col("CSHRUB"),
          `Foliage C (tC/ha)`         = cc_col("CFAS"),
          `Stem C (tC/ha)`            = cc_col("CSTEM"),
          `Live Root C (tC/ha)`       = cc_col("CROOTL"),
          `Dead Root C (tC/ha)`       = cc_col("CROOTD"),
          `Branch Live C (tC/ha)`     = cc_col("C_branch_live"),
          `Branch Dead C (tC/ha)`     = cc_col("C_branch_dead"),
          `Needle Litter C (tC/ha)`   = cc_col("C_needle_litter"),
          `Branch Litter C (tC/ha)`   = cc_col("C_branch_litter"),
          `Stem Litter C (tC/ha)`     = cc_col("C_stem_litter"),
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

