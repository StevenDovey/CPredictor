# ==========================================================================
# Input Validation - R port of VBA Modules 3 and 8
#
# Validation routines for 300 Index, multi-species, and tree-list inputs.
# Returns a list of errors/warnings; stops execution if critical errors found.
# ==========================================================================

# ---------------------------------------------------------------------------
# Validate 300 Index inputs (Module 3)
# ---------------------------------------------------------------------------
validate_300index_inputs <- function(data_300_index, data_300_indexX = NULL,
                                    implementation = 1) {
  errors <- character(0)
  warnings <- character(0)

  # Site Index check
  si_val <- data_300_index[4, 3]
  if (is.na(si_val) || si_val < 5 || si_val > 60) {
    errors <- c(errors, "Site Index missing or outside range [5, 60]")
  }

  # 300 Index / I300 check
  i300_age <- data_300_index[7, 3]
  i300_stk <- data_300_index[8, 3]
  if (is.na(i300_age) || is.na(i300_stk) ||
      i300_age < 0.1 || i300_age > 100 ||
      i300_stk < 10 || i300_stk > 15000) {
    errors <- c(errors, "300 Index measurement age/stocking missing or outside range")
  }
  if (all(is.na(data_300_index[9:11, 3]))) {
    errors <- c(errors, "300 Index: no height/DBH/BA measurement provided")
  }

  # Height/age measurement
  if (any(is.na(data_300_index[14:15, 3])) ||
      any(data_300_index[14:15, 3] < 0.1) ||
      any(data_300_index[14:15, 3] > 100)) {
    errors <- c(errors, "Height/age measurement missing or outside range [0.1, 100]")
  }

  # Initial stocking
  init_stk <- data_300_index[19, 3]
  if (is.na(init_stk) || init_stk < 1 || init_stk > 80000) {
    errors <- c(errors, "Initial stocking missing or outside range [1, 80000]")
  }

  # Fell age / rotation length
  fell_age <- data_300_index[47, 3]
  if (is.na(fell_age) || fell_age < 1 || fell_age > 100) {
    errors <- c(errors, "Fell age (rotation length) missing or outside range [1, 100]")
  }

  # Step length
  step_len <- data_300_index[48, 3]
  if (!is.na(step_len) && (step_len < 0.01 || step_len > 2)) {
    errors <- c(errors, "Step length outside range [0.01, 2]")
  }

  # Height function check
  if (!is.null(data_300_indexX)) {
    ht_mods <- sum(tolower(as.character(data_300_indexX[64:65, 4])) == "x", na.rm = TRUE)
    if (ht_mods != 1) {
      errors <- c(errors, "Exactly one height function must be selected")
    }

    # Volume function check
    vol_mods <- sum(tolower(as.character(data_300_indexX[51:61, 4])) == "x", na.rm = TRUE)
    if (vol_mods != 1) {
      errors <- c(errors, "Exactly one volume table must be selected")
    }

    # Mortality function check
    mort_mods <- sum(tolower(as.character(data_300_indexX[68:72, 4])) == "x", na.rm = TRUE)
    if (mort_mods != 1) {
      warnings <- c(warnings, "Exactly one mortality function should be selected; defaulting to model 6")
    }
  }

  # Pruning history check
  for (i in 40:44) {
    if (!is.na(data_300_index[i, 2])) {
      if (i > 40 && is.na(data_300_index[i - 1, 2])) {
        errors <- c(errors, sprintf("Pruning row %d: gap in pruning ages", i))
      }
      if (is.na(data_300_index[i, 3]) ||
          data_300_index[i, 3] < 0 || data_300_index[i, 3] > 20) {
        errors <- c(errors, sprintf("Pruning row %d: height missing or outside range [0, 20]", i))
      }
    }
  }

  # Stocking history check
  old_age <- 0
  old_sph <- if (!is.na(data_300_index[19, 3])) data_300_index[19, 3] else 0
  for (i in 20:36) {
    age_val <- data_300_index[i, 2]
    sph1 <- data_300_index[i, 3]
    sph2 <- data_300_index[i, 4]
    if (!is.na(age_val) && age_val != 0) {
      if (age_val > 100) {
        errors <- c(errors, sprintf("Stocking row %d: age > 100", i))
      }
      if (age_val <= old_age && old_age > 0) {
        errors <- c(errors, sprintf("Stocking row %d: age not increasing", i))
      }
      if (!is.na(sph1) && sph1 != 0 && sph1 > old_sph) {
        errors <- c(errors, sprintf("Stocking row %d: post-thin stocking exceeds previous", i))
      }
      old_age <- age_val
      old_sph <- if (!is.na(sph1) && sph1 != 0) sph1 else old_sph
      if (!is.na(sph2) && sph2 > 0) old_sph <- sph2
    }
  }

  list(errors = errors, warnings = warnings, valid = length(errors) == 0)
}

# ---------------------------------------------------------------------------
# Validate multi-species inputs (Module 8)
# ---------------------------------------------------------------------------
validate_multispecies_inputs <- function(input_data) {
  errors <- character(0)
  warnings <- character(0)

  # Rotation length
  rot <- input_data[6, 5]
  if (is.na(rot) || rot < 1 || rot > 100) {
    errors <- c(errors, "Rotation length outside allowed range [1, 100]")
  }

  # Stocking at planting
  stk <- input_data[5, 5]
  if (is.na(stk) || stk < 1 || stk > 10000) {
    errors <- c(errors, "Stocking at planting missing or outside range [1, 10000]")
  }

  # 300 Index
  i300 <- input_data[3, 5]
  if (!is.na(i300) && (i300 < 1 || i300 > 70)) {
    errors <- c(errors, "300 Index outside allowed range [1, 70]")
  }

  # Site Index
  si <- input_data[4, 5]
  if (!is.na(si) && (si < 1 || si > 60)) {
    errors <- c(errors, "Site Index outside allowed range [1, 60]")
  }

  # Thinning schedule validation
  for (thin_idx in 1:4) {
    col <- thin_idx + 4
    thin_age <- input_data[11, col]
    if (!is.na(thin_age) && thin_age != 0) {
      if (thin_age < 1 || thin_age > 100) {
        errors <- c(errors, sprintf("Thinning %d: age outside range [1, 100]", thin_idx))
      }
      thin_sph <- input_data[12, col]
      if (is.na(thin_sph) || thin_sph < 1 || thin_sph > 10000) {
        errors <- c(errors, sprintf("Thinning %d: stocking after thin outside range [1, 10000]", thin_idx))
      }
      thin_coeff <- input_data[13, col]
      if (!is.na(thin_coeff) && thin_coeff != -999 && (thin_coeff < 0 || thin_coeff > 10)) {
        errors <- c(errors, sprintf("Thinning %d: coefficient outside range [0, 10]", thin_idx))
      }
    }
  }

  # Pruning schedule validation
  for (lift_idx in 1:4) {
    col <- lift_idx + 12
    prune_age <- input_data[9, col]
    if (!is.na(prune_age) && prune_age != 0) {
      if (prune_age < 1 || prune_age > 100) {
        errors <- c(errors, sprintf("Pruning %d: age outside range [1, 100]", lift_idx))
      }
      prune_sph <- input_data[10, col]
      if (is.na(prune_sph) || prune_sph < 1 || prune_sph > 10000) {
        errors <- c(errors, sprintf("Pruning %d: stems pruned outside range [1, 10000]", lift_idx))
      }
      prune_ht <- input_data[11, col]
      if (is.na(prune_ht) || prune_ht < 0 || prune_ht > 15) {
        errors <- c(errors, sprintf("Pruning %d: height outside range [0, 15]", lift_idx))
      }
    }
  }

  # Calibration measurement validation
  meas_age <- input_data[20, 5]
  if (!is.na(meas_age) && meas_age != 0) {
    if (meas_age < 1 || meas_age > 200) {
      errors <- c(errors, "Measurement age outside range [1, 200]")
    }
    meas_stk <- input_data[21, 5]
    if (is.na(meas_stk) || meas_stk < 1 || meas_stk > 10000) {
      errors <- c(errors, "Measurement stocking outside range [1, 10000]")
    }
    meas_ht <- input_data[22, 5]
    if (is.na(meas_ht) || meas_ht < 1 || meas_ht > 200) {
      errors <- c(errors, "Measurement height outside range [1, 200]")
    }
    meas_dbh <- input_data[23, 5]
    if (is.na(meas_dbh) || meas_dbh < 1 || meas_dbh > 500) {
      errors <- c(errors, "Measurement DBH/BA outside range [1, 500]")
    }
  }

  list(errors = errors, warnings = warnings, valid = length(errors) == 0)
}

# ---------------------------------------------------------------------------
# Validate tree list inputs (Module 8, Error_checks_4)
# ---------------------------------------------------------------------------
validate_tree_list <- function(tree_list, plot_area, age) {
  errors <- character(0)
  warnings <- character(0)

  if (is.na(plot_area) || plot_area < 0.001 || plot_area > 100) {
    errors <- c(errors, "Plot area missing or outside range [0.001, 100]")
  }

  if (is.na(age) || age < 1 || age > 200) {
    errors <- c(errors, "Age of tree list missing or outside range [1, 200]")
  }

  nstems <- nrow(tree_list)
  if (nstems < 2 || nstems > 1000) {
    errors <- c(errors, sprintf("Number of stems (%d) outside range [2, 1000]", nstems))
  }

  no_dbh <- sum(!is.na(tree_list[, 2]))
  no_ht <- sum(!is.na(tree_list[, 3]))

  if (no_dbh != nstems) {
    errors <- c(errors, sprintf("Missing DBH: %d of %d stems have DBH", no_dbh, nstems))
  }

  if (no_ht < 3) {
    errors <- c(errors, sprintf("Only %d height measurements; need at least 3", no_ht))
  }

  list(errors = errors, warnings = warnings, valid = length(errors) == 0)
}

# ---------------------------------------------------------------------------
# Validate yield table (Module 8, Error_checks_5)
# ---------------------------------------------------------------------------
validate_yield_table <- function(N_vector, rotlength) {
  warnings <- character(0)
  for (t_val in 2:min(length(N_vector), rotlength + 1)) {
    if (!is.na(N_vector[t_val]) && !is.na(N_vector[t_val - 1]) &&
        N_vector[t_val] > N_vector[t_val - 1]) {
      warnings <- c(warnings, sprintf(
        "Yield table stocking increases at age %d (from %.0f to %.0f) - check inputs",
        t_val - 1, N_vector[t_val - 1], N_vector[t_val]))
    }
  }
  list(errors = character(0), warnings = warnings, valid = TRUE)
}

# ---------------------------------------------------------------------------
# Run all validations and report
# ---------------------------------------------------------------------------
run_all_validations <- function(data_300_index = NULL, data_300_indexX = NULL,
                                input_data = NULL, tree_list = NULL,
                                plot_area = NULL, age = NULL,
                                implementation = 1, stop_on_error = TRUE) {
  all_errors <- character(0)
  all_warnings <- character(0)

  if (!is.null(data_300_index)) {
    v <- validate_300index_inputs(data_300_index, data_300_indexX, implementation)
    all_errors <- c(all_errors, v$errors)
    all_warnings <- c(all_warnings, v$warnings)
  }

  if (!is.null(input_data)) {
    v <- validate_multispecies_inputs(input_data)
    all_errors <- c(all_errors, v$errors)
    all_warnings <- c(all_warnings, v$warnings)
  }

  if (!is.null(tree_list)) {
    v <- validate_tree_list(tree_list, plot_area, age)
    all_errors <- c(all_errors, v$errors)
    all_warnings <- c(all_warnings, v$warnings)
  }

  if (length(all_warnings) > 0) {
    for (w in all_warnings) warning(w)
  }

  if (length(all_errors) > 0) {
    msg <- paste("Input validation failed:\n", paste("-", all_errors, collapse = "\n"))
    if (stop_on_error) stop(msg)
    message(msg)
  }

  list(
    errors = all_errors,
    warnings = all_warnings,
    valid = length(all_errors) == 0
  )
}
