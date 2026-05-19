# ---------------------------------------------------------------------------
# Model environment: model functions reference free variables via lexical
# scoping, which resolves to .GlobalEnv.  MODEL_ENV therefore points to
# .GlobalEnv so that assign(..., envir = MODEL_ENV) puts values where the
# functions can see them.  reset_model_env() clears non-function objects
# between batch iterations so state does not leak across plots.
# ---------------------------------------------------------------------------
MODEL_ENV <- .GlobalEnv

# VBA module-level arrays — shared across Growth(), Height(), Ageshifts(),
# Diameter(), Newlift(), thinning().  Initialised at source time so every
# function can find them; Growth() resets them with <<- each call.
Meanht     <- numeric(11)
adjageel   <- numeric(11)
initiallag <- numeric(9)
ThinLag    <- numeric(9)
agethin    <- numeric(9)

# VBA module-level string/numeric defaults (VBA Dim initialises strings to ""
# and doubles to 0).  Radiata pine never sets MTH_model etc., so they must
# exist as empty strings for TP_fn / AgeBH to fall through and return 0.
MTH_model  <- ""
MTH_form   <- ""
MTH_a      <- 0
MTH_b      <- 0
MTH_c      <- 0
DBH_model  <- ""
DBH_form   <- ""
DBH_a      <- 0; DBH_b <- 0; DBH_c <- 0; DBH_d <- 0
DBH_f      <- 0; DBH_g <- 0; DBH_h <- 0; DBH_k <- 0
Check_errors <- FALSE
Minimal_run  <- FALSE
Error_flag   <- FALSE

# VBA output subroutine stubs — not yet ported; Growth() calls these when
# OUTPUT=TRUE.  Stubs are no-ops so the batch pathway doesn't crash.
OutStep     <- function() invisible(NULL)
OutThin     <- function() invisible(NULL)
OutPrune    <- function() invisible(NULL)
OutElements <- function() invisible(NULL)
earlyield   <- function() invisible(NULL)
mortvol     <- function() invisible(NULL)

# ---------------------------------------------------------------------------
# density(growth_df) — Port of VBA Module 1 Sub density().
# Computes growth-sheath wood density (g/cm3) for each row in growth_df.
# Returns the growth_df with a WoodDensity column appended.
# Requires: sheathdens(), outdens26(), outdens(), old_outdens(),
#           Calcagezero()  (all defined in 300index2025V1.2.R or this file).
# ---------------------------------------------------------------------------
density <- function(growth_df) {
  if (!is.data.frame(growth_df) || nrow(growth_df) == 0) return(growth_df)

  # Read density inputs from data_300_index (VBA Cells references)
  loc_SoilC      <- suppressWarnings(as.numeric(data_300_index[75, 4]))
  loc_SoilN      <- suppressWarnings(as.numeric(data_300_index[76, 4]))
  loc_Temp       <- suppressWarnings(as.numeric(data_300_index[77, 4]))
  loc_CoreDens   <- suppressWarnings(as.numeric(data_300_index[78, 4]))
  loc_CoreAge    <- suppressWarnings(as.numeric(data_300_index[79, 4]))
  loc_InnerRing  <- suppressWarnings(as.numeric(data_300_index[80, 4]))
  loc_OuterRing  <- suppressWarnings(as.numeric(data_300_index[81, 4]))
  loc_GeneticAdj <- suppressWarnings(as.numeric(data_300_index[82, 4]))
  loc_densitymodel <- suppressWarnings(as.numeric(data_300_index[83, 4]))

  nz <- function(x) is.finite(x) && x != 0

  densityinfo <-
    if (nz(loc_SoilC) && nz(loc_SoilN) && nz(loc_Temp)) 1L
    else if (nz(loc_CoreDens) && nz(loc_CoreAge)) 2L
    else if (nz(loc_CoreDens) && nz(loc_InnerRing) && nz(loc_OuterRing)) 3L
    else if (nz(loc_CoreDens)) 4L
    else 5L

  agezero <- Calcagezero()

  outdensring <- 0
  Wcal <- 3

  if (densityinfo == 1L) {
    loc_CoreAge <- 26
    loc_CoreDens <- outdens26(loc_SoilC, loc_SoilN, loc_Temp,
                              stocking = 250, GeneticAdj = if (is.finite(loc_GeneticAdj)) loc_GeneticAdj else 0)
    outdensring <- 18.95 - 0.024 * SI
    Wcal <- 10.19 + 0.0893 * I300 - 0.255 * SI + 0.00373 * SI^2 - 0.00339 * I300 * SI
  } else if (densityinfo == 5L) {
    loc_CoreDens <- 470
    outdensring <- 23
  }

  n <- nrow(growth_df)
  ages <- growth_df$Age
  dbhs <- growth_df$DBH
  vols <- growth_df$Vol

  # First ring width (mm)
  first_ringwidth <- 1.5
  for (i in 2:n) {
    if (is.finite(dbhs[i - 1]) && dbhs[i - 1] > 0) {
      da <- ages[i] - ages[i - 1]
      if (da > 0) first_ringwidth <- 10 * (dbhs[i] - dbhs[i - 1]) / da / 2
      break
    }
  }

  # Per-row ring width
  ringwidth <- numeric(n)
  for (i in seq_len(n)) {
    if (i == 1 || !is.finite(dbhs[i - 1]) || dbhs[i - 1] == 0) {
      ringwidth[i] <- first_ringwidth
    } else {
      da <- ages[i] - ages[i - 1]
      if (da > 0) {
        ringwidth[i] <- 10 * (dbhs[i] - dbhs[i - 1]) / da / 2
      } else {
        ringwidth[i] <- first_ringwidth
      }
    }
  }

  # Sheath density (g/cm3) for each row
  wd <- numeric(n)
  prevvol <- 0
  prev_stem <- 0
  for (i in seq_len(n)) {
    age_i <- ages[i]
    vol_i <- vols[i]

    if (densityinfo == 4L) {
      sheath <- loc_CoreDens / 1000
    } else {
      if (isTRUE(as.integer(loc_densitymodel) == 2L)) {
        ring <- max(1, age_i - agezero)
        od <- outdens(ring, ringwidth[i], loc_CoreDens, outdensring, Wcal)
      } else {
        od <- old_outdens(age_i, loc_CoreDens, outdensring)
      }
      sheath <- sheathdens(od, age_i) / 1000
    }

    # Whole-stem density (volume-weighted average)
    if (i == 1 || !is.finite(prevvol) || prevvol <= 0 || !is.finite(vol_i) || vol_i <= 0) {
      stem <- sheath
    } else {
      stem <- (sheath * (vol_i - prevvol) + prev_stem * prevvol) / vol_i
    }
    wd[i] <- sheath
    prev_stem <- stem
    prevvol <- vol_i
  }

  growth_df$WoodDensity <- wd
  growth_df
}

# Snapshot of variable names present after sourcing all model files.
# Populated by snapshot_initial_env(); used by reset_model_env() so that
# constants/defaults defined at source time are never wiped between plots.
.initial_env_names <- character(0)

snapshot_initial_env <- function() {
  .initial_env_names <<- ls(envir = .GlobalEnv)
}

reset_model_env <- function() {
  keep <- .initial_env_names
  all_names <- ls(envir = .GlobalEnv)
  to_remove <- setdiff(all_names, keep)
  # Never remove functions or environments
  to_remove <- Filter(function(nm) {
    obj <- get(nm, envir = .GlobalEnv)
    !is.function(obj) && !is.environment(obj)
  }, to_remove)
  if (length(to_remove) > 0) rm(list = to_remove, envir = .GlobalEnv)
}

if (!exists("read_data", mode = "function")) source("io_utils.R")

# ---------------------------------------------------------------------------
# Load species parameters from CSV (extracted from VBA Module 2 constants).
# Called once at startup by run_batch_psp() — never at source time.
# ---------------------------------------------------------------------------
load_parameters_from_csv <- function(csv_path) {
  params <- read.csv(csv_path, stringsAsFactors = FALSE)
  for (i in seq_len(nrow(params))) {
    vname <- trimws(params$Variable[i])
    if (is.na(vname) || vname == "") next
    val <- params$Value[i]
    if (params$Type[i] == "String") {
      assign(vname, val, envir = MODEL_ENV)
    } else {
      assign(vname, as.numeric(val), envir = MODEL_ENV)
    }
  }
}

check_input_I300 <- function(data_300_index, implementation) {
  # Check conditions based on the equivalent cell references in R
  if (is.na(data_300_index[7, 3]) || is.na(data_300_index[8, 3]) || 
      data_300_index[7, 3] < 0.1 || data_300_index[7, 3] > 100 || 
      data_300_index[8, 3] < 10 || data_300_index[8, 3] > 15000 || 
      (is.na(data_300_index[9, 3]) && is.na(data_300_index[10, 3]) && is.na(data_300_index[11, 3]))) {
    
    # Display error message if conditions are not met and implementation matches
    if (implementation == 1 || implementation == 5) {
      print("Input Error: 300 Index measurement")  # Display modal message (adjust according to your R environment)
    }
    return(FALSE)  # Return false if any condition fails
  } else {
    return(TRUE)  # Return true if all conditions are met
  }
}
check_input_SI <- function(data_300_index, implementation) {
  if (is.na(data_300_index[4, 3]) || data_300_index[4, 3] < 5 || data_300_index[4, 3] > 60) {
    if (implementation %in% c(1, 5)) print("Input Error: no Site Index")
    return(FALSE)
  }
  TRUE
}
check_input_htage <- function(data_300_index, implementation) {
  if (any(is.na(data_300_index[14:15, 3])) || 
      any(data_300_index[14:15, 3] < 0.1) || 
      any(data_300_index[14:15, 3] > 100)) {
    if (implementation %in% c(1, 5)) print("Input Error: height/age measurement")
    return(FALSE)
  }
  TRUE
}
check_input_htfn <- function(data_300_index, implementation) {
  mods <- sum(tolower(data_300_indexX[64:65, 4]) == "x")
  if (mods != 1) {
    if (implementation %in% c(1, 5)) print("Input Error: Height function")
    return(FALSE)
  }
  TRUE
}
Error_checks_4 <- function(Starting_tree_list, Plot_area, Age) {
  Error_flag <- FALSE
  
  # Check the plot area
  if (Plot_area < 0.001 || Plot_area > 100) { 
    Error_flag <- TRUE
    message("Plot area of tree list not specified or outside allowed range.")
  }
  
  # Check the age of trees
  if (Age < 1 || Age > 200) {  
    Error_flag <- TRUE
    message("Age of trees in tree list not specified or outside allowed range.")
  }
  
  # Count non-NA values for DBH and height measurements
  No_dbh <- sum(!is.na(Starting_tree_list[, 2]))  # Assuming DBH is in the second column
  No_ht <- sum(!is.na(Starting_tree_list[, 3]))   # Assuming height is in the third column
  
  # Check for valid number of stems
  nstems <- nrow(Starting_tree_list)
  if (nstems < 2 || nstems > 1000) {  
    Error_flag <- TRUE
    message("Number of stems in tree list outside allowed range.")
  }
  
  # Check for DBH and height measurements
  if (No_dbh != nstems) {  
    Error_flag <- TRUE
    message("At least one stem in tree list has a missing DBH.")
  }
  
  if (No_ht < 3) {  
    Error_flag <- TRUE
    message("At least 3 stems in tree list must have a measured height.")
  }
  
  return(Error_flag)
}

voltab <- function(workbook = NULL) {
  # VBA Module 1 hardcoded volume-table coefficients.
  # VBA indexes V(voltable, coef); R stores V[coef, voltable].
  V <- data.frame(matrix(0, nrow = 8, ncol = 11))
  V[1,1]<-0.942;   V[2,1]<- -1.161;   V[3,1]<-0.317                                            # Kimberley & Beets 2007
  V[1,2]<-0.989;   V[2,2]<- -1.2752;  V[3,2]<-0.3191                                           # Kimberley 2006
  V[1,3]<-1.492912924; V[2,3]<- -0.999113309; V[3,3]<-1.250753941; V[4,3]<- -0.397037159;
  V[5,3]<-0.027218164; V[6,3]<- -0.063166205; V[7,3]<-0.064609459; V[8,3]<- -0.030665365       # Vol fn 182
  V[1,4]<-1.633105986; V[2,4]<- -1.039327204; V[3,4]<-1.212696953; V[4,4]<- -0.359131176;
  V[5,4]<-0.026454943; V[6,4]<- -0.067457458; V[7,4]<-0.066992488; V[8,4]<- -0.030528278       # Vol fn 236
  V[1,5]<-0.730448717; V[2,5]<- -0.617440226; V[3,5]<-1.095616037; V[4,5]<- -0.222220223;
  V[5,5]<-0.013858949; V[6,5]<- -0.11022445;  V[7,5]<-0.059157535; V[8,5]<- -0.016942593       # Vol fn 328
  V[1,6]<-1.09857999;  V[2,6]<- -0.883862258; V[3,6]<-1.165375013; V[4,6]<- -0.28047221;
  V[5,6]<-0.022081234; V[6,6]<- -0.059261776; V[7,6]<-0.053187392; V[8,6]<- -0.025226521       # Vol fn 358
  V[1,7]<-1.403009551; V[2,7]<- -0.96392392;  V[3,7]<-1.221046594; V[4,7]<- -0.358337009;
  V[5,7]<-0.024975712; V[6,7]<- -0.061374804; V[7,7]<-0.061895757; V[8,7]<- -0.028672533       # Vol fn 11
  V[1,8]<-2.834246614; V[2,8]<- -1.856804825; V[3,8]<-1.152097786; V[4,8]<- -0.201346156;
  V[5,8]<- -0.000721117; V[6,8]<-0.081503044; V[7,8]<-0.024428222; V[8,8]<-0.001938887         # Vol fn 430
  V[1,9]<-2.7023; V[2,9]<- -2.1301; V[3,9]<-1.3901; V[4,9]<- -0.5056;
  V[5,9]<-0.0548; V[6,9]<-0.0991;   V[7,9]<-0.1478; V[8,9]<- -0.088                           # 3-point-taper
  V[1,10]<-6.2733; V[2,10]<-0.1284; V[3,10]<- -0.00097                                         # NSW1
  V[1,11]<-2.1819; V[2,11]<-0.2504; V[3,11]<- -0.00081                                         # NSW2
  assign("V", V, envir = MODEL_ENV)
}
Inputparms <- function() {
  

  # Input details of stand    # READ from input data
   implementation <- (data_300_index[8, 6])  # Operating mode: 1=Standard mode, 2=Offset mode, 3=Index mode
  assign("implementation", implementation, envir = MODEL_ENV)  
  SI <- ((data_300_index[4, 3]))    # Set global variable SI
  if (is.na(SI)) stop("Inputparms: SI (data_300_index[4,3]) is NA — check Plots CSV Site_Index column")
  assign("SI", SI, envir = MODEL_ENV)  
    initialstocking <- (data_300_index[19, 3])  # Set global variable initialstocking
  if (is.na(initialstocking) || initialstocking < 1) stop("Inputparms: initialstocking (data_300_index[19,3]) is NA or < 1 — check PSP Summary E record")
  assign("initialstocking", initialstocking, envir = MODEL_ENV)  
    drift <- (data_300_index[64, 6])  # Set global variable drift
  if (is.na(drift)) drift <- 0
  assign("drift", drift, envir = MODEL_ENV)  
  
  # Determine bias corrections to be used in BA model
  bias_old <- !is.na(data_300_indexX[1, 3]) && tolower(data_300_indexX[1, 3]) == "x"
  bias_young <- !is.na(data_300_indexX[2, 3]) && tolower(data_300_indexX[2, 3]) == "x"
  bias_SI <- !is.na(data_300_indexX[3, 3]) && tolower(data_300_indexX[3, 3]) == "x"
  assign("bias_old", bias_old, envir = MODEL_ENV)  
  assign("bias_young", bias_young, envir = MODEL_ENV)  
  assign("bias_SI", bias_SI, envir = MODEL_ENV)  
  
  
  
  maxage <- (data_300_index[47, 3])  # Set global variable maxAge
  if (is.na(maxage) || maxage < 1) stop("Inputparms: maxage (data_300_index[47,3]) is NA or < 1 — check rotation length in c_change_control.csv")
  assign("maxage", maxage, envir = MODEL_ENV)  
  
  steplength <- (data_300_index[48, 3])  # Set global variable stepLength
  if (is.na(steplength) || steplength < 0.01) steplength <- 1.0
  assign("steplength", steplength, envir = MODEL_ENV)  
  
  # Determine height model
  heightmodel <- heightmod()  # Call heightmod function to set height model
  assign("heightmodel", heightmodel, envir = MODEL_ENV)  
  height_coeffs <- calcheightcoeff(SI, heightmodel)
  assign("height_coeffs", height_coeffs, envir = MODEL_ENV)  
  ha <- height_coeffs$ha
  hb <- height_coeffs$hb
  assign("ha", ha, envir = MODEL_ENV)  
  assign("hb", hb, envir = MODEL_ENV)  
  
  # Determine mortality model
  mortmodel <- 6  # Default mortality model
  if (!is.na(data_300_indexX[18, 1]) && tolower(data_300_indexX[18, 1]) == "x") mortmodel <- 1
  if (!is.na(data_300_indexX[19, 1]) && tolower(data_300_indexX[19, 1]) == "x") mortmodel <- 2
  if (!is.na(data_300_indexX[20, 1]) && tolower(data_300_indexX[20, 1]) == "x") mortmodel <- 3
  if (!is.na(data_300_indexX[21, 1]) && tolower(data_300_indexX[21, 1]) == "x") mortmodel <- 5
  assign("mortmodel", mortmodel, envir = MODEL_ENV)  
  
  attrition <- 0
  pctmortadj <- 0
  # Debug print removed for batch SI/300 runs.
  
  
  ####Need to fix this
    # Adjust attrition and mortality percentage if mortmodel >= 4
  #if (mortmodel >= 4) {
  #  attrition <- ifelse(!is.na(data_300_index[68, 6]) && data_300_index[68, 6] != 0, 
  #                      data_300_index[68, 6] / 100, 
  #                      if (mortmodel==4) mortu)
  #  pctmortadj <- ifelse(!is.na(data_300_index[69, 6]) && data_300_index[69, 6] != 0, 
  #                       data_300_index[69, 6], 0)
  #}

  # Assign values to the global environment
  assign("attrition", attrition, envir = MODEL_ENV)  
  assign("pctmortadj", pctmortadj, envir = MODEL_ENV)  

  
  # Obtain volume table number from volume table array
  voltabarray <- sapply(1:11, function(i) tolower(data_300_indexX[i, 1]))
  voltable <- which(voltabarray == "x")
  assign("voltable", voltable, envir = MODEL_ENV)  
  steps <- maxage / steplength
  assign("steps", steps, envir = MODEL_ENV)  
  
  height_coeffs <- calcheightcoeff(SI, heightmodel)
  ha <- height_coeffs$ha; assign("ha", ha, envir = MODEL_ENV)  
  hb <- height_coeffs$hb; assign("hb", hb, envir = MODEL_ENV) 

  # Create the Stocking_history dataframe by copying the specified portion of data_300_index
  Stocking_history <- as.data.frame(data_300_index[20:23, 2:6])
  Stocking_history[is.na(Stocking_history)] <- 0    # Replace all NA values in the Stocking_history dataframe with 0
  colnames(Stocking_history) <- c("shist_T", "shist_N1", "shist_N2", "shist_thincoeff", "shist_thinratio")
  Stocking_history$Mortality <- 0  # Initialize the Mortality column with 0
  Stocking_history[] <- lapply(Stocking_history, as.numeric)  # Convert all columns in Stocking_history to numeric
  Nshist <- sum(!is.na(data_300_index[20:36, 2]))
  assign("Stocking_history", Stocking_history, envir = MODEL_ENV)  
  assign("Nshist", Nshist, envir = MODEL_ENV)  
  
  Stocking_history<- mort()
  assign("Stocking_history", Stocking_history, envir = MODEL_ENV)  
  
  # Read pruning history
  Pruning_history <- as.data.frame(data_300_index[40:44, 2:5])
  Pruning_history[is.na(Pruning_history)] <- 0    # Replace all NA values in the Pruning_history dataframe with 0
  colnames(Pruning_history) <- c("lift_T", "lift_height", "lift_sph", "lift_prunecoeff")
  Pruning_history[] <- lapply(Pruning_history, as.numeric)  # Convert all columns in Pruning_history to numeric
  Pruning_history$lift_sph[Pruning_history$lift_sph == 0] <- 10000
  Nlifts <- sum(!is.na(data_300_index[40:44, 2]))
  assign("Pruning_history", Pruning_history, envir = MODEL_ENV)  
  assign("Nlifts", Nlifts, envir = MODEL_ENV)  
}
Input_parameters <- function() {
  Species = Species
  I300 <- input_data[3, 4]
  H30 <- input_data[4, 4]
  T1 <- input_data[20, 4]  # Age at calibration
  if (is.na(T1) || T1 < 0.1) stop("Input_parameters: T1 (calibration age, input_data[20,4]) is NA or invalid — check PSP Summary M record Age")
  H1 <- input_data[22, 4]  # Height at calibration
  N1 <- input_data[21, 4]  # Stocking at calibration
  if (is.na(N1) || N1 < 1) stop("Input_parameters: N1 (calibration stocking, input_data[21,4]) is NA or invalid — check PSP Summary M record Stocking")
  
  if (input_data[22, 5] == 2) {
    H1 <- MTH_from_MnHt(H1, N1, MTH_MnHt_a, MTH_MnHt_b)}  # Convert calibration height from mean height to MTH if necessary
  
  D1 <- input_data[23, 4]  # Calibration BA
  if (input_data[23, 5] == 1) {
    D1 <- 200 * sqrt(D1 / N1 / pi)}  # Calculate calibration qDBH from BA if necessary
  T2 <- input_data[24, 4]
  H2 <- input_data[25, 4]
  
  {
    if (Species == "Coast redwood") {
      MTH_model <- MTH_model_red
      MTH_form <- MTH_form_red
      MTH_a <- MTH_a_red
      MTH_b <- MTH_b_red
      MTH_c <- MTH_c_red
      DBH_model <- DBH_model_red
      DBH_form <- DBH_form_red
      DBH_a <- DBH_a_red
      DBH_b <- DBH_b_red
      DBH_c <- DBH_c_red
      DBH_d <- DBH_d_red
      DBH_f <- DBH_f_red
      DBH_g <- DBH_g_red
      DBH_h <- DBH_h_red
      DBH_k <- DBH_k_red
      MTH_MnHt_a <- MTH_MnHt_a_red
      MTH_MnHt_b <- MTH_MnHt_b_red
      VOL_type <- VOL_type_red2
      VOL_u <- VOL_u_red2
      VOL_v <- VOL_v_red2
      VOL_w <- VOL_w_red2
      VOL_z <- VOL_z_red2
      THINCOEF <- THINCOEF_red
      MORT_k <- MORT_k_red
      MORT_m <- MORT_m_red
      MORT_n <- MORT_n_red
      Den_a <- DEN_a_red
      Den_b <- DEN_b_red
      
    } else if (Species == "Cupressus macrocarpa (N.I.)") {
      MTH_model <- MTH_model_mac
      MTH_form <- MTH_form_mac
      MTH_a <- MTH_a_mac
      MTH_b <- MTH_b_mac
      MTH_c <- MTH_c_mac
      DBH_model <- DBH_model_mac
      DBH_form <- DBH_form_mac
      DBH_a <- DBH_a_mac
      DBH_b <- DBH_b_mac
      DBH_c <- DBH_c_mac
      DBH_d <- DBH_d_mac
      DBH_f <- DBH_f_mac
      DBH_g <- DBH_g_mac
      DBH_h <- DBH_h_mac
      DBH_k <- DBH_k_mac
      MTH_MnHt_a <- MTH_MnHt_a_mac
      MTH_MnHt_b <- MTH_MnHt_b_mac
      VOL_type <- VOL_type_mac
      VOL_u <- VOL_u_mac
      VOL_v <- VOL_v_mac
      VOL_w <- VOL_w_mac
      VOL_z <- VOL_z_mac
      THINCOEF <- THINCOEF_cyp
      MORT_k <- MORT_k_mac_NI
      MORT_m <- MORT_m_mac_NI
      MORT_n <- MORT_n_mac_NI
      Den_a <- DEN_a_cyp
      Den_b <- DEN_b_cyp
      
    } else if (Species == "Cupressus macrocarpa (S.I.)") {
      MTH_model <- MTH_model_mac
      MTH_form <- MTH_form_mac
      MTH_a <- MTH_a_mac
      MTH_b <- MTH_b_mac
      MTH_c <- MTH_c_mac
      DBH_model <- DBH_model_mac
      DBH_form <- DBH_form_mac
      DBH_a <- DBH_a_mac
      DBH_b <- DBH_b_mac
      DBH_c <- DBH_c_mac
      DBH_d <- DBH_d_mac
      DBH_f <- DBH_f_mac
      DBH_g <- DBH_g_mac
      DBH_h <- DBH_h_mac
      DBH_k <- DBH_k_mac
      MTH_MnHt_a <- MTH_MnHt_a_mac
      MTH_MnHt_b <- MTH_MnHt_b_mac
      VOL_type <- VOL_type_mac
      VOL_u <- VOL_u_mac
      VOL_v <- VOL_v_mac
      VOL_w <- VOL_w_mac
      VOL_z <- VOL_z_mac
      THINCOEF <- THINCOEF_cyp
      MORT_k <- MORT_k_mac_SI
      MORT_m <- MORT_m_mac_SI
      MORT_n <- MORT_n_mac_SI
      Den_a <- DEN_a_cyp
      Den_b <- DEN_b_cyp
      
    } else if (Species == "Cupressus lusitanica (N.I.)") {
      MTH_model <- MTH_model_lus
      MTH_form <- MTH_form_lus
      MTH_a <- MTH_a_lus
      MTH_b <- MTH_b_lus
      MTH_c <- MTH_c_lus
      DBH_model <- DBH_model_lus
      DBH_form <- DBH_form_lus
      DBH_a <- DBH_a_lus
      DBH_b <- DBH_b_lus
      DBH_c <- DBH_c_lus
      DBH_d <- DBH_d_lus
      DBH_f <- DBH_f_lus
      DBH_g <- DBH_g_lus
      DBH_h <- DBH_h_lus
      DBH_k <- DBH_k_lus
      MTH_MnHt_a <- MTH_MnHt_a_lus
      MTH_MnHt_b <- MTH_MnHt_b_lus
      VOL_type <- VOL_type_lus
      VOL_u <- VOL_u_lus
      VOL_v <- VOL_v_lus
      VOL_w <- VOL_w_lus
      VOL_z <- VOL_z_lus
      THINCOEF <- THINCOEF_cyp
      MORT_k <- MORT_k_lus_NI
      MORT_m <- MORT_m_lus_NI
      MORT_n <- MORT_n_lus_NI
      Den_a <- DEN_a_cyp
      Den_b <- DEN_b_cyp
      
    } else if (Species == "Cupressus lusitanica (S.I.)") {
      MTH_model <- MTH_model_lus
      MTH_form <- MTH_form_lus
      MTH_a <- MTH_a_lus
      MTH_b <- MTH_b_lus
      MTH_c <- MTH_c_lus
      DBH_model <- DBH_model_lus
      DBH_form <- DBH_form_lus
      DBH_a <- DBH_a_lus
      DBH_b <- DBH_b_lus
      DBH_c <- DBH_c_lus
      DBH_d <- DBH_d_lus
      DBH_f <- DBH_f_lus
      DBH_g <- DBH_g_lus
      DBH_h <- DBH_h_lus
      DBH_k <- DBH_k_lus
      MTH_MnHt_a <- MTH_MnHt_a_lus
      MTH_MnHt_b <- MTH_MnHt_b_lus
      VOL_type <- VOL_type_lus
      VOL_u <- VOL_u_lus
      VOL_v <- VOL_v_lus
      VOL_w <- VOL_w_lus
      VOL_z <- VOL_z_lus
      THINCOEF <- THINCOEF_cyp
      MORT_k <- MORT_k_lus_SI
      MORT_m <- MORT_m_lus_SI
      MORT_n <- MORT_n_lus_SI
      Den_a <- DEN_a_cyp
      Den_b <- DEN_b_cyp
      
    } else if (Species == "Blackwood") {
      MTH_model <- MTH_model_bla
      MTH_form <- MTH_form_bla
      MTH_a <- MTH_a_bla
      MTH_b <- MTH_b_bla
      MTH_c <- MTH_c_bla
      DBH_model <- DBH_model_bla
      DBH_form <- DBH_form_bla
      DBH_a <- DBH_a_bla
      DBH_b <- DBH_b_bla
      DBH_c <- DBH_c_bla
      DBH_d <- DBH_d_bla
      DBH_f <- DBH_f_bla
      DBH_g <- DBH_g_bla
      DBH_h <- DBH_h_bla
      DBH_k <- DBH_k_bla
      MTH_MnHt_a <- MTH_MnHt_a_bla
      MTH_MnHt_b <- MTH_MnHt_b_bla
      VOL_type <- VOL_type_bla
      VOL_u <- VOL_u_bla
      VOL_v <- VOL_v_bla
      VOL_w <- VOL_w_bla
      VOL_z <- VOL_z_bla
      THINCOEF <- THINCOEF_bla
      MORT_k <- MORT_k_bla
      MORT_m <- MORT_m_bla
      MORT_n <- MORT_n_bla
      Den_a <- DEN_a_bla
      Den_b <- DEN_b_bla
      
    } else if (Species == "Eucalyptus regnans") {
      MTH_model <- MTH_model_reg
      MTH_form <- MTH_form_reg
      MTH_a <- MTH_a_reg
      MTH_b <- MTH_b_reg
      MTH_c <- MTH_c_reg
      DBH_model <- DBH_model_reg
      DBH_form <- DBH_form_reg
      DBH_a <- DBH_a_reg
      DBH_b <- DBH_b_reg
      DBH_c <- DBH_c_reg
      DBH_d <- DBH_d_reg
      DBH_f <- DBH_f_reg
      DBH_g <- DBH_g_reg
      DBH_h <- DBH_h_reg
      DBH_k <- DBH_k_reg
      MTH_MnHt_a <- MTH_MnHt_a_reg
      MTH_MnHt_b <- MTH_MnHt_b_reg
      VOL_type <- VOL_type_reg
      VOL_u <- VOL_u_reg
      VOL_v <- VOL_v_reg
      VOL_w <- VOL_w_reg
      VOL_z <- VOL_z_reg
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_reg
      MORT_m <- MORT_m_reg
      MORT_n <- MORT_n_reg
      Den_a <- DEN_a_reg
      Den_b <- DEN_b_reg
      
    } else if (Species == "Eucalyptus fastigata") {
      MTH_model <- MTH_model_fas
      MTH_form <- MTH_form_fas
      MTH_a <- MTH_a_fas
      MTH_b <- MTH_b_fas
      MTH_c <- MTH_c_fas
      DBH_model <- DBH_model_fas
      DBH_form <- DBH_form_fas
      DBH_a <- DBH_a_fas
      DBH_b <- DBH_b_fas
      DBH_c <- DBH_c_fas
      DBH_d <- DBH_d_fas
      DBH_f <- DBH_f_fas
      DBH_g <- DBH_g_fas
      DBH_h <- DBH_h_fas
      DBH_k <- DBH_k_fas
      MTH_MnHt_a <- MTH_MnHt_a_fas
      MTH_MnHt_b <- MTH_MnHt_b_fas
      VOL_type <- VOL_type_fas
      VOL_u <- VOL_u_fas
      VOL_v <- VOL_v_fas
      VOL_w <- VOL_w_fas
      VOL_z <- VOL_z_fas
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_fas
      MORT_m <- MORT_m_fas
      MORT_n <- MORT_n_fas
      Den_a <- DEN_a_fas
      Den_b <- DEN_b_fas
      
    } else if (Species == "Eucalyptus nitens (N.I.)") {
      MTH_model <- MTH_model_nit
      MTH_form <- MTH_form_nit
      MTH_a <- MTH_a_nit
      MTH_b <- MTH_b_nit
      MTH_c <- MTH_c_nit
      DBH_model <- DBH_model_nit
      DBH_form <- DBH_form_nit
      DBH_a <- DBH_a_nit
      DBH_b <- DBH_b_nit
      DBH_c <- DBH_c_nit
      DBH_d <- DBH_d_nit
      DBH_f <- DBH_f_nit
      DBH_g <- DBH_g_nit
      DBH_h <- DBH_h_nit
      DBH_k <- DBH_k_nit
      MTH_MnHt_a <- MTH_MnHt_a_nit
      MTH_MnHt_b <- MTH_MnHt_b_nit
      VOL_type <- VOL_type_nit
      VOL_u <- VOL_u_nit
      VOL_v <- VOL_v_nit
      VOL_w <- VOL_w_nit
      VOL_z <- VOL_z_nit
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_nit_NI
      MORT_m <- MORT_m_nit_NI
      MORT_n <- MORT_n_nit_NI
      Den_a <- DEN_a_nit
      Den_b <- DEN_b_nit
      
    } else if (Species == "Eucalyptus nitens (S.I.)") {
      MTH_model <- MTH_model_nit
      MTH_form <- MTH_form_nit
      MTH_a <- MTH_a_nit
      MTH_b <- MTH_b_nit
      MTH_c <- MTH_c_nit
      DBH_model <- DBH_model_nit
      DBH_form <- DBH_form_nit
      DBH_a <- DBH_a_nit
      DBH_b <- DBH_b_nit
      DBH_c <- DBH_c_nit
      DBH_d <- DBH_d_nit
      DBH_f <- DBH_f_nit
      DBH_g <- DBH_g_nit
      DBH_h <- DBH_h_nit
      DBH_k <- DBH_k_nit
      MTH_MnHt_a <- MTH_MnHt_a_nit
      MTH_MnHt_b <- MTH_MnHt_b_nit
      VOL_type <- VOL_type_nit
      VOL_u <- VOL_u_nit
      VOL_v <- VOL_v_nit
      VOL_w <- VOL_w_nit
      VOL_z <- VOL_z_nit
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_nit_SI
      MORT_m <- MORT_m_nit_SI
      MORT_n <- MORT_n_nit_SI
      Den_a <- DEN_a_nit
      Den_b <- DEN_b_nit
      
    } else if (Species == "Eucalyptus delegatensis") {
      MTH_model <- MTH_model_del
      MTH_form <- MTH_form_del
      MTH_a <- MTH_a_del
      MTH_b <- MTH_b_del
      MTH_c <- MTH_c_del
      DBH_model <- DBH_model_del
      DBH_form <- DBH_form_del
      DBH_a <- DBH_a_del
      DBH_b <- DBH_b_del
      DBH_c <- DBH_c_del
      DBH_d <- DBH_d_del
      DBH_f <- DBH_f_del
      DBH_g <- DBH_g_del
      DBH_h <- DBH_h_del
      DBH_k <- DBH_k_del
      MTH_MnHt_a <- MTH_MnHt_a_del
      MTH_MnHt_b <- MTH_MnHt_b_del
      VOL_type <- VOL_type_del
      VOL_u <- VOL_u_del
      VOL_v <- VOL_v_del
      VOL_w <- VOL_w_del
      VOL_z <- VOL_z_del
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_del
      MORT_m <- MORT_m_del
      MORT_n <- MORT_n_del
      Den_a <- DEN_a_del
      Den_b <- DEN_b_del
      
    } else if (Species == "Eucalyptus saligna") {
      MTH_model <- MTH_model_sal
      MTH_form <- MTH_form_sal
      MTH_a <- MTH_a_sal
      MTH_b <- MTH_b_sal
      MTH_c <- MTH_c_sal
      DBH_model <- DBH_model_sal
      DBH_form <- DBH_form_sal
      DBH_a <- DBH_a_sal
      DBH_b <- DBH_b_sal
      DBH_c <- DBH_c_sal
      DBH_d <- DBH_d_sal
      DBH_f <- DBH_f_sal
      DBH_g <- DBH_g_sal
      DBH_h <- DBH_h_sal
      DBH_k <- DBH_k_sal
      MTH_MnHt_a <- MTH_MnHt_a_sal
      MTH_MnHt_b <- MTH_MnHt_b_sal
      VOL_type <- VOL_type_sal
      VOL_u <- VOL_u_sal
      VOL_v <- VOL_v_sal
      VOL_w <- VOL_w_sal
      VOL_z <- VOL_z_sal
      THINCOEF <- THINCOEF_euc
      MORT_k <- MORT_k_sal
      MORT_m <- MORT_m_sal
      MORT_n <- MORT_n_sal
      Den_a <- DEN_a_sal
      Den_b <- DEN_b_sal
      
    } else if (Species == "Radiata pine") {
      MTH_MnHt_a <- MTH_MnHt_a_rad
      MTH_MnHt_b <- MTH_MnHt_b_rad
      VOL_type <- VOL_type_rad
      VOL_u <- VOL_u_rad
      VOL_v <- VOL_v_rad
      VOL_w <- VOL_w_rad
      VOL_z <- VOL_z_rad
      THINCOEF <- THINCOEF_rad
      
    } else if (Species == "Douglas-fir") {
      MTH_MnHt_a <- MTH_MnHt_a_dfr
      MTH_MnHt_b <- MTH_MnHt_b_dfr
      VOL_type <- VOL_type_dfr
      VOL_u <- VOL_u_dfr
      VOL_v <- VOL_v_dfr
      VOL_w <- VOL_w_dfr
      VOL_z <- VOL_z_dfr
      THINCOEF <- THINCOEF_dfr
    }
    
  } # Species-specific parameters

  # Export all species-specific variables to MODEL_ENV so downstream
  # functions (AgeBH, DBH_mod, CalcVol, etc.) can find them.
  vars_to_export <- c("MTH_model","MTH_form","MTH_a","MTH_b","MTH_c",
                       "DBH_model","DBH_form","DBH_a","DBH_b","DBH_c",
                       "DBH_d","DBH_f","DBH_g","DBH_h","DBH_k",
                       "MTH_MnHt_a","MTH_MnHt_b",
                       "VOL_type","VOL_u","VOL_v","VOL_w","VOL_z",
                       "THINCOEF","MORT_k","MORT_m","MORT_n",
                       "Den_a","Den_b",
                       "WoodDensity_Adjustment")
  for (vn in vars_to_export) {
    if (exists(vn, inherits = FALSE))
      assign(vn, get(vn), envir = MODEL_ENV)
  }

  rotlength <- input_data[6, 4]
  if (is.na(rotlength) || rotlength < 1) stop("Input_parameters: rotlength (input_data[6,4]) is NA or < 1 — check c_change_control.csv 1st Rotation")
  if (rotlength > 200) rotlength <- 200  # Maximum allowed rotation length is 200 years
  
  
  # Create the Thinning_schedule dataframe by copying the specified portion of input_data
  Thinning_schedule <- cbind(rep(0, 5),as.data.frame(input_data[9:13, 4:7]))
  Thinning_schedule[is.na(Thinning_schedule)] <- 0    # Replace all NA values in the Thinning_schedule dataframe with 0
  colnames(Thinning_schedule) <- c("Start","Thin1", "Thin2",  "Thin3", "Thin4")
  rownames(Thinning_schedule) <- c("Stock_hist_Type", "thin_age", "Stock_hist_T",  "Stock_hist_N", "Stock_hist_thin_coeff")
  Thinning_schedule["Stock_hist_N", 1] <- input_data[5, 4] 
  Thinning_schedule["Stock_hist_T", 1] <- 0
  Thinning_schedule["thin_age",]<-  Thinning_schedule["Stock_hist_T",]
  # Replace any value in columns 2 to 5 with THINCOEF if it is 0, -999, or NA
  Thinning_schedule["Stock_hist_thin_coeff", 2:5] <- ifelse(
    is.na(Thinning_schedule["Stock_hist_thin_coeff", 2:5]) | 
      Thinning_schedule["Stock_hist_thin_coeff", 2:5] == 0 | 
      Thinning_schedule["Stock_hist_thin_coeff", 2:5] == -999, 
    THINCOEF, 
    Thinning_schedule["Stock_hist_thin_coeff", 2:5] )
  #Thinning_schedule[] <- lapply(Thinning_schedule, as.numeric)  # Convert all columns in Thinning_schedule to numeric
  Nthins <- sum(Thinning_schedule["Stock_hist_T", 2:5]>0)  
  assign("Nthins", Nthins, envir = MODEL_ENV)  
  Thinning_schedule<-cbind(Thinning_schedule,"end"=0)
  Thinning_schedule["Stock_hist_Type", "end"]<- 1
  Thinning_schedule["Stock_hist_thin_coeff", "end"]<- 1
  assign("Thinning_schedule", Thinning_schedule, envir = MODEL_ENV)  
  
  
  mode <- input_data[16, 3]  # Mode: 1 = Use specified indices, 2 = Calibrate using stand metrics, 3 = Calibrate using tree list
  assign("mode", mode, envir = MODEL_ENV)  
    DiaDist <- input_data[16, 4]  # Diameter distribution method: 1 = Weibull, 2 = Derive from tree list
    assign("DiaDist", DiaDist, envir = MODEL_ENV)  
    
  if (input_data[26, 4] != -999) {MORT_k <- input_data[26, 4] / 100}
  
  weibull_CV <- input_data[27, 4]
  if (weibull_CV == -999) {weibull_CV <- 0.27}  # Default DBH CV
  weibull_b <- 1.010369 * weibull_CV ^ (-1.078517)  # Approximate Weibull b parameter from CV
    assign("weibull_CV", weibull_CV, envir = MODEL_ENV)  
    assign("weibull_b", weibull_b, envir = MODEL_ENV)  
  
  
  
  # Taper function coefficients
  alpha0 <- alpha0_458
  alpha1 <- alpha1_458
  alpha2 <- alpha2_458
  beta1 <- beta1_458
  beta2 <- beta2_458
  beta3 <- beta3_458
  beta4 <- beta4_458
  beta5 <- beta5_458
  
  # Other parameters
  log_length <- input_data[28, 4]
  if (log_length == -999) log_length <- 6
   min_SED <- input_data[29, 4]
   if (min_SED == -999) min_SED <- 150
   break_height <- input_data[30, 4] / 100
  if (break_height == -999) break_height <- 0.65
   log_losses <- input_data[31, 4]
  if (log_losses == -999) log_losses <- 4
  
  WoodDensity_Adjustment <- 1  # No wood density adjustment
  if (input_data[32, 4] != -999) {WoodDensity_Adjustment <- 1 + input_data[32, 4] / 100 }
  
  AGCWD_half_life <- input_data[33, 4]
  if (AGCWD_half_life == -999) AGCWD_half_life <- 15  # Natural forest NZ
  BGCWD_half_life <- input_data[34, 4]
  if (BGCWD_half_life == -999) BGCWD_half_life <- 15  # Natural forest NZ
  latitude <- input_data[35, 4]
  if (latitude == -999) latitude <- 36
  elevation <- input_data[36, 4]
  if (elevation == -999) elevation <- 200
  Soil_C <- input_data[37, 4]
  if (Soil_C == -999) Soil_C <- 5.57
  Soil_N <- input_data[38, 4]
  if (Soil_N == -999) Soil_N <- 0.296
  MAT <- input_data[39, 4]
  if (MAT == -999) MAT <- 12
  drift <- input_data[40, 4]
  if (drift == -999) drift <- 0
  
    # Pruning information
  PRUNEHT <- 0 #Final Prune Height ????
  
  
  # Create the Pruning_schedule dataframe by copying the specified portion of input_data
  Pruning_schedule <- as.data.frame(input_data[9:11, 11:14])
  Pruning_schedule[is.na(Pruning_schedule)] <- 0    # Replace all NA values in the Pruning_schedule dataframe with 0
 colnames(Pruning_schedule) <- c("Lift1", "Lift2",  "Lift3", "Lift4")
  rownames(Pruning_schedule) <- c("prune_age", "prune_N", "prune_height")
  
    # Assign the value of the last height to PRUNEHT
  for (i in 1:4) {
    if (Pruning_schedule["prune_height", i] != 0) {PRUNEHT <- Pruning_schedule["prune_height", i]
      break}}
  
  assign("Pruning_schedule", Pruning_schedule, envir = MODEL_ENV)  
  

  
  # Radiata pine: Copy inputs into 300 Index worksheet
  if (Species == "Radiata pine") {
    # Clear specific columns 
    # Fill in the specific cells with data
    data_300_index[3, 3] <- I300
    data_300_index[4, 3] <- H30
    data_300_index[3, 6] <- latitude
    data_300_index[4, 6] <- elevation
    data_300_index[75,4] <- Soil_C
    data_300_index[76,4] <- Soil_N
    data_300_index[77,4] <- MAT
    data_300_index[7, 3] <- T1
    data_300_index[8, 3] <- N1
    data_300_index[14,3] <- T1
    data_300_index[15,3] <- H1
    data_300_index[10,3] <- N1 * pi * (D1 / 200) ^ 2
    data_300_index[19,3] <- Thinning_schedule["Stock_hist_N", 1]
    
    Thinning_schedule["Stock_hist_N", i]
    
    # Loop through thinning history
    for (i in 2:5) {
      if (Thinning_schedule["Stock_hist_N", i] != 0) {
        data_300_index[19 + i, "B"] <- Thinning_schedule["Stock_hist_T", i]  # Ages of thinning in integer years
        data_300_index[19 + i, "D"] <- Thinning_schedule["Stock_hist_N", i]
        data_300_index[19 + i, "E"] <- ifelse(Thinning_schedule["Stock_hist_thin_coeff", i] == -999, THINCOEF, Thinning_schedule["Stock_hist_thin_coeff", i])
      }
    }
    
    # Loop through pruning information
    for (i in 1:4) {
      if (Pruning_schedule["prune_age", i] != 0) {
        data_300_index[39 + i, "B"] <- Pruning_schedule["prune_age", i]
        data_300_index[39 + i, "C"] <- Pruning_schedule["prune_height", i]
        data_300_index[39 + i, "D"] <- Pruning_schedule["prune_N", i]
      }
    }
    
    # Rotation length
    data_300_index[47, 3] <- rotlength
  }
  
  # Douglas-fir: Copy inputs into the "500 Index" dataframe (same logic if needed)
  if (Species == "Douglas-fir") {
    # Implement similar logic if you have a "500 Index" dataframe like for Radiata pine
  }
  
  # Export remaining local variables to MODEL_ENV so downstream functions
  # (Yield_Table_radiata, density, etc.) can find them.
  vars_to_export2 <- c("WoodDensity_Adjustment", "log_length", "min_SED",
                        "break_height", "log_losses", "AGCWD_half_life",
                        "BGCWD_half_life", "latitude", "elevation",
                        "Soil_C", "Soil_N", "MAT", "drift", "PRUNEHT",
                        "rotlength", "Cali")
  for (vn in vars_to_export2) {
    if (exists(vn, inherits = FALSE))
      assign(vn, get(vn), envir = MODEL_ENV)
  }

  # Return the modified "300 Index" dataframe
  return(data_300_index) 
}
Input_tree_list <- function() {
  # Generate unscaled DBHs for tree list - either using Weibull distribution or reading from user-supplied tree list
  if (DiaDist == 1) {
    # Generate 100 stems from Weibull distribution
    nstems <- 100
    proportion <- 0.005
    Treelist <- matrix(0, nstems, 6)
    for (tree in 1:nstems) {
      Treelist[tree, 3] <- (-log(1 - proportion))^(1 / weibull_b)
      proportion <- proportion + 0.01
    }
  } else if (DiaDist == 2) {
    # Use simple scaling for DBH in tree list based on analysis showing DBH CV remains constant over time
    nstems <- nrow(Starting_tree_list)
    if (nstems > 10000) nstems <- 1  # Set nstems to 1 if greater than 10,000
    if (nstems > 0 && nstems < 10000) { Treelist <- (Starting_tree_list)} # removed matrix(Starting_tree_list)
  }
  
  return(Treelist)
} ###2 DiaDist, weibull_b, Starting_tree_list
siteIndex <- function() {
  # Error Check ######## if (!checkinput_htage() || !checkinput_htfn()) return(NULL)
  
  #checkinput_htage<- check_input_htage(input_data, implementation)
  #checkinput_htfn<-check_input_htfn(input_data, implementation)
  #if(!checkinput_htage) print("htage error")
  #if(!checkinput_htfn) print("htfn error")
      
  
  HAge <- data_300_index[14, 3]
  HMTH <- data_300_index[15, 3]
  
  if (HAge == 20) {
    SI <- HMTH
  } else {
    heightmodel <- heightmod()  # Determine height model
    assign("heightmodel", heightmodel, envir = MODEL_ENV)  
    
    SI <- BisectionFn(5, 60, 15, 2, HMTH, HAge, 0, 0)
  }
  
  data_300_index[4, 3] <- SI
  assign("data_300_index", data_300_index, envir = MODEL_ENV)  
  
}
Calc300Index <- function() {
  
  #check if all ionput data exists in errorchecking  sub
  #I300, SI, htfn, initialstock, stocking, prune, fellage,steplth, volfn, mortfn
 
  #Call  DONE

  # Function to check input for I300 index measurements
#if (!check_input_I300(input_data, implementation)) {return(NULL)}
#if (!check_input_SI(input_data, implementation)) {return(NULL)}
# if (!checkinput_htfn(input_data, implementation)) {return(NULL)}
# if (!checkinput_initialstock(input_data, implementation)) {return(NULL)}
# if (!checkinput_stocking(input_data, implementation)) {return(NULL)}
# if (!checkinput_prune(input_data, implementation)) {return(NULL)}
# if (!checkinput_fellage(input_data, implementation)) {return(NULL)}
# if (!checkinput_steplth(input_data, implementation)) {return(NULL)}
# if (!checkinput_volfn(input_data, implementation)) {return(NULL)}
# if (!checkinput_mortfn(input_data, implementation)) {return(NULL)}
    # Extract necessary values from the Excel sheet
  
  Inputparms()
  voltab()
  age300 <- data_300_index[7, 3] 
  assign("age300", age300, envir = MODEL_ENV)
  
  maxage <- age300
  assign("maxage", maxage, envir = MODEL_ENV)
  
  Stock300 <- data_300_index[8, 3] 
  assign("Stock300", Stock300, envir = MODEL_ENV)
  
  HAge <- data_300_index[14, 3]
  assign("HAge", HAge, envir = MODEL_ENV)
  
  HMTH <- data_300_index[15, 3]
  assign("HMTH", HMTH, envir = MODEL_ENV)
  
  # Calculate maxage and steps
  steps <- round(as.numeric(maxage / steplength))
  assign("steps", steps, envir = MODEL_ENV)
  
  # Calculate MTH300 using CalcMTH function (assuming it's defined elsewhere)
  MTH300 <- CalcMTH(SI, age300)  # SI should be defined earlier in the script
  assign("MTH300", MTH300, envir = MODEL_ENV)
  
  # Get DBH300 from Excel sheet
  DBH300 <- data_300_index[9, 3]
  BA300  <- data_300_index[10, 3]
  if (is.na(BA300)) BA300 <- 0

  # If DBH300 is 0 or NA, calculate it using BA300 or Vol300
  if (is.na(DBH300) || DBH300 == 0) {
    if (BA300 != 0) {
      DBH300 <- CalcDBHfromBA(BA300, Stock300)
    } else {
      Vol300 <- data_300_index[11, 3]
      DBH300 <- CalcDBHfromBA(calcBAfromVol(MTH300, Vol300, Stock300), Stock300)
    }
  } else {
    BA300 <- Stock300 * pi * (DBH300 / 200)^2
  }
  assign("DBH300", DBH300, envir = MODEL_ENV)
  assign("BA300", BA300, envir = MODEL_ENV)
  
  
  #I300<-0
  #Index300()  # This function calculates I300 Replaced with below
  I300 <- BisectionFn(1.328, 60, 14, 1, 0, 0, 0, 0)
  assign("I300", I300, envir = MODEL_ENV)

  # Write I300 back into the synthetic matrix so subsequent Inputparms() calls
  # (inside OutputGrowth / CalcOffsets) pick up the computed value, not NA.
  data_300_index[3, 3] <- I300
  assign("data_300_index", data_300_index, envir = MODEL_ENV)
}
Index300 <- function() {
  # Calculate 300 Index using Bisection method (assuming Bisection is defined)
  I300 <- BisectionFn(1.328, 60, 14, 1, 0, 0, 0, 0)  # I300 is updated globally
}

#xlower=1.328; xupper=60; niterations=14; fnno=1; p1=0; p2=0; p3=0; p4=0; X=xlower
BisectionFn <- function(xlower, xupper, niterations, fnno, p1, p2, p3, p4) {
 # 'Find when function number fnno equals zero using the bisection method  
  xA <- xlower
  FA <- fn(xA, fnno, p1, p2, p3, p4) #values checked
  
  xB <- xupper
  FB <- fn(xB, fnno, p1, p2, p3, p4) #values checked
  
  for (j in 1:niterations) {
    xC <- (xA + xB) / 2
    FC <- fn(xC, fnno, p1, p2, p3, p4)
    if (FA * FC < 0) {
      xB <- xC
      FB <- FC
    } else {
      xA <- xC
      FA <- FC
    }
  }
  
    return(xC)
} # Seems to Work
fn <- function(X, fnno, p1, p2, p3, p4) {
  # Function to be zeroed using the bisection method
  if (fnno == 1) {
    I300 <<- X
    gr <- Growth(FALSE, I300)
    return(DBH300 - gr$DBH_end)
  } else if (fnno == 2) {
    return(p1 - CalcMTH(X, p2))  # p1 = MTH, p2 = age
  } else if (fnno == 3) {
    return(p4 - DBHmodelFn(p1, p2, X, p3))  # p1 = A200, p2 = SI, p3 = stock, p4 = DBH
  } else if (fnno == 4) {
    return(p4 - DBHmodelFn(X, p1, p2, p3))  # p1 = SI, p2 = Age, p3 = stock, p4 = DBH
  }
} # seems to Work

Growth <- function(OUTPUT, I300) {
  # Initialise variables
  N <- initialstocking
  shist <- 1
  thin <- 0
  lift <- 0
  Age <- 0 
  Nelements <- 1
  A200 <- CalcA200start(Age, I300, SI)
  DBH <- 0
  Meanht <<- numeric(11)
  adjageel <<- numeric(11)
  initiallag <<- numeric(9)
  ThinLag <<- numeric(9)
  agethin <<- numeric(9)
  # Loop to reset dbhelement array (size 10)
  dbhelement <- rep(0, 10)
  Vol <- 0.0000064 * N  # Volume of seedling at planting (Beets)
  BA <- 0
  MTH <- 0.25
  mnheight <- 0.25
  nelement <- rep(N, Nelements)
  ncum <- rep(N, Nelements)
  PRHT <- rep(0, Nelements)
  prlag <- rep(0, Nelements)
  totalthinlag <- 0
  sellag <- rep(0, Nelements)
  total_prlag <- rep(0, Nelements)
  outputline <- 5
  lineprinted <- FALSE
  
  # Set offsets to neutral values if implementation is not 2
  if (implementation != 2) {
    DBHsqd_add_offset <- 0
    DBHsqd_mult_offset <- 1
    MTH_add_offset <- 0
    MTH_mult_offset <- 1
    DBH_calibration_age <- 0
    MTH_calibration_age <- 0
  }
  
  # Print age zero if OUTPUT is TRUE
  if (OUTPUT) OutStep()
  
  output_rows <- list()
  if (OUTPUT && file.exists("output.csv")) file.remove("output.csv")

  for (j in 1:steps) {
     # Save previous stand parameters
  tl_prev_standDBH <- DBH
  tl_prev_standN <- N
  tl_prev_standBA <- BA
  tl_prev_standmnheight <- mnheight
  tl_prev_standage <- Age
  
  # Update Age and calculate A200
  Age <- Age + steplength
  A200 <- CalcA200start(Age, I300, SI)
  stock_out <- stock(Nelements, nelement, ncum, N, shist, DBH)
  N <- stock_out$N
  nelement <- stock_out$nelement
  ncum <- stock_out$ncum
  HeightCalc<- (Height(N, nelement, ncum, Nelements, Age))
  mnheight <-   HeightCalc$mnheight
  MTH <-HeightCalc$MTH
  Meanht <<- HeightCalc$Meanht
  adjage <- Ageshifts(N, PRHT, prlag, total_prlag, thin, sellag, Nelements, Age, nelement) #?????????
  DBH <- Diameter(ncum, dbhelement, Nelements, I300,nelement)
  BA <- CalcBAfromDBH(DBH, N)
  Vol <- CalcVol(MTH, BA, N)
    
    # Create a single row for the current iteration
   { iteration_data <- data.frame(
      Iteration = j,
      Age = Age,
      A200 = A200,
      N = N,
      mnheight = mnheight,
      MTH = MTH,
      DBH = DBH,
      BA = BA,
      Vol = Vol
    )}
     # Append to CSV file without overwriting
    output_rows[[length(output_rows) + 1]] <- iteration_data
    
    # Print step if certain conditions are met
    if (OUTPUT && (abs(Age - floor(Age)) < 0.001 || abs(Age - maxage) < 0.001)) {
      OutStep()
    }
    # Thinning and lifting
    if (Stocking_history$shist_N2[shist] != 0 && Age >= Stocking_history$shist_T[shist] - 0.001) {
      if (OUTPUT && !lineprinted) OutStep()
      thin <- thin + 1
      thin_out <- thinning(N, DBH, shist, Nelements, dbhelement, nelement, ncum, adjage, thin, totalthinlag, sellag, prlag, I300, Age, MTH)
      N <- thin_out$N
      DBH <- thin_out$DBH
      Nelements <- thin_out$Nelements
      dbhelement <- thin_out$dbhelement
      nelement <- thin_out$nelement
      ncum <- thin_out$ncum
      adjage <- thin_out$adjage
      totalthinlag <- thin_out$totalthinlag
      sellag <- thin_out$sellag
      BA <- thin_out$BA
      Vol <- thin_out$Vol
      mnheight <- thin_out$mnheight
      if (OUTPUT) {
        OutThin()
        OutElements()
      }
    }
    if (lift < Nlifts && Age >= Pruning_history$lift_T[lift + 1] - 0.001) {
      if (OUTPUT && !lineprinted) OutStep()
      lift <- lift + 1
      lift_out <- Newlift(lift, Nelements, PRHT, nelement, ncum, total_prlag, prlag, dbhelement, Meanht, adjageel, sellag, MTH, A200, N, SI, pra, prb, prc, Age, totalthinlag, Pruning_history)
      Nelements <- lift_out$Nelements
      PRHT <- lift_out$PRHT
      nelement <- lift_out$nelement
      ncum <- lift_out$ncum
      total_prlag <- lift_out$total_prlag
      prlag <- lift_out$prlag
      dbhelement <- lift_out$dbhelement
      Meanht <- lift_out$Meanht
      adjageel <- lift_out$adjageel
      sellag <- lift_out$sellag
      if (OUTPUT) {
        OutPrune()
        OutElements()
      }
    }
    if (shist < Nshist && Age >= Stocking_history$shist_T[shist] - 0.001) {
      shist <- shist + 1
    }
    if (lineprinted) {
      outputline <- outputline + 1
      lineprinted <- FALSE
    }
  }
  growth_df <- if (length(output_rows)) do.call(rbind, output_rows) else data.frame()
  if (OUTPUT && nrow(growth_df) > 0) {
    write.csv(growth_df, "output.csv", row.names = FALSE)
  }
  list(DBH_end = DBH, MTH_end = MTH, BA_end = BA, Vol_end = Vol, N_end = N, growth_df = growth_df)
}  #??? Working

Calibrate_radiata <- function() {
  siteIndex()
  Calc300Index()
  OutputGrowth()

  I300 <- data_300_index[3, 3]
  H30  <- data_300_index[4, 3]
  assign("I300", I300, envir = MODEL_ENV)
  assign("H30",  H30,  envir = MODEL_ENV)
}
CalcMTH <- function(SI, HAge) {
  
  height_coeffs<- calcheightcoeff(SI, heightmodel)
  ha <- height_coeffs$ha
  hb <- height_coeffs$hb
    0.25 + (SI - 0.25) * ((1 - exp(-ha * HAge)) / (1 - exp(-ha * 20)))^hb
} # Working
CalcA200start <- function(Age, I300, SI) {
  # Calculate A200 from the 300 Index and SI
  adjI300 <- I300 #'Standard model
  b_adj <- 0.0206 / 19.488
  c_adj <- -0.0182 / 100 / 19.488
  k1_adj <- 25
  k2_adj <- 55
  k3_adj <- 215.97
  k4_adj <- -0.05532
  k_adj <- k3_adj * exp(k4_adj * I300)
  
  #'300 Index bias correction for ages<6.77
  if (bias_young && Age < 6.77) {
    i300adjustment <- 180.5 * adjI300^(-3.256) * (Age - 6.77)^2
    if (i300adjustment > 5) i300adjustment <- 5
    adjI300 <- adjI300 + i300adjustment
  }
  
  if (bias_SI) {
    if (SI < 25 && SI >= 15) adjI300 <- adjI300 * (30 - 0.02 * (25 - SI) * (Age - 28.6)) / 30
    if (SI < 15) adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
    if (SI > 35 && SI <= 45) adjI300 <- adjI300 * (30 - 0.02 * (SI - 35) * (Age - 28.6)) / 30
    if (SI > 45) adjI300 <- adjI300 * (30 - 0.2 * (Age - 28.6)) / 30
  }
  
  if (Age < 30) {
    adjI300 <- adjI300 * (30 + drift * (Age - 28.6)) / 30
  }
  
  BA300_30 <- calcBAfromVol(CalcMTH(SI, 30), adjI300 * 30, 300)
  DBH300_30 <- CalcDBHfromBA(BA300_30, 300)
  CalcA200start<-CalcA200(DBH300_30, 28.7, 300, SI)
  return(CalcA200start)
} # Working if set parameters used
MTD <- function(Treelist, nstems, stocking) {
  # Calculate mean top diameter of stem measurements in TreeList
  nMTD <- 100
  sumWt <- 0
  sumDBH2Wt <- 0
  
  # Sort Treelist based on DBH values
  sorted_Treelist <- Treelist[order(Treelist[, 3], decreasing = FALSE), ]
  
  for (j in nstems:1) {
    if (sumWt + stocking / nstems > nMTD) {
      wt <- nMTD - sumWt
    } else {
      wt <- stocking / nstems
    }
    sumWt <- sumWt + wt
    sumDBH2Wt <- sumDBH2Wt + wt * sorted_Treelist[j, 3]^2
    if (sumWt >= nMTD) break
  }
  
  MTD <- sqrt(sumDBH2Wt / sumWt)
  return(MTD)
}
FitPettersonType1 <- function(Treelist, nstems) {
  # Fit type 1 Petterson Height / DBH curve to stem measurements in TreeList
  Nheights <- 0
  sum_x <- 0
  sum_y <- 0
  sum_x2 <- 0
  sum_y2 <- 0
  sum_xy <- 0
  
  qDBH <- sqrt(sum(Treelist[, 3]^2)/nstems)
  BA <- (nstems * pi * ((qDBH / 200)^2))/Plot_area
  
  # Filter for valid stems where neither Height nor DBH are zero or NA
  valid_stems <- Treelist[!is.na(Treelist[, 4]) & Treelist[, 4] != 0 & 
                            !is.na(Treelist[, 3]) & Treelist[, 3] != 0, ]
  
  # Calculate Heights > 1.4m for valid stems
  Y_values <- valid_stems[, 3] / (valid_stems[, 4] - 1.4)^0.4
  X_values <- valid_stems[, 3] #dbh
  
  # Count the number of valid heights
  Nheights <- nrow(valid_stems)
  
  # Calculate the sums using vectorized operations
  sum_x <- sum(X_values)
  sum_y <- sum(Y_values)
  sum_x2 <- sum(X_values^2)
  sum_y2 <- sum(Y_values^2)
  sum_xy <- sum(X_values * Y_values)
  
  petA <- 0
  petB <- 0
  
  if (Nheights > 1) {
    petA <- (sum_xy - sum_x * sum_y / Nheights) / (sum_x2 - sum_x^2 / Nheights)
    petB <- sum_y / Nheights - petA * (sum_x / Nheights)
    
    if (petB < 0) {
      petB <- 0
      petA <- sum_y / sum_x
    }
    if (petA < 0) {
      petA <- 0
      petB <- sum_y / Nheights
    }
  }
  
  return(list(petA = petA, petB = petB, qDBH = qDBH, BA = BA))
}
Process_tree_list <- function(Treelist, Plot_area, Age, nstems) {
  # Calculate stocking
  stocking <- nstems / Plot_area
  if (nstems > 10000) { nstems <- 1 }
  
  # Insert the new column at the beginning of Treelist
  Treelist <- data.frame(NewColumn = rep(1/Plot_area, nrow(Treelist)), Treelist) # Add a new column of 1/plot area
  
  MTDia <- MTD(Treelist, nstems, stocking)
  result <- FitPettersonType1(Treelist, nstems)
  petA <- result$petA
  petB <- result$petB
  BA <- result$BA
  qDBH <- result$qDBH
  
  MTH <- 1.4 + (petA + petB / MTDia) ^ (-2.5)
  
  # Return the results or perform further actions
  return(list(MTDia = MTDia, stocking = stocking, petA = petA, petB=petB, qDBH = qDBH, BA = BA, MTH = MTH, Age = Age))
}

populate_300index_inputs_from_tree <- function(tree_metrics) {
  # Optional enrichment: tree-derived values populate summary inputs.
  data_300_index[14, 3] <<- tree_metrics$Age
  data_300_index[15, 3] <<- tree_metrics$MTH
  data_300_index[8, 3]  <<- tree_metrics$stocking
  data_300_index[9, 3]  <<- tree_metrics$qDBH
  data_300_index[10, 3] <<- tree_metrics$BA
}
MTH_from_MnHt <- function(MeanHeight, N, MTH_MnHt_a, MTH_MnHt_b) {
  return(MeanHeight / (1 - MTH_MnHt_a * (1 - exp(MTH_MnHt_b * (N - 100)))))
}
heightmod <- function() {
  # Determine height model - 1 = NSW, 2 = Simple NZ, 3 = Environmental NZ
  if (!is.na(data_300_indexX[14, 1]) && tolower(data_300_indexX[14, 1]) == "x") {
    return(1)  # NSW model
    } else if (is.na(data_300_index[3, 6]) || is.na(data_300_index[4, 6]) ||(data_300_index[3, 6] < 30 || data_300_index[3, 6] > 48)) {
    return(2)  # Simple NZ model after 'Test whether latitude is present and within NZ range
      } else {
    return(3)  # Environmental NZ model
  }
} # Working
calcheightcoeff <- function(SI, heightmodel) {
  if (heightmodel == 1) {
    ha <- exp(hNSWa)
    hb <- 1 / (hNSWb + hNSWp * SI)
  } else if (heightmodel == 2) {
    ha <- exp(ha0 + ha1 * SI)
    hb <- 1 / (hb0 + hb1 * SI)
  } else {
    latitude <- (data_300_index[3, 6])
    altitude <- (data_300_index[4, 6])
    ha <- exp(hae0 + hae1 * latitude + hae2 * altitude)
    hb <- 1 / (hbe0 + hbe1 * SI)
  }
  return(list(ha = ha, hb = hb))
} # Working
mort <- function() {
  if(Nshist==0) {Nshist <- Nshist + 1}  #Array starts at 1
    # Check and extend stand history if necessary
  if (Stocking_history$shist_T[Nshist] < maxage) {
    Nshist <- Nshist + 1  # Use <- to modify the global variable Nshist
    Stocking_history$shist_T[Nshist] <- maxage  # Update the maximum age in the dataframe
  }

  # Initialize previous age and stocking
  prevage <- 0
  prevN <- initialstocking  # Ensure this variable is defined in the global environment
  
  # Calculate mortality for each stand history element
  for (shist in 2:Nshist) {
    if (is.na(Stocking_history$shist_N1[shist]) || Stocking_history$shist_N1[shist] == 0) {
      Stocking_history$Mortality[shist] <- -1  # Will calculate mortality rate later
    } else {
      Stocking_history$Mortality[shist] <- 100 * log(prevN / Stocking_history$shist_N1[shist]) / 
        (Stocking_history$shist_T[shist] - prevage)
    }
    prevage <- Stocking_history$shist_T[shist]  # Update previous age
    prevN <- ifelse(is.na(Stocking_history$shist_N2[shist]) || Stocking_history$shist_N2[shist] == 0, 
                    Stocking_history$shist_N1[shist], 
                    Stocking_history$shist_N2[shist])  # Update previous stocking
  }
  
  # Return the updated Stocking_history dataframe
  return(Stocking_history)
} # Working
Calcagezero <- function() {
  height_coeffs<- calcheightcoeff(SI, heightmodel)
  ha <- height_coeffs$ha
  hb <- height_coeffs$hb
  -log(-(1 - exp(-ha * 20)) * ((1.4 - 0.25) / (SI - 0.25))^(1 / hb) + 1) / ha 
  } # Working
CalcDBHfromBA <- function(BA, N) {
  sqrt(1.273 * BA / N) * 100
}# Working
calcBAfromVol <- function(MTH, Vol, stockn) {
  if (Vol <= 0 || MTH <= 1.6 || stockn <= 0) {
    return(0)
  } else if (voltable %in% c(1, 2)) {
    return(Vol / (MTH * (V[1, voltable] * (MTH - 1.4)^V[2, voltable] + V[3, voltable])))
  } else if (voltable %in% c(10, 11)) {
    return(Vol / (V[1,voltable] + V[2,voltable] * MTH + V[3,voltable] * stockn))
  } else {
      return(exp(V[1,voltable] + 
                 V[2,voltable] * log(MTH) + 
                 V[3,voltable] * log(Vol) + 
                 V[4,voltable] * log(stockn) + 
                 V[5,voltable] * log(stockn)^2 + 
                 V[6,voltable] * log(MTH)^2 + 
                 V[7,voltable] * log(MTH) * log(stockn) + 
                 V[8,voltable] * log(Vol) * log(stockn)))
  }
}# Working
DBHmodelFn <- function(A200, SI, Age, stockn) {
  # Predict DBH at given age and stocking using 300 Index model
  # A200 is DBH at 200 sph and age 30 with no pruning or thinning
  
  # Initialize variables
  stk <- stockn
  agezero <- Calcagezero()
  site_effect <- A200 / da1 - 1
  
  A <- da1 * (1 + site_effect)
  B <- db2 * (db1 + dbSI * (SI - 28) + dbdia * site_effect + dbsidia * (SI - 28) * site_effect)

  # Ensure B is within reasonable bounds
  if (B > -0.05) {
    B <- -0.05
  }
  
  # Handle cases where Age is less than agezero
  if (Age < agezero) {
    DBHmodel <- 0
  } else {
        D200 <- OldAgeCorrection(Age, agezero, B) * A * ((1 - exp(B * (Age - agezero))) / (1 - exp(B * (30 - agezero))))^dc
    
      # Modify q to eliminate unnatural behaviour at stockings near 200 sph
    if (stk > 220) {
      qq <- (log(stk) - log(200)) ^ dr2
    } else {
      qq <- 2 * (log(220) - log(200)) ^ dr2 - (log(242) - log(stk)) ^ dr2
    }
    
    q <- dr * (1 + drsi * (SI - 28)) * qq
    p <- dl + dm * stk + dn * site_effect
    
    # Calculate DBHmodel
    DBHmodel <- D200 - q * log(1 + exp(Ds * (D200 - p)))
    
    # High stocking correction, apply if BA is decreasing with stocking
    if (stk > 250) {
      if (dBA_dN(D200, p, q, stk) <= 0) {
        N_MaxBA <- MaxBAStocking(D200, site_effect, SI, stk)
        q <- dr * (1 + drsi * (SI - 28)) * sign(N_MaxBA - 200) * (abs(log(N_MaxBA) - log(200))) ^ dr2
        p <- dl + dm * N_MaxBA + dn * site_effect
        DBHmodel <- (D200 - q * log(1 + exp(Ds * (D200 - p)))) * sqrt(N_MaxBA / stk)
      }
    }
  }
  
  # Ensure DBHmodel is non-negative
  if (DBHmodel < 0) {
    DBHmodel <- 0
  }
  
  return(DBHmodel)
}# Working
OldAgeCorrection <- function(Age, agez, B) {
  T <- (Age - agez) - 25
  if (T < 0) T <- 0
  return(1 + 4.350585474 * (1 - exp(-0.001473784 * T))^0.973636099)
} # Working
approxDBH <- function(D200, p, q) {
  # Calculate the approximate DBH
  approxDBH <- D200 - q * Ds * (D200 - p)
  
  return(approxDBH)
}# Working
dBA_dN <- function(D200, p, q, N) {
  # Derivative of predicted BA with respect to stocking
  dp_dN <- dm
  dq_dN <- q * dr2 / N / (log(N) - log(200))
  dD_dN <- -Ds * D200 * dq_dN + Ds * p * dq_dN + Ds * q * dp_dN
  D <- approxDBH(D200, p, q)
  
  if (D < 0) {
    return(0)
  } else {
    return(D * (D + 2 * N * dD_dN))
  }
}# Working
DBH_mod <- function(T, D300_30, SI, TBH, N, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c, DBH_d, DBH_f, DBH_g, DBH_h, DBH_k, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c) {
  D300_est <- D300(T, TBH, D300_30, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c)
  Ntemp <- N
  if (Ntemp > 200 * exp(1 / DBH_d)) {
    Ntemp <- 200 * exp(1 / DBH_d)
  }
  DBH_mod <- D300_est - DBH_d / DBH_f * (log(Ntemp) - log(300)) * log(1 + exp(DBH_f * (D300_est - (DBH_g + DBH_h * (SI - 30) + DBH_k * log(Ntemp)))))
  return(DBH_mod)
} # seems to Work
CalcA200 <- function(DBH, Age, stockn, SI) {
  Age<-28.7
  stockn<-300
  # Calculate A200 from DBH, Age, Stocking, and SI using Bisection Method
  agezero <- Calcagezero()
  Bisection <- BisectionFn(10, 150, 20, 4, SI, Age, stockn, DBH)
} # Working


#N is decreasing quite a lot on the second run
stock <- function(Nelements, nelement, ncum,N, shist, DBH) {
      prevN <- N
  # Calculate new stocking based on mortality model
  if (Stocking_history$Mortality[shist+1] >= 0) {
    N <- prevN / exp(Stocking_history$Mortality[shist+1] * steplength / 100)
  } else if (mortmodel == 1) {
    mortrate <- mortNSW
    N <- prevN / exp(mortrate * steplength / 100)
  } else if (mortmodel == 2) {
    if (DBH == 0) {
      mortrate <- 0
    } else {
      X <- exp(mortb + morte * SI + mortc * (log(N) + mortd * log(DBH^2)))
      mortrate <- (morta + (1 - morta) * X / (1 + X)) * 100
      N <- prevN / exp(mortrate * steplength / 100)
    }
  } else if (mortmodel == 3) {
    if (DBH == 0) {
      mortrate <- 0
    } else {
      X <- exp(mortv + mortw * (log(N) + mortx * log(DBH)))
      mortrate <- (morty + (1 - morty) * X / (1 + X)) * 100
      N <- prevN / exp(mortrate * steplength / 100)
    }
  } else if (mortmodel == 4) {
    if (DBH == 0) {
      mortrate <- 0
    } else {
      X <- exp(mortp + mortq * I300 / SI + morts * (log(N) + mortt * log(DBH)))
      mortrate <- (attrition + (1 - attrition) * X / (1 + X)) * 100
      N <- prevN / exp(mortrate * steplength / 100)
    }
  } else if (mortmodel == 5) {
    if (DBH == 0) {
      mortrate <- 0
    } else {
      X <- exp(mortb1 + morte1 * I300 + mortf1 * SI + mortc1 * (log(N) + mortd1 * log(DBH)))
      mortrate <- (attrition + (1 - attrition) * X / (1 + X)) * 100
      N <- prevN / exp(mortrate * steplength / 100)
    }
  } else if (mortmodel == 6) {
    if (DBH == 0) {
      mortrate <- 0
    } else {
      sdi <- exp(mort2007_f * I300 + mort2007_g * SI + log(N) + mort2007_d * log(DBH / 100) + mort2007_h * (log(DBH / 100))^2) / 1000
      mortrate <- attrition * 100 + 100 * (1 + pctmortadj / 100) * (mort2007_a + mort2007_b * sdi^mort2007_c)
      mortrate <- pmin(pmax(mortrate, 0), 95)
      N <- prevN * (1 - mortrate / 100)^steplength
    }
  }
    # Apply mortality to each element
  for (el in 1:Nelements) {
    nelement[el] <- nelement[el] * N / prevN
    ncum[el] <- ncum[el] * N / prevN  }
  #assign("nelement", nelement, envir = MODEL_ENV)  
  #assign("ncum", ncum, envir = MODEL_ENV)
  return(list(N = N, nelement = nelement, ncum = ncum))
} # Works
calcMeanht <- function(MTH, stockn) {
  A <- 0.07
  B <- -0.00399
  if (!is.na(MTH) && !is.na(stockn)) {
    return(as.numeric(MTH) * (1 - A * (1 - exp(B * (stockn - 100)))))
  }
} # works check returned value
Height <- function(N, nelement, ncum, Nelements, Age) {
  # Calculate MTH and mean height for each element
  MTH <- (CalcMTH(SI, Age))
  mnheight <- calcMeanht(MTH, N)
  Meanht[1] <- calcMeanht(MTH, nelement[[1]])
  
  if (Nelements > 1) {
    for (el in 2:Nelements) {
      Meanht[el] <- (ncum[el] * calcMeanht(MTH, ncum[el]) -
                       ncum[el - 1] * calcMeanht(MTH, ncum[el - 1])) / 
        (ncum[el] - ncum[el - 1])
    }
  }
  #assign("MTH", MTH, envir = MODEL_ENV)  
  #assign("mnheight", mnheight, envir = MODEL_ENV)  
  #assign("Meanht", Meanht, envir = MODEL_ENV)
  return(list(MTH=MTH,mnheight=mnheight, Meanht=Meanht))
  } # working - to update check returned value
#Returns NA thin>0
Ageshifts <- function(N,PRHT, prlag, total_prlag, thin, sellag, Nelements,Age,nelement) {
  # Calculate pruning and thinning time shifts for each element
  for (el in 1:Nelements) {
    if (PRHT[el] > 0) {
      prlag[el] <- prlag[el] + 0.3 * steplength
    }
    prlag[el] <- min(prlag[el], total_prlag[el])
  }
  
  totalthinlag <- 0
  if (thin>0) {
    for (th in 1:thin) {
      timesincethin <- Age - agethin[th]
      ThinLag[th] <- initiallag[th] + min(initiallag[th], tha) * thb * (1 - exp(thc * timesincethin))
      totalthinlag <- totalthinlag + ThinLag[th]
    }}
  
  adjage <- 0
  for (el in 1:Nelements) {
    adjageel[el] <- Age - prlag[el] - sellag[el] - totalthinlag
    adjage <- adjage + as.numeric(adjageel[el]) * as.numeric(nelement[el])
    
  }
  adjage <- adjage / N
  assign("adjageel", adjageel, envir = MODEL_ENV)  
  return(adjage)
  } # Works - to update check returned value
Diameter <- function(ncum, dbhelement, Nelements, I300,nelement) {
  N <- (ncum[Nelements])
  dbhsqd <- 0
  
  for (el in 1:Nelements) {
    prevdbh_el <- dbhelement[el]
    dbhelement[el] <- CalcDBH(I300, SI, (adjageel[el]), N)
    
    if (dbhelement[el] < prevdbh_el) {
      dbhelement[el] <- prevdbh_el
    }
    
    dbhsqd <- dbhsqd + (nelement[el]) * dbhelement[el]^2
  }
  DBH <- sqrt(dbhsqd / N)
  return(DBH)
}# Works - to update check returned value
CalcBAfromDBH <- function(DBH, N) {
  N / 1.273 * (DBH / 100)^2
} # Works - to update check returned value
CalcVol <- function(MTH, BA, stockn) {
  # Check for invalid inputs
  if (BA <= 0 || MTH <= 1.6 || stockn <= 0) { 
    return(0)
  } else if (voltable %in% c(1, 2)) {
    # Switch the reference for voltable
    return(MTH * BA * (V[1, voltable] * (MTH - 1.4)^V[2, voltable] + V[3, voltable]))
  } else if (voltable %in% c(10, 11)) {
    return(BA * (V[1, voltable] + V[2, voltable] * MTH + V[3, voltable] * stockn))
  } else {
    return(exp(-(V[1, voltable] + 
                   V[2, voltable] * log(MTH) + 
                   V[4, voltable] * log(stockn) + 
                   V[5, voltable] * log(stockn)^2 + 
                   V[6, voltable] * log(MTH)^2 + 
                   V[7, voltable] * log(MTH) * log(stockn) - 
                   log(BA)) / 
                 (V[3, voltable] + V[8, voltable] * log(stockn))))
  }
}
#Error Here if (Age <= 20 || Age >= 40)
CalcDBH <- function(I300, SI, Age, stockn) {
  if (Age <= 20 || Age >= 40) {
    A200 <- CalcA200start(Age, I300, SI)
    return(DBHmodelFn(A200, SI, Age, stockn))
  } else {
    # Interpolation logic for age between 20 and 40
    DBH1 <- DBHmodelFn(CalcA200start(19.5, I300, SI), SI, 19.5, stockn)
    DBH2 <- DBHmodelFn(CalcA200start(20.5, I300, SI), SI, 20.5, stockn)
    DBH3 <- DBHmodelFn(CalcA200start(39.5, I300, SI), SI, 39.5, stockn)
    DBH4 <- DBHmodelFn(CalcA200start(40.5, I300, SI), SI, 40.5, stockn)
    Y0 <- (DBH1 + DBH2) / 2
    Y1 <- (DBH3 + DBH4) / 2
    Y0p <- (DBH2 - DBH1)
    Y1p <- (DBH4 - DBH3)
    A <- Y0
    B <- Y0p
    D <- (2 * (Y0 + Y0p * 20 - Y1) + 20 * (Y1p - Y0p)) / (20^3)
    C <- (Y1p - Y0p - 3 * D * 20^2) / (2 * 20)
    return(A + B * (Age - 20) + C * (Age - 20)^2 + D * (Age - 20)^3)
  }
}# works
Newlift <- function(lift, Nelements, PRHT, nelement, ncum, total_prlag, prlag, dbhelement, Meanht, adjageel, sellag, MTH, A200, N, SI, pra, prb, prc, Age, totalthinlag, Pruning_history) {
  if (Pruning_history$lift_sph[lift] + 0.0001 < nelement[1]) {
    Nelements <- Nelements + 1
    
    for (el in Nelements:2) {
      PRHT[el] <- PRHT[el - 1]
      nelement[el] <- nelement[el - 1]
      ncum[el] <- ncum[el - 1]
      total_prlag[el] <- total_prlag[el - 1]
      prlag[el] <- prlag[el - 1]
      dbhelement[el] <- dbhelement[el - 1]
      Meanht[el] <- Meanht[el - 1]
      adjageel[el] <- adjageel[el - 1]
      sellag[el] <- sellag[el - 1]
    }
    
    PRHT[1] <- Pruning_history$lift_height[lift]
    nelement[1] <- Pruning_history$lift_sph[lift]
    nelement[2] <- nelement[2] - nelement[1]
    ncum[1] <- nelement[1]
    Meanht[1] <- calcMeanht(MTH, nelement[1])
    crlth <- Meanht[1] - PRHT[1]
    dbhb4pr <- dbhelement[1]
    
    prunecoeff <- ifelse(Pruning_history$lift_prunecoeff[lift] != 0, Pruning_history$lift_prunecoeff[lift], thincoeff)
    
    dbhelement[1] <- dbhelement[1] * (nelement[1] / ncum[2])^((prunecoeff - 1) / 2)
    dbhelement[2] <- sqrt((ncum[2] * dbhelement[2]^2 - nelement[1] * dbhelement[1]^2) / nelement[2])
    
    sellag[1] <- sellag[1] + adjageel[1] - CalcAge(dbhelement[1], A200, N, SI)
    sellag[2] <- sellag[2] + adjageel[2] - CalcAge(dbhelement[2], A200, N, SI)
    
    adjageel[1] <- Age - prlag[1] - totalthinlag - sellag[1]
    adjageel[2] <- Age - prlag[2] - totalthinlag - sellag[2]
    
    total_prlag[1] <- total_prlag[2] + pra * (PRHT[1]^prb - PRHT[2]^prb) * exp(-prc * crlth)
  } else {
    prevprht <- PRHT[1]
    PRHT[1] <- Pruning_history$lift_height[lift]
    crlth <- Meanht[1] - PRHT[1]
    total_prlag[1] <- total_prlag[1] + pra * (PRHT[1]^prb - prevprht^prb) * exp(-prc * crlth)
  }
  list(
    Nelements = Nelements,
    PRHT = PRHT,
    nelement = nelement,
    ncum = ncum,
    total_prlag = total_prlag,
    prlag = prlag,
    dbhelement = dbhelement,
    Meanht = Meanht,
    adjageel = adjageel,
    sellag = sellag
  )
}
CalcAge <- function(DBH, A200, stockn, SI) {
  # Calculate Age from DBH, A200 (DBH at 200 sph, age 30), Stocking, and SI using Bisection Method
  BisectionFn(0.001, 150, 15, 3, A200, SI, stockn, DBH)
}
thinning <- function(N, DBH, shist, Nelements, dbhelement, nelement, ncum, adjage, thin, totalthinlag, sellag, prlag, I300, Age, MTH) {
  prevN <- N
  prevdbh <- DBH
  
  kcoeff <- ifelse(Stocking_history$shist_thincoeff[shist] != 0, Stocking_history$shist_thincoeff[shist], thincoeff)
  
  thinN <- prevN - Stocking_history$shist_N2[shist]
  
  for (el in Nelements:1) {
    if (thinN + 0.0001 >= nelement[el]) {
      thinN <- thinN - nelement[el]
      Nelements <- Nelements - 1
    } else {
      prevNel <- nelement[el]
      prevNcum <- ncum[el]
      nelement[el] <- nelement[el] - thinN
      ncum[el] <- ncum[el] - thinN
      
      if (el != 1) {
        dbhelement[el] <- dbhelement[el] * (ncum[el]^((kcoeff + 1) / 2) - ncum[el - 1]^((kcoeff + 1) / 2)) * (prevNcum - ncum[el - 1]) / ((prevNcum^((kcoeff + 1) / 2) - ncum[el - 1]^((kcoeff + 1) / 2)) * (ncum[el] - ncum[el - 1]))
      } else {
        dbhelement[el] <- dbhelement[el] * (ncum[el] / prevNcum)^((kcoeff - 1) / 2)
      }
      
      Nelements <- el
      break
    }
  }
  
  N <- ncum[Nelements]
  dbhsqd <- 0
  
  for (el in 1:Nelements) {
    dbhsqd <- dbhsqd + nelement[el] * dbhelement[el]^2
  }
  
  DBH <- sqrt(dbhsqd / N)
  
  if (Stocking_history$shist_thinratio[shist] != 0 && prevdbh != 0) {
    current_thinratio <- DBH / prevdbh
    for (el in 1:Nelements) {
      dbhelement[el] <- dbhelement[el] * Stocking_history$shist_thinratio[shist] / current_thinratio
    }
    DBH <- DBH * Stocking_history$shist_thinratio[shist] / current_thinratio
  }
  
  #VolBA() removed function replaced with below
  BA <- CalcBAfromDBH(DBH, N)
  Vol <- CalcVol(MTH, BA, N)
  
  A200 <- CalcA200start(adjage, I300, SI)
  initiallag[thin] <- adjage - CalcAge(prevdbh, A200, N, SI)#Thin is zero it fails
  A200 <- CalcA200start(adjage - initiallag[thin], I300, SI)
  initiallag[thin] <- adjage - CalcAge(prevdbh, A200, N, SI)
  A200 <- CalcA200start(adjage - initiallag[thin], I300, SI)
  initiallag[thin] <- adjage - CalcAge(prevdbh, A200, N, SI)
  
  ThinLag[thin] <- initiallag[thin]
  agethin[thin] <- Age
  totalthinlag <- totalthinlag + initiallag[thin]
  
  for (el in 1:Nelements) {
    if (dbhelement[el] == 0) {
      sellag[el] <- 0
    } else {
      sellag[el] <- Age - prlag[el] - totalthinlag - CalcAge(dbhelement[el], A200, N, SI)
    }
    adjageel[el] <- Age - prlag[el] - totalthinlag - sellag[el]
  }
  
  adjage <- sum(adjageel[1:length(nelement)] * nelement) / N
  mnheight <- calcMeanht(MTH, N)
  list(
    N = N,
    DBH = DBH,
    Nelements = Nelements,
    dbhelement = dbhelement,
    nelement = nelement,
    ncum = ncum,
    adjage = adjage,
    totalthinlag = totalthinlag,
    sellag = sellag,
    BA = BA,
    Vol = Vol,
    mnheight = mnheight
  )
}
D300 <- function(T, TBH, D300_30, DBH_model, DBH_form, DBH_a, DBH_b, DBH_c) {
  return(Y(T - TBH, DBH_model, DBH_form, D300_30, 30 - TBH, DBH_a, DBH_b, DBH_c))
}
Y <- function(T, model, form, Y0, t0, A, B, C) {
  R0 <- 0
  if (model == "Richards") {
    if (form == "Anamorphic") {
      return(Y0 * ((1 - exp(-B * T)) / (1 - exp(-B * t0))) ^ C)
    } else if (form == "CA") {
      return(A * (1 - (1 - (Y0 / A) ^ (1 / C)) ^ (T / t0)) ^ C)
    } else if (form == "GADA") {
      R0 <- (log(Y0) - log((1 - exp(-B * t0)) ^ C)) / (1 + log((1 - exp(-B * t0)) ^ A))
      return(exp(R0) * (1 - exp(-B * T)) ^ (C + A * R0))
    } else {
      return(0)
    }
  } else if (model == "Korf") {
    if (form == "Anamorphic") {
      return(Y0 * exp(-B * T ^ (-C)) / exp(-B * t0 ^ (-C)))
    } else if (form == "CA") {
      return(A * (Y0 / A) ^ ((t0 / T) ^ C))
    } else if (form == "GADA") {
      R0 <- log(Y0) + sqrt((log(Y0)) ^ 2 + 4 * B * (t0) ^ (-C))
      return(exp(R0 / 2 - 2 * B / (R0 * (T ^ C))))
    } else {
      return(0)
    }
  } else if (model == "Hossfeld") {
    if (form == "Anamorphic") {
      return((T ^ C) / (B + T ^ C * (1 / Y0 - B / (t0 ^ C))))
    } else if (form == "CA") {
      return((T ^ C) / ((t0 ^ C) / Y0 + (T ^ C - t0 ^ C) / A))
    } else if (form == "GADA") {
      R0 <- (Y0 - A + sqrt((Y0 - A) ^ 2 + 4 * Y0 * B * (t0) ^ (-C))) / 2
      return((A + R0) / (1 + B * (T ^ (-C)) / R0))
    } else {
      return(0)
    }
  } else {
    return(0)
  }
}

# ---------------------------------------------------------------------------
# CalcOffsets — VBA Module 1 port
# Derives additive/multiplicative offsets for MTH and DBH from a measurement.
# Used when implementation == 2 (Offset mode for PSP data).
# ---------------------------------------------------------------------------
CalcOffsets <- function() {
  Inputparms()
  voltab()

  DBH_calibration_age <- data_300_index[7, 3]
  Stock300 <- data_300_index[8, 3]
  DBH300 <- data_300_index[9, 3]
  if (is.na(DBH300) || DBH300 == 0) {
    BA300 <- data_300_index[10, 3]
    if (!is.na(BA300) && BA300 != 0) {
      DBH300 <- CalcDBHfromBA(BA300, Stock300)
    } else {
      Vol300 <- data_300_index[11, 3]
      MTH300_local <- CalcMTH(SI, age300)
      DBH300 <- CalcDBHfromBA(calcBAfromVol(MTH300_local, Vol300, Stock300), Stock300)
    }
  }
  MTH_calibration_age <- data_300_index[14, 3]
  MTH300 <- data_300_index[15, 3]

  OUTPUT <- FALSE

  # Run Growth to DBH calibration age to get predicted DBH
  old_maxage <- maxage
  old_steps <- steps
  maxage <<- DBH_calibration_age
  steps <<- as.integer(maxage / steplength)
  gr <- Growth(OUTPUT, I300)
  DBHsqd_add_offset <<- DBH300^2 - gr$DBH_end^2
  DBHsqd_mult_offset <<- DBH300^2 / gr$DBH_end^2

  # Run Growth to MTH calibration age to get predicted MTH
  maxage <<- MTH_calibration_age
  steps <<- as.integer(maxage / steplength)
  gr <- Growth(OUTPUT, I300)
  MTH_add_offset <<- MTH300 - gr$MTH_end
  MTH_mult_offset <<- MTH300 / gr$MTH_end

  # Restore maxage/steps
  maxage <<- old_maxage
  steps <<- old_steps

  assign("DBH_calibration_age", DBH_calibration_age, envir = MODEL_ENV)
  assign("MTH_calibration_age", MTH_calibration_age, envir = MODEL_ENV)
  assign("DBHsqd_add_offset", DBHsqd_add_offset, envir = MODEL_ENV)
  assign("DBHsqd_mult_offset", DBHsqd_mult_offset, envir = MODEL_ENV)
  assign("MTH_add_offset", MTH_add_offset, envir = MODEL_ENV)
  assign("MTH_mult_offset", MTH_mult_offset, envir = MODEL_ENV)
}

OutputGrowth <- function() {
  # Define a helper function to check inputs (equivalent to VBA checks)
  check_input <- function() {
  #  if (!checkinput_site()) return(FALSE)
  #  if (!checkinput_SI()) return(FALSE)
  #  if (!checkinput_htfn()) return(FALSE)
  ##  if (!checkinput_initialstock()) return(FALSE)
  #  if (!checkinput_stocking()) return(FALSE)
  #  if (!checkinput_prune()) return(FALSE)
  #  if (!checkinput_fellage()) return(FALSE)
  #  if (!checkinput_steplth()) return(FALSE)
  #  if (!checkinput_volfn()) return(FALSE)
  #  if (!checkinput_mortfn()) return(FALSE)
  #  return(TRUE)
  }
  
  # Exit if any input check fails
 # if (!check_input()) return(NULL)
  
  # Set OUTPUT to TRUE
  OUTPUT <- TRUE
  
  # Read the '300 Index' sheet from the Excel file
  
  # Clear contents in the equivalent range in R (rows 5-150 and columns G-BR)
  # This would typically involve modifying the dataframe, but in Excel, we clear cells in the range G5:BR150
  outputrange <- data_300_index[]

  # Write the cleared output range back to the Excel file
  #print(outputrange)
  
  # Call other functions 
  Inputparms()  # Call Inputparms function
 # voltab()      # Call voltab function
  
  # If implementation equals 2, call CalcOffsets and Inputparms again
  if (implementation == 2) {
    CalcOffsets()  # Call CalcOffsets function
    Inputparms()   # Call Inputparms again
    OUTPUT <- TRUE
  }
  
  # Call remaining functions in sequence
  gr <- Growth(OUTPUT, I300)  # Call Growth function
  earlyield()     # Call earlyield function
  mortvol()       # Call mortvol function

  # Compute wood density for each row and add WoodDensity column
  if (!is.null(gr$growth_df) && nrow(gr$growth_df) > 0) {
    gr$growth_df <- density(gr$growth_df)
  }
  gr
}



# ---------------------------------------------------------------------------
# load_input_sheets(workbook) — loads Inputs, 300 Index, and parameters
# sheets from an Excel workbook into MODEL_ENV.
# Only called explicitly (e.g. by single-site workflows), never at source time.
# For batch PSP workflow, use load_parameters_from_csv() and
# build_synthetic_matrices() instead.
# ---------------------------------------------------------------------------
load_input_sheets <- function(workbook) {
  Meanht <- numeric(11)
  adjageel <- numeric(11)
  initiallag <- numeric(9)
  ThinLag <- numeric(9)
  agethin <- numeric(9)
  DBH <- 0

  input_data <- read_sheet(workbook, "Inputs", col_names = FALSE)
  Species <- input_data[2, 4]
  assign("Species", Species, envir = MODEL_ENV)
  input_data <- as.data.frame(lapply(input_data, as.numeric))
  assign("input_data", input_data, envir = MODEL_ENV)

  data_300_index <- read_sheet(workbook, "300 Index", col_names = FALSE)
  data_300_indexX <- as.data.frame(data_300_index[51:72, 4:6])
  data_300_index <- as.data.frame(lapply(data_300_index, as.numeric))
  assign("data_300_index", data_300_index, envir = MODEL_ENV)
  assign("data_300_indexX", data_300_indexX, envir = MODEL_ENV)

  parameters <- read_sheet(workbook, "parameters", col_names = TRUE)
  assign("parameters", parameters, envir = MODEL_ENV)
  if ("Variable" %in% names(parameters)) {
    variable_names <- parameters$Variable
  } else {
    variable_names <- parameters$name
  }
  if ("Coefficients" %in% names(parameters)) {
    variable_values <- parameters$Coefficients
  } else {
    variable_values <- parameters$value
  }
  for (i in seq_along(variable_names)) {
    if (is.na(variable_names[i]) || as.character(variable_names[i]) == "") next
    val <- type.convert(as.character(variable_values[i]), as.is = TRUE)
    assign(variable_names[i], val, envir = MODEL_ENV)
  }
}


run_model <- function() {
  # Run Growth model
  Error_flag <- FALSE
  Cali <- FALSE
  
  # Activate worksheet "Inputs" and input parameters
  
  Inputparms()  
  data_300_index<- Input_parameters()
  # Check for errors
  ##  if (Check_errors) Error_checks_1() #not yet defined
  ##  if (Error_flag) return()
  
  # Input unscaled tree list or generate it using a Weibull distribution
  Treelist <- Input_tree_list()
  
  if (mode == 3) {
    #estimate from starting tree list
    if (Check_errors) Error_checks_4()
    if (Error_flag) return()
    Treelist_results <- Process_tree_list(Treelist, Plot_area, Age, nstems)# Derive stand metrics from tree list
    MTDia<-Treelist_results$MTDia
    stocking<-Treelist_results$stocking
    petA <-Treelist_results$petA
    petB<-Treelist_results$petB
    qDBH<-Treelist_results$qDBH
    BA<-Treelist_results$BA
    populate_300index_inputs_from_tree(Treelist_results)
  }
  
  if (mode == 2 || mode == 3) {
    if (Check_errors) Error_checks_3()
    if (Error_flag) return()
    siteIndex()
    Calc300Index()
    
    if (Species == "Radiata pine") {
      Calibrate_radiata()
    } else if (Species == "Douglas-fir") {
      Calibrate_dfir()
    } else {
      Calibrate()
    }
  }
  
  # Growth prediction from summary pathway (optionally enriched by tree metrics)
  growth_result <- OutputGrowth()
  if (!is.null(growth_result$growth_df) && nrow(growth_result$growth_df) > 0) {
    write.csv(growth_result$growth_df, "growth_check_output.csv", row.names = FALSE)
  }
  
  if (Check_errors) Error_checks_2()
  if (Error_flag) return()
  
  # Estimate breast height age
  Input_parameters()  # Input stocking history and other parameters
  
  if (Species == "Coast redwood" && T2 != 0) {
    MTH_b <- MTHmodel_b(30, H30, T2, H2)
  }
  
  TBH <- AgeBH(30, H30, MTH_model, MTH_form, MTH_a, MTH_b, MTH_c)
  
  stemno <- 1
  logno <- 1
  D300_30_est <- D300_30_from_I300_SI(I300, H30, VOL_type, VOL_u, VOL_v, VOL_w, VOL_z)
  
  # Generate yield table based on species
  if (Species == "Radiata pine") {
    Yield_Table_radiata()
  } else if (Species == "Douglas-fir") {
    Yield_Table_dfir()
  } else {
    Yield_Table()
  }
  
  if (Check_errors) Error_checks_5()
  
  # Set up C_Change inputs and estimate mortality volume
  Initialise_C_Change()
  Mortality_Volume()
  
  # Run C_Change for non-redwood species or use Kizha Han models for redwood
  if (!Minimal_run) {
    if (Species != "Coast redwood") {
      Run_C_Change()
    } else {
      Kizha_Han()
    }
  }
  
  # Output table
  Output_table()
  
  # Check diameter distribution
  if (DiaDist == 2) {
    if (Check_errors) Error_checks_4b()
    if (Error_flag) return()
  }
  
  # Export key results to MODEL_ENV so the batch loop can retrieve them
  if (exists("yield_table", inherits = FALSE))
    assign("yield_table", yield_table, envir = MODEL_ENV)
  if (exists("carbon_results", inherits = FALSE))
    assign("carbon_results", carbon_results, envir = MODEL_ENV)
  if (exists("cchange_result", inherits = FALSE))
    assign("cchange_result", cchange_result, envir = MODEL_ENV)
  if (exists("felled_stems_df", inherits = FALSE))
    assign("felled_stems_df", felled_stems_df, envir = MODEL_ENV)
  if (exists("logs_df", inherits = FALSE))
    assign("logs_df", logs_df, envir = MODEL_ENV)
  if (exists("harvest_sum", inherits = FALSE))
    assign("harvest_sum", harvest_sum, envir = MODEL_ENV)
  
  invisible(growth_result)
}




# ---------------------------
# Tree-level input entrypoints
# ---------------------------

run_treelevel_input <- function(workbook,
                                output_path = "plot_summary_from_tree.csv",
                                tree_sheet = "Starting tree list") {
  Input_parameters()
  Inputparms()

  starting_tree_list <- read_sheet(workbook, tree_sheet, col_names = TRUE, skip = 5)
  tree_header <- read_sheet(workbook, tree_sheet, col_names = FALSE)
  plot_area <- as.numeric(tree_header[3, 2])
  age <- as.numeric(tree_header[4, 2])
  nstems <- nrow(starting_tree_list)

  if (Error_checks_4(starting_tree_list, plot_area, age)) {
    stop("Tree list failed validation checks.")
  }

  tree_metrics <- Process_tree_list(starting_tree_list, plot_area, age, nstems)
  populate_300index_inputs_from_tree(tree_metrics)
  siteIndex()
  Calc300Index()

  summary_df <- data.frame(
    PlotID = if ("PlotSampleID" %in% names(starting_tree_list)) as.character(starting_tree_list$PlotSampleID[[1]]) else "single_plot",
    Age = as.numeric(tree_metrics$Age),
    Stocking = as.numeric(tree_metrics$stocking),
    qDBH = as.numeric(tree_metrics$qDBH),
    BA = as.numeric(tree_metrics$BA),
    MTH = as.numeric(tree_metrics$MTH),
    SI = as.numeric(data_300_index[4, 3]),
    Index300 = as.numeric(data_300_index[3, 3])
  )

  write.csv(summary_df, file = output_path, row.names = FALSE)

  message("Wrote ", output_path)
  invisible(summary_df)
}

# Batch-ready helper: expects one row per stem and groups by plot id.
run_treelevel_input_batch <- function(tree_data,
                                      plot_id_col,
                                      plot_area_col,
                                      age_col,
                                      dbh_col,
                                      height_col,
                                      output_path = "plot_summary_from_tree_batch.csv") {
  Input_parameters()
  Inputparms()

  ids <- unique(tree_data[[plot_id_col]])
  out <- vector("list", length(ids))

  for (i in seq_along(ids)) {
    pid <- ids[[i]]
    pdat <- tree_data[tree_data[[plot_id_col]] == pid, , drop = FALSE]
    treelist <- data.frame(
      StemID = seq_len(nrow(pdat)),
      DBH = as.numeric(pdat[[dbh_col]]),
      Height = as.numeric(pdat[[height_col]])
    )
    plot_area <- as.numeric(pdat[[plot_area_col]][1])
    age <- as.numeric(pdat[[age_col]][1])
    nstems <- nrow(treelist)

    if (Error_checks_4(treelist, plot_area, age)) next

    tm <- Process_tree_list(treelist, plot_area, age, nstems)
    populate_300index_inputs_from_tree(tm)
    siteIndex()
    Calc300Index()

    out[[i]] <- data.frame(
      PlotID = as.character(pid),
      Age = as.numeric(tm$Age),
      Stocking = as.numeric(tm$stocking),
      qDBH = as.numeric(tm$qDBH),
      BA = as.numeric(tm$BA),
      MTH = as.numeric(tm$MTH),
      SI = as.numeric(data_300_index[4, 3]),
      Index300 = as.numeric(data_300_index[3, 3])
    )
  }

  batch_df <- do.call(rbind, Filter(Negate(is.null), out))
  write.csv(batch_df, file = output_path, row.names = FALSE)
  message("Wrote ", output_path)
  invisible(batch_df)
}
