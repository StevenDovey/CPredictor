setwd("C:/R/CPredictor/")

source("Run_Model_Chain.R")
source("infill_plots.R")

# =============================================================================
# Configuration
# =============================================================================
input_dir  <- "csv_inputs"
output_dir <- "batch_output"

# Set to TRUE to run the infill step (computes plot metrics: SI, drift stats).
# Set to FALSE to skip infill and use plots.csv as-is in the batch.
run_infill <- TRUE

# Set to TRUE to run the main yield/carbon batch.
# Set to FALSE to do infill only (useful when you only want plot metrics).
run_main_batch <- TRUE

# =============================================================================
# Pipeline
# =============================================================================
params_csv <- file.path(input_dir, "parameters.csv")
plots_csv  <- file.path(input_dir, "plots.csv")
psp_csv    <- file.path(input_dir, "psp_summary.csv")

pars <- load_parameters_from_csv(params_csv)

if (run_infill) {
  message("=== Step 1: Infill plot metrics ===")
  infill_plots(plots_csv, psp_csv, output_dir, pars)
  batch_input <- file.path(output_dir, "plots_infilled.csv")
} else {
  batch_input <- NULL  # use plots.csv directly
}

if (run_main_batch) {
  message("\n=== Step 2: Run yield/carbon batch ===")
  run_batch_psp(input_source   = input_dir,
                plots_override = batch_input,
                output_dir     = output_dir)
}


source("compare_model_outputs.R")