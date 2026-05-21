setwd(dirname(rstudioapi::getSourceEditorContext()$path))

source("Run_Model_Chain.R")

# Point to the directory containing the 3 input CSVs:
#   c_change_control.csv, psp_summary.csv, plots.csv
input_dir <- "examples/csv_inputs/"

# Output directory for results
output_dir <- "batch_output"

run_batch_psp(input_source = input_dir, output_dir = output_dir)
