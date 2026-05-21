setwd(dirname(rstudioapi::getSourceEditorContext()$path))

source("Tree_To_PlotSummary.R")

# Convert individual tree data to plot-level summaries (SI, 300 Index, BA, etc.)
# This is an optional data preparation step — only needed if you have tree-level
# data and need to build the plot input files for run_batch.R
run_tree_to_plot_summary(
  tree_file       = "NZFM Fert Trial Ind Tree Data as at Nov25.xlsx",
  tree_sheet      = 1L,
  input_file      = "input.xlsx",
  plot_site_file  = "Plot_site_data.xlsx",
  output_file     = "Plot_Summary.xlsx"
)
