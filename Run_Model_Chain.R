setwd(dirname(rstudioapi::getSourceEditorContext()$path))

source("TreeLevel_Input.R")
source("Growth_From_PlotSummary.R")
source("03_Yield_From_Growth.R")
source("04_Carbon_From_Yield.R")
source("05_Model_Report.R")

run_full_chain <- function(use_tree_level = TRUE) {
  if (use_tree_level) {
    run_treelevel_input(output_path = "plot_summary_from_tree.xlsx")
  }

  run_growth_from_plot_summary()
  run_yield_from_growth()
  run_carbon_from_yield()
  run_model_report()
  message("Full model chain complete.")
}
