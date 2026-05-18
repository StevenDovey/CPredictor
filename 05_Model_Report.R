library(readxl)
library(writexl)
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

run_model_report <- function(
  summary_file = "plot_summary_from_tree.xlsx",
  growth_file = "growth_from_plot_summary.xlsx",
  yield_file = "yield_from_growth.xlsx",
  carbon_file = "carbon_from_yield.xlsx",
  output_file = "model_chain_outputs.xlsx"
) {
  plot_summary <- as.data.frame(read_excel(summary_file, sheet = "plot_summary"))
  growth <- as.data.frame(read_excel(growth_file, sheet = "growth"))
  annual_yield <- as.data.frame(read_excel(yield_file, sheet = "annual_yield"))
  annual_carbon <- as.data.frame(read_excel(carbon_file, sheet = "annual_carbon"))

  write_xlsx(
    list(
      plot_summary = plot_summary,
      growth = growth,
      annual_yield = annual_yield,
      annual_carbon = annual_carbon
    ),
    path = output_file
  )
  message("Wrote ", output_file)
  output_file
}
