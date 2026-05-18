library(readxl)
library(writexl)

if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}

run_yield_from_growth <- function(
  growth_file = "growth_from_plot_summary.xlsx",
  output_file = "yield_from_growth.xlsx"
) {
  growth <- as.data.frame(read_excel(growth_file, sheet = "growth"))
  if (nrow(growth) == 0) stop("Growth sheet is empty.")

  names(growth) <- gsub("\\.", "_", names(growth))
  required <- c("Age", "N", "DBH", "BA", "Vol")
  missing <- setdiff(required, names(growth))
  if (length(missing)) {
    stop(paste("Missing expected growth columns:", paste(missing, collapse = ", ")))
  }

  annual_yield <- growth[, c("Age", "N", "DBH", "BA", "Vol")]
  names(annual_yield) <- c("Age", "Stocking", "DBH_cm", "BA_m2_ha", "StemVol_m3_ha")
  annual_yield$VolIncrement_m3_ha <- c(NA_real_, diff(annual_yield$StemVol_m3_ha))

  rotation_summary <- data.frame(
    StartAge = min(annual_yield$Age, na.rm = TRUE),
    EndAge = max(annual_yield$Age, na.rm = TRUE),
    FinalStocking = tail(annual_yield$Stocking, 1),
    FinalDBH_cm = tail(annual_yield$DBH_cm, 1),
    FinalBA_m2_ha = tail(annual_yield$BA_m2_ha, 1),
    FinalStemVol_m3_ha = tail(annual_yield$StemVol_m3_ha, 1),
    PeakStemVol_m3_ha = max(annual_yield$StemVol_m3_ha, na.rm = TRUE)
  )

  write_xlsx(
    list(
      annual_yield = annual_yield,
      rotation_summary = rotation_summary
    ),
    path = output_file
  )
  message("Wrote ", output_file)
  list(annual_yield = annual_yield, rotation_summary = rotation_summary)
}
