library(readxl)
library(writexl)

if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}

run_carbon_from_yield <- function(
  yield_file = "yield_from_growth.xlsx",
  output_file = "carbon_from_yield.xlsx",
  wood_density_t_per_m3 = 0.42,
  carbon_fraction = 0.50
) {
  yld <- as.data.frame(read_excel(yield_file, sheet = "annual_yield"))
  if (nrow(yld) == 0) stop("annual_yield sheet is empty.")

  req <- c("Age", "StemVol_m3_ha")
  miss <- setdiff(req, names(yld))
  if (length(miss)) stop(paste("Missing yield columns:", paste(miss, collapse = ", ")))

  carbon <- yld[, c("Age", "StemVol_m3_ha")]
  carbon$StemBiomass_t_ha <- carbon$StemVol_m3_ha * wood_density_t_per_m3
  carbon$StemCarbon_tC_ha <- carbon$StemBiomass_t_ha * carbon_fraction
  carbon$StemCarbonIncrement_tC_ha <- c(NA_real_, diff(carbon$StemCarbon_tC_ha))

  carbon_summary <- data.frame(
    FinalStemCarbon_tC_ha = tail(carbon$StemCarbon_tC_ha, 1),
    PeakStemCarbon_tC_ha = max(carbon$StemCarbon_tC_ha, na.rm = TRUE),
    MeanAnnualStemCarbonIncrement_tC_ha = mean(carbon$StemCarbonIncrement_tC_ha, na.rm = TRUE),
    wood_density_t_per_m3 = wood_density_t_per_m3,
    carbon_fraction = carbon_fraction
  )

  write_xlsx(
    list(
      annual_carbon = carbon,
      carbon_summary = carbon_summary
    ),
    path = output_file
  )
  message("Wrote ", output_file)
  list(annual_carbon = carbon, carbon_summary = carbon_summary)
}
