library(readxl)
library(writexl)

if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}

if (!exists("read_data", mode = "function")) source("io_utils.R")
source("CChange_model.R")

run_carbon_from_yield <- function(
  yield_file = "yield_from_growth.xlsx",
  output_file = "carbon_from_yield.xlsx",
  wood_density_t_per_m3 = 0.42,
  carbon_fraction = 0.50,
  use_cchange = TRUE,
  disturbance_schedule = NULL,
  soil_c = 5.57, soil_n = 0.296,
  soil_organic_p = 333,
  MATEMP = 12
) {
  yld <- as.data.frame(read_data(yield_file, sheet = "annual_yield"))
  if (nrow(yld) == 0) stop("annual_yield sheet is empty.")

  req <- c("Age", "StemVol_m3_ha")
  miss <- setdiff(req, names(yld))
  if (length(miss)) stop(paste("Missing yield columns:", paste(miss, collapse = ", ")))

  # Simple stem-only carbon accounting (always computed as baseline)
  carbon <- yld[, c("Age", "StemVol_m3_ha")]
  carbon$StemBiomass_t_ha <- carbon$StemVol_m3_ha * wood_density_t_per_m3
  carbon$StemCarbon_tC_ha <- carbon$StemBiomass_t_ha * carbon_fraction
  carbon$StemCarbonIncrement_tC_ha <- c(NA_real_, diff(carbon$StemCarbon_tC_ha))

  # Full CChange compartment model
  cchange_result <- NULL
  if (use_cchange) {
    growth_table <- data.frame(
      Age = yld$Age,
      SPHA = if ("Stocking_stems_ha" %in% names(yld)) yld$Stocking_stems_ha else
             if ("SPHA" %in% names(yld)) yld$SPHA else rep(1000, nrow(yld)),
      MTH = if ("MTH_m" %in% names(yld)) yld$MTH_m else
            if ("MTH" %in% names(yld)) yld$MTH else rep(0, nrow(yld)),
      Vol = yld$StemVol_m3_ha,
      GrossVol = if ("GrossVol_m3_ha" %in% names(yld)) yld$GrossVol_m3_ha else yld$StemVol_m3_ha,
      BA = if ("BA_m2_ha" %in% names(yld)) yld$BA_m2_ha else
           if ("BA" %in% names(yld)) yld$BA else rep(0, nrow(yld)),
      WholeStemDens = rep(wood_density_t_per_m3, nrow(yld)),
      RingDens = rep(0, nrow(yld)),
      stringsAsFactors = FALSE
    )

    tryCatch({
      cchange_result <- run_cchange(
        growth_table = growth_table,
        disturbance_schedule = disturbance_schedule,
        soil_c = soil_c, soil_n = soil_n,
        soil_organic_p = soil_organic_p,
        MATEMP = MATEMP
      )
      message("CChange compartment model completed successfully.")
    }, error = function(e) {
      warning(paste("CChange model failed, using simple carbon only:", e$message))
    })
  }

  carbon_summary <- data.frame(
    FinalStemCarbon_tC_ha = tail(carbon$StemCarbon_tC_ha, 1),
    PeakStemCarbon_tC_ha = max(carbon$StemCarbon_tC_ha, na.rm = TRUE),
    MeanAnnualStemCarbonIncrement_tC_ha = mean(carbon$StemCarbonIncrement_tC_ha, na.rm = TRUE),
    wood_density_t_per_m3 = wood_density_t_per_m3,
    carbon_fraction = carbon_fraction
  )

  sheets <- list(
    annual_carbon = carbon,
    carbon_summary = carbon_summary
  )

  if (!is.null(cchange_result)) {
    sheets$cchange_annual <- cchange_result$annual_carbon
    sheets$cchange_summary <- cchange_result$carbon_summary
  }

  write_xlsx(sheets, path = output_file)
  message("Wrote ", output_file)

  list(
    annual_carbon = carbon,
    carbon_summary = carbon_summary,
    cchange = cchange_result
  )
}
