# Set working directory portably (works outside RStudio)
if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
} else {
  this_dir <- getSrcDirectory(function(x) x)
  if (nzchar(this_dir)) {
    setwd(this_dir)
  } else {
    # fallback: assume scripts live in the current working directory
    message("Note: could not detect script directory; using current working directory.")
  }
}

source("Tree_To_PlotSummary.R")
source("TreeLevel_Input.R")
source("Growth_From_PlotSummary.R")
source("03_Yield_From_Growth.R")
source("04_Carbon_From_Yield.R")
source("05_Model_Report.R")
source("DouglasFir_500Index.R")
source("input_validation.R")
source("MultiSpecies_Growth.R")

run_full_chain <- function(input_workbook = NULL,
                           use_tree_level = TRUE,
                           output_dir = ".",
                           species = "Radiata pine",
                           validate_inputs = TRUE,
                           use_cchange = TRUE,
                           soil_c = 5.57, soil_n = 0.296,
                           soil_organic_p = 333,
                           MATEMP = 12) {
  if (!is.null(input_workbook)) {
    assign("INPUT_WORKBOOK", input_workbook, envir = MODEL_ENV)
  }
  old_wd <- getwd()
  on.exit(setwd(old_wd))
  if (output_dir != ".") {
    if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
    setwd(output_dir)
  }

  # Douglas-fir pathway
  if (tolower(species) %in% c("douglas-fir", "douglas fir", "pseudotsuga menziesii")) {
    message("Using Douglas-fir 500 Index pathway")
    result <- dfir_yield(
      SI = if (exists("SI", envir = MODEL_ENV)) get("SI", envir = MODEL_ENV) else NULL,
      I500 = if (exists("I300", envir = MODEL_ENV)) get("I300", envir = MODEL_ENV) else NULL,
      latitude = if (exists("latitude", envir = MODEL_ENV)) get("latitude", envir = MODEL_ENV) else -42
    )
    if (!is.null(result)) {
      writexl::write_xlsx(list(annual_yield = result$annual), path = "yield_from_growth.xlsx")
      run_carbon_from_yield(use_cchange = use_cchange,
                            soil_c = soil_c, soil_n = soil_n,
                            soil_organic_p = soil_organic_p,
                            MATEMP = MATEMP)
      run_model_report()
    }
    message("Douglas-fir model chain complete.")
    return(invisible(result))
  }

  # Standard radiata pine / multi-species pathway
  if (validate_inputs && !is.null(input_workbook)) {
    tryCatch({
      inp <- as.data.frame(readxl::read_excel(
        if (exists("INPUT_WORKBOOK", envir = MODEL_ENV)) get("INPUT_WORKBOOK", envir = MODEL_ENV) else input_workbook,
        sheet = 1, col_names = FALSE, .name_repair = "minimal"))
      v <- run_all_validations(data_300_index = inp, stop_on_error = FALSE)
      if (!v$valid) message("Input validation issues found; continuing anyway.")
    }, error = function(e) {
      message(paste("Skipping validation:", e$message))
    })
  }

  # Multi-species pathway: use run_model() which routes to the correct
  # species-specific yield table, calibration, and carbon model
  assign("Species", species, envir = MODEL_ENV)
  message(sprintf("Running multi-species model for: %s", species))
  tryCatch({
    run_model()
    if (exists("yield_table", envir = MODEL_ENV)) {
      writexl::write_xlsx(list(yield_table = get("yield_table", envir = MODEL_ENV)), path = "yield_from_growth.xlsx")
    }
    if (exists("carbon_results", envir = MODEL_ENV)) {
      writexl::write_xlsx(list(carbon = get("carbon_results", envir = MODEL_ENV)), path = "carbon_output.xlsx")
    }
    if (exists("felled_stems_df", envir = MODEL_ENV)) {
      writexl::write_xlsx(list(felled_stems = get("felled_stems_df", envir = MODEL_ENV)), path = "felled_stems.xlsx")
    }
    if (exists("logs_df", envir = MODEL_ENV)) {
      writexl::write_xlsx(list(logs = get("logs_df", envir = MODEL_ENV)), path = "harvest_logs.xlsx")
    }
  }, error = function(e) {
    message(sprintf("Multi-species run_model() failed: %s", e$message))
    message("Falling back to standard radiata pathway...")
    if (use_tree_level) {
      run_treelevel_input(output_path = "plot_summary_from_tree.xlsx")
    }
    run_growth_from_plot_summary()
    run_yield_from_growth()
  })

  run_carbon_from_yield(use_cchange = use_cchange,
                        soil_c = soil_c, soil_n = soil_n,
                        soil_organic_p = soil_organic_p,
                        MATEMP = MATEMP)
  run_model_report()
  message("Full model chain complete.")
}

# Batch runner: run the full chain for multiple input workbooks
run_batch <- function(workbook_paths,
                      use_tree_level = TRUE,
                      output_root = "batch_output",
                      species = "Radiata pine") {
  results <- vector("list", length(workbook_paths))
  for (i in seq_along(workbook_paths)) {
    wb <- workbook_paths[i]
    site_name <- tools::file_path_sans_ext(basename(wb))
    out_dir <- file.path(output_root, site_name)
    message(sprintf("[%d/%d] Running site: %s (%s)", i, length(workbook_paths), site_name, species))
    tryCatch({
      reset_model_env()
      run_full_chain(
        input_workbook = normalizePath(wb, mustWork = TRUE),
        use_tree_level = use_tree_level,
        output_dir = out_dir,
        species = species
      )
      results[[i]] <- list(site = site_name, status = "success", error = NA_character_)
    }, error = function(e) {
      message(sprintf("  ERROR in %s: %s", site_name, conditionMessage(e)))
      results[[i]] <<- list(site = site_name, status = "error", error = conditionMessage(e))
    })
  }
  summary_df <- do.call(rbind, lapply(results, as.data.frame, stringsAsFactors = FALSE))
  message(sprintf("\nBatch complete: %d succeeded, %d failed",
                  sum(summary_df$status == "success"),
                  sum(summary_df$status == "error")))
  summary_df
}
