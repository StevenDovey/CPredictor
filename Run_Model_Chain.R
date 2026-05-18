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

source("TreeLevel_Input.R")
source("Growth_From_PlotSummary.R")
source("03_Yield_From_Growth.R")
source("04_Carbon_From_Yield.R")
source("05_Model_Report.R")

run_full_chain <- function(input_workbook = NULL,
                           use_tree_level = TRUE,
                           output_dir = ".") {
  if (!is.null(input_workbook)) {
    assign("INPUT_WORKBOOK", input_workbook, envir = .GlobalEnv)
  }
  old_wd <- getwd()
  on.exit(setwd(old_wd))
  if (output_dir != ".") {
    if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)
    setwd(output_dir)
  }

  if (use_tree_level) {
    run_treelevel_input(output_path = "plot_summary_from_tree.xlsx")
  }

  run_growth_from_plot_summary()
  run_yield_from_growth()
  run_carbon_from_yield()
  run_model_report()
  message("Full model chain complete.")
}

# Batch runner: run the full chain for multiple input workbooks
run_batch <- function(workbook_paths,
                      use_tree_level = TRUE,
                      output_root = "batch_output") {
  results <- vector("list", length(workbook_paths))
  for (i in seq_along(workbook_paths)) {
    wb <- workbook_paths[i]
    site_name <- tools::file_path_sans_ext(basename(wb))
    out_dir <- file.path(output_root, site_name)
    message(sprintf("[%d/%d] Running site: %s", i, length(workbook_paths), site_name))
    tryCatch({
      run_full_chain(
        input_workbook = normalizePath(wb, mustWork = TRUE),
        use_tree_level = use_tree_level,
        output_dir = out_dir
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
