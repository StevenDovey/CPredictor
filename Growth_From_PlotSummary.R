library(readxl)
library(writexl)

if (requireNamespace("rstudioapi", quietly = TRUE) &&
    rstudioapi::isAvailable()) {
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
}

source("TreeLevel_Input.R")

pick_summary_file <- function() {
  candidates <- c("plot_summary_input.xlsx", "plot_summary_from_tree.xlsx")
  hit <- candidates[file.exists(candidates)][1]
  if (is.na(hit) || is.null(hit)) {
    stop("No plot summary file found. Provide plot_summary_input.xlsx or run TreeLevel_Input.R first.")
  }
  hit
}

run_growth_from_plot_summary <- function(
  summary_file = NULL,
  output_file = "growth_from_plot_summary.xlsx"
) {
  Input_parameters()
  Inputparms()

  if (is.null(summary_file)) summary_file <- pick_summary_file()
  plot_summary <- as.data.frame(read_excel(summary_file, sheet = "plot_summary"))
  if (nrow(plot_summary) < 1) stop("plot_summary sheet is empty.")
  r <- plot_summary[1, ]

  if (!is.na(r$Age)) data_300_index[14, 3] <- as.numeric(r$Age)
  if (!is.na(r$MTH)) data_300_index[15, 3] <- as.numeric(r$MTH)
  if (!is.na(r$Stocking)) data_300_index[8, 3] <- as.numeric(r$Stocking)
  if (!is.na(r$qDBH)) data_300_index[9, 3] <- as.numeric(r$qDBH)
  if (!is.na(r$BA)) data_300_index[10, 3] <- as.numeric(r$BA)
  if ("SI" %in% names(r) && !is.na(r$SI)) data_300_index[4, 3] <- as.numeric(r$SI)
  if ("Index300" %in% names(r) && !is.na(r$Index300)) data_300_index[3, 3] <- as.numeric(r$Index300)

  if (is.na(data_300_index[4, 3]) || data_300_index[4, 3] == 0) siteIndex()
  if (is.na(data_300_index[3, 3]) || data_300_index[3, 3] == 0) Calc300Index()

  SI <<- as.numeric(data_300_index[4, 3])
  I300 <<- as.numeric(data_300_index[3, 3])

  gr <- OutputGrowth()
  growth_df <- gr$growth_df
  if (is.null(growth_df) || nrow(growth_df) == 0) {
    if (file.exists("output.csv")) {
      growth_df <- read.csv("output.csv")
    } else {
      stop("Growth output not produced.")
    }
  }

  write_xlsx(
    list(
      growth = growth_df,
      summary_used = plot_summary
    ),
    path = output_file
  )

  message("Wrote ", output_file)
  growth_df
}
