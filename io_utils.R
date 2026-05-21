# io_utils.R — CSV / Excel dispatcher for CPredictor inputs
#
# All model scripts use read_data() instead of read_excel() directly.
# This lets users supply either .csv or .xlsx files as input.
#
# For multi-sheet workbooks (e.g. input.xlsx with sheets: parameters, VolTab,
# 300 Index, Inputs, Starting tree list), the user can provide either:
#   1. A single .xlsx file (existing behaviour)
#   2. A directory of .csv files named after the sheets:
#        input_dir/parameters.csv
#        input_dir/voltab.csv
#        input_dir/300_index.csv
#        input_dir/inputs.csv
#        input_dir/starting_tree_list.csv

library(readxl)
library(writexl)

is_csv <- function(path) grepl("\\.(csv|tsv)$", tolower(path))
is_xlsx <- function(path) grepl("\\.(xlsx|xls|xlsm)$", tolower(path))

# Read a single-table file (CSV or Excel).
# For Excel files, `sheet` selects the sheet (name or number).
# For CSV files, `sheet` is ignored.
# Additional arguments (...) are passed to read_excel() or read.csv().
read_data <- function(path, sheet = 1L, col_names = TRUE, skip = 0, ...) {
  if (!file.exists(path)) stop("File not found: ", path)
  if (is_csv(path)) {
    df <- read.csv(path, header = col_names, skip = skip,
                   stringsAsFactors = FALSE, check.names = FALSE, ...)
    return(as.data.frame(df))
  }
  as.data.frame(read_excel(path, sheet = sheet, col_names = col_names,
                           skip = skip, .name_repair = "minimal", ...))
}

# Normalise a sheet name to a filename stem (lowercase, replace spaces with _).
sheet_to_csv_name <- function(sheet_name) {
  s <- tolower(trimws(sheet_name))
  s <- gsub("[^a-z0-9]+", "_", s)
  s <- gsub("^_|_$", "", s)
  paste0(s, ".csv")
}

# Read one "sheet" from either a multi-sheet .xlsx workbook or a directory of CSVs.
# `input_source` can be:
#   - a path to an .xlsx/.xlsm file  →  reads the named sheet
#   - a path to a directory           →  reads the matching .csv file
read_sheet <- function(input_source, sheet_name, col_names = TRUE, skip = 0, ...) {
  if (dir.exists(input_source)) {
    csv_name <- sheet_to_csv_name(sheet_name)
    csv_path <- file.path(input_source, csv_name)
    if (!file.exists(csv_path)) {
      stop("Expected CSV file '", csv_name, "' in directory: ", input_source)
    }
    return(read_data(csv_path, col_names = col_names, skip = skip, ...))
  }
  if (!file.exists(input_source)) stop("Input not found: ", input_source)
  read_data(input_source, sheet = sheet_name, col_names = col_names, skip = skip, ...)
}

# Read a specific cell range from a multi-sheet workbook or CSV directory.
# For Excel: uses the `range` argument.
# For CSV directory: reads the full CSV, then subsets to the specified row/col.
read_sheet_range <- function(input_source, sheet_name, row, col, skip = 0) {
  if (dir.exists(input_source)) {
    df <- read_sheet(input_source, sheet_name, col_names = FALSE, skip = skip)
    if (row > nrow(df) || col > ncol(df)) return(NA)
    return(df[row, col])
  }
  if (!file.exists(input_source)) stop("Input not found: ", input_source)
  # For Excel, read the full sheet and index
  df <- as.data.frame(read_excel(input_source, sheet = sheet_name,
                                 col_names = FALSE, skip = skip,
                                 .name_repair = "minimal"))
  if (row > nrow(df) || col > ncol(df)) return(NA)
  df[row, col]
}

# Write output data — supports both CSV and Excel based on output_file extension.
write_output <- function(data, output_file) {
  if (is_csv(output_file)) {
    if (is.data.frame(data)) {
      write.csv(data, file = output_file, row.names = FALSE)
    } else if (is.list(data)) {
      # For multi-sheet data written to CSV, write the first (or named) element
      # and additional elements as separate files
      base <- tools::file_path_sans_ext(output_file)
      ext <- tools::file_ext(output_file)
      nms <- names(data)
      if (is.null(nms)) nms <- paste0("sheet", seq_along(data))
      for (i in seq_along(data)) {
        fn <- if (length(data) == 1L) output_file else
              paste0(base, "_", nms[i], ".", ext)
        write.csv(data[[i]], file = fn, row.names = FALSE)
      }
    }
  } else {
    if (is.data.frame(data)) data <- list(data = data)
    write_xlsx(data, path = output_file)
  }
  message("Wrote ", output_file)
}
