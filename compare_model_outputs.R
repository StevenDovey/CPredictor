# =============================================================================
# compare_model_outputs.R
# Compares VBA (Excel) model vs R model outputs for CPredictor
# Excel outputs: C:/R/CPredictor/FCP_Output    (VBA files)
# R outputs:     C:/R/CPredictor/batch_output  (R files)
#
# Outputs:
# - Standard 1:1 scatter comparison plots
# - Per-plot Stocking b4 thin XY comparison curves (VBA vs R)
# =============================================================================

# =============================================================================
# Libraries
# =============================================================================

library(ggplot2)
library(readr)
library(dplyr)
library(tidyr)

# =============================================================================
# Paths
# =============================================================================

excel_dir <- "C:/R/CPredictor/FCP_Output"
r_dir     <- "C:/R/CPredictor/batch_output"
out_dir   <- "C:/R/CPredictor/comparison_plots"

stocking_dir <- file.path(out_dir, "Stocking")

dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
dir.create(stocking_dir, showWarnings = FALSE, recursive = TRUE)

# =============================================================================
# Read data
# =============================================================================

vba_yield <- read_csv(
  file.path(excel_dir, "yield_tables_VBA.csv"),
  show_col_types = FALSE
)

vba_pp <- read_csv(
  file.path(excel_dir, "plots_processed_VBA.csv"),
  show_col_types = FALSE
)

r_yield <- read_csv(
  file.path(r_dir, "yield_tables_R.csv"),
  show_col_types = FALSE
)

r_pp <- read_csv(
  file.path(r_dir, "plots_processed_R.csv"),
  show_col_types = FALSE
)

# =============================================================================
# Forward-fill VBA yield table
# VBA output leaves Plot and Index age blank after first row
# =============================================================================

vba_yield <- vba_yield %>%
  fill(Plot, `Index age`, .direction = "down")

# =============================================================================
# Remove duplicate rows before joins
# Prevent accidental Cartesian expansion
# =============================================================================

vba_yield <- vba_yield %>%
  distinct(Plot, Age, .keep_all = TRUE)

r_yield <- r_yield %>%
  distinct(Plot, Age, .keep_all = TRUE)

vba_pp <- vba_pp %>%
  distinct(Plot, Age, Species, .keep_all = TRUE)

r_pp <- r_pp %>%
  distinct(Plot, Age, Species, .keep_all = TRUE)

# =============================================================================
# Join VBA and R outputs
# =============================================================================

yt <- inner_join(
  vba_yield,
  r_yield,
  by = c("Plot", "Age"),
  suffix = c("_vba", "_r")
)

pp <- inner_join(
  vba_pp,
  r_pp,
  by = c("Plot", "Age", "Species"),
  suffix = c("_vba", "_r")
)

# =============================================================================
# Generic 1:1 scatter plot helper
# =============================================================================

save_plot <- function(x, y, title, filename) {
  
  df <- data.frame(
    vba = suppressWarnings(as.numeric(x)),
    r   = suppressWarnings(as.numeric(y))
  )
  
  df <- df %>%
    filter(is.finite(vba), is.finite(r))
  
  if (nrow(df) == 0) {
    message("Skipped empty plot: ", filename)
    return(NULL)
  }
  
  axis_lim <- range(c(df$vba, df$r), na.rm = TRUE)
  
  p <- ggplot(df, aes(x = vba, y = r)) +
    geom_point(
      alpha = 0.5,
      size = 1.2,
      colour = "#2166ac"
    ) +
    geom_abline(
      intercept = 0,
      slope = 1,
      colour = "red",
      linewidth = 0.7
    ) +
    coord_equal(
      xlim = axis_lim,
      ylim = axis_lim
    ) +
    labs(
      title = title,
      x = "VBA model",
      y = "R model"
    ) +
    theme_bw(base_size = 9)
  
  ggsave(
    file.path(out_dir, filename),
    plot = p,
    width = 7.5,
    height = 5,
    units = "cm",
    dpi = 300
  )
  
  message("Saved: ", filename)
}

# =============================================================================
# Per-plot Stocking b4 thin comparison
# Composite XY plot by Age
# =============================================================================

save_stocking_plot <- function(plot_id, data) {
  
  df <- data %>%
    filter(Plot == plot_id) %>%
    select(
      Plot,
      Age,
      `Stocking b4 thin_vba`,
      `Stocking b4 thin_r`
    ) %>%
    mutate(
      `Stocking b4 thin_vba` =
        suppressWarnings(as.numeric(`Stocking b4 thin_vba`)),
      
      `Stocking b4 thin_r` =
        suppressWarnings(as.numeric(`Stocking b4 thin_r`))
    ) %>%
    filter(
      is.finite(Age),
      is.finite(`Stocking b4 thin_vba`),
      is.finite(`Stocking b4 thin_r`)
    ) %>%
    arrange(Age)
  
  if (nrow(df) == 0) {
    return(NULL)
  }
  
  df_long <- df %>%
    pivot_longer(
      cols = c(
        `Stocking b4 thin_vba`,
        `Stocking b4 thin_r`
      ),
      names_to = "Model",
      values_to = "Stocking"
    ) %>%
    mutate(
      Model = recode(
        Model,
        `Stocking b4 thin_vba` = "VBA",
        `Stocking b4 thin_r`   = "R"
      )
    )
  
  p <- ggplot(
    df_long,
    aes(
      x = Age,
      y = Stocking,
      colour = Model
    )
  ) +
    geom_line(linewidth = 0.8) +
    geom_point(size = 1.3) +
    labs(
      title = paste("Plot", plot_id),
      x = "Age",
      y = "Stocking b4 thin"
    ) +
    theme_bw(base_size = 9)
  
  safe_name <- gsub("[^A-Za-z0-9_\\-]", "_", plot_id)
  
  ggsave(
    file.path(
      stocking_dir,
      paste0(safe_name, ".png")
    ),
    plot = p,
    width = 9,
    height = 6,
    units = "cm",
    dpi = 300
  )
}

# =============================================================================
# plots_processed comparisons
# =============================================================================

save_plot(
  pp$`Site Index_vba`,
  pp$`Site Index_r`,
  "Site Index",
  "01_SI.png"
)

save_plot(
  pp$`300 Index_vba`,
  pp$`300 Index_r`,
  "300 Index",
  "02_300i.png"
)

# =============================================================================
# yield_tables comparisons
# =============================================================================

save_plot(
  yt$`Stocking b4 thin_vba`,
  yt$`Stocking b4 thin_r`,
  "Stocking b4 thin",
  "08_stocking_b4_thin.png"
)

save_plot(
  yt$MTH_vba,
  yt$MTH_r,
  "MTH",
  "09_MTH.png"
)

save_plot(
  yt$`Volume b4 thin_vba`,
  yt$`Volume b4 thin_r`,
  "Volume b4 thin",
  "10_volume_b4_thin.png"
)

save_plot(
  yt$`BA b4 thin_vba`,
  yt$`BA b4 thin_r`,
  "BA b4 thin",
  "11_BA_b4_thin.png"
)

save_plot(
  yt$`DBH b4 thin_vba`,
  yt$`DBH b4 thin_r`,
  "DBH b4 thin",
  "12_DBH_b4_thin.png"
)

save_plot(
  yt$`Mean Height_vba`,
  yt$`Mean Height_r`,
  "Mean Height",
  "13_mean_height.png"
)

# =============================================================================
# Create per-plot Stocking b4 thin plots
# =============================================================================

plot_ids <- unique(yt$Plot)

for (p in plot_ids) {
  save_stocking_plot(p, yt)
}

message("\nAll comparison plots written to: ", out_dir)
message("Stocking plots written to: ", stocking_dir)