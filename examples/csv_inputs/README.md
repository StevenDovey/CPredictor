# CSV Input Templates for CPredictor

These CSV files replace the Excel input workbook sheets. You can use either
CSV files **or** the original `.xlsx` workbook — the R code auto-detects the
format from the file extension.

## Multi-sheet workbook → directory of CSVs

The Excel workbook `input.xlsx` contains multiple sheets. To use CSVs instead,
create a **directory** (e.g. `my_inputs/`) and place one CSV per sheet:

| Excel sheet          | CSV filename              | Description                                    |
|----------------------|---------------------------|------------------------------------------------|
| `Inputs`             | `inputs.csv`              | Stand info, thinning/pruning, site parameters  |
| `300 Index`          | `300_index.csv`           | Site productivity indices, stocking history     |
| `Starting tree list` | `starting_tree_list.csv`  | Individual tree DBH and height measurements    |
| `parameters`         | `parameters.csv`          | Model coefficients (usually not user-edited)   |
| `VolTab`             | `voltab.csv`              | Volume table coefficients (usually not edited)  |

Then pass the directory path instead of the `.xlsx` path:

```r
# Instead of:
run_full_chain(input_workbook = "input.xlsx")

# Use:
run_full_chain(input_workbook = "my_inputs/")
```

## Single-file CSVs

For pipeline steps that read a single-sheet file, just use `.csv` directly:

```r
# Tree data
run_tree_to_plot_summary(tree_file = "tree_data.csv")

# Plot site data
run_tree_to_plot_summary(plot_site_file = "plot_site_data.csv")

# Growth / yield / carbon intermediate files
run_yield_from_growth(growth_file = "growth.csv")
run_carbon_from_yield(yield_file = "yield.csv")
```

## Spatial raster files

Raster inputs (DEM, soil C/N/P, temperature) remain as GeoTIFF `.tif` files.
These are **not** converted to CSV.

## File descriptions

### inputs.csv
Two-column format: `field,value`. Contains stand information (species, 300-index,
site index, stocking, rotation length), thinning schedule (up to 4 thins),
pruning schedule (up to 4 lifts), site productivity options, measurement metrics,
and environmental parameters (latitude, elevation, soil, temperature).

### 300_index.csv
Site productivity indices and stocking history for the 300 Index model.
Top section is `field,value` pairs; bottom section is the stocking history table.

### starting_tree_list.csv
Row 1: `Plot_area_ha,<value>`
Row 2: `Age_years,<value>`
Row 3: blank
Row 4+: `Stem_number,DBH_cm,Height_m` (header + data rows).
Height can be blank for stems without height measurements.

### plot_site_data.csv
Standard tabular CSV with columns:
`Plot_Id, Latitude, Longitude, Easting, Northing, Elevation, Soil_C, Soil_N, Soil_P, Mean_Temp, Species`
