# Run Order

## 1) Tree-level to plot summary

Use `TreeLevel_Input.R`.

- Single plot (from workbook `Starting tree list`):
  - source script
  - run `run_treelevel_input()`
  - output: `plot_summary_from_tree.xlsx`

- Batch (data frame with many plots):
  - source script
  - run `run_treelevel_input_batch(...)`
  - output: `plot_summary_from_tree_batch.xlsx`

## 2) Growth from plot summary

Use `Growth_From_PlotSummary.R`.

- It reads:
  - `plot_summary_input.xlsx` if present, else
  - `plot_summary_from_tree.xlsx`
- It writes:
  - `growth_from_plot_summary.xlsx`

Run:
- source script
- run `run_growth_from_plot_summary()`

## 3) Yield from growth

Use `03_Yield_From_Growth.R`.

- It reads:
  - `growth_from_plot_summary.xlsx` (sheet `growth`)
- It writes:
  - `yield_from_growth.xlsx` (sheets `annual_yield`, `rotation_summary`)

Run:
- source script
- run `run_yield_from_growth()`

## 4) Carbon from yield

Use `04_Carbon_From_Yield.R`.

- It reads:
  - `yield_from_growth.xlsx` (sheet `annual_yield`)
- It writes:
  - `carbon_from_yield.xlsx` (sheets `annual_carbon`, `carbon_summary`)

Run:
- source script
- run `run_carbon_from_yield()`

## 5) Consolidated model report

Use `05_Model_Report.R`.

- It reads:
  - `plot_summary_from_tree.xlsx`
  - `growth_from_plot_summary.xlsx`
  - `yield_from_growth.xlsx`
  - `carbon_from_yield.xlsx`
- It writes:
  - `model_chain_outputs.xlsx`

Run:
- source script
- run `run_model_report()`

## 6) One-command chain

Use `Run_Model_Chain.R`.

Run:
- source script
- run `run_full_chain()`

## Minimal workflow

1. Run tree-level script to create summary (`plot_summary_from_tree.xlsx`).
2. Run growth script to generate growth output (`growth_from_plot_summary.xlsx`).
3. Run yield script (`yield_from_growth.xlsx`).
4. Run carbon script (`carbon_from_yield.xlsx`).
5. Run report script (`model_chain_outputs.xlsx`).

## Notes

- Growth is summary-driven; tree-derived values are optional upstream inputs.
- Keep workbook files in the same folder as scripts.
- Carbon step currently uses configurable stem-only conversion defaults; replace with full C_Change logic as you port more VBA functions.
