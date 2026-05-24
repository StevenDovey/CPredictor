# VBA to R Workflow (for `300index2025V1.2.R`)

## 1) Export and organize VBA source

- Keep workbook source as: `C:\R\C Predictor\Multi-Species-Carbon-Calculator-Version-1.0ul.xlsm`
- Keep exported code in: `C:\R\C Predictor\VBA`
- Use purpose-based file names (already applied):
  - `Module1_300Index_Growth_and_WoodDensity.bas`
  - `Module2_MultiSpecies_GrowthModel.bas`
  - `Module3_300Index_InputChecks.bas`
  - `Module4_DouglasFir_500Index_Model.bas`
  - `Module5_CChange_CoreModel.bas`
  - `Module6_CChange_IO_Procedures.bas`
  - `Module7_ForestCarbonPredictor_Orchestrator_and_Batch.bas`
  - `Module8_MultiSpecies_InputErrorChecks.bas`
  - `Sheet41_300Index_RunButton.cls` (key entry button on `300 Index` worksheet)
  - Other `Sheet*.cls` and `ThisWorkbook_WorkbookModule.cls` are workbook/sheet object modules.

## 2) Set translation scope (recommended order)

1. **Core 300 Index parity first**
   - Primary sources:
     - `Module1_300Index_Growth_and_WoodDensity.bas`
     - `Module3_300Index_InputChecks.bas`
     - `Sheet41_300Index_RunButton.cls`
   - Target R file:
     - `C:\R\C Predictor\300index2025V1.2.R`

2. **Multi-species growth**
   - Primary source:
     - `Module2_MultiSpecies_GrowthModel.bas`
   - Keep separate from core 300 Index until core parity is stable.

3. **Carbon/Nutrient model integration**
   - Primary sources:
     - `Module5_CChange_CoreModel.bas`
     - `Module6_CChange_IO_Procedures.bas`
     - `Module7_ForestCarbonPredictor_Orchestrator_and_Batch.bas`

4. **Douglas-fir / 500 Index pathway**
   - Primary source:
     - `Module4_DouglasFir_500Index_Model.bas`

5. **Input validation for non-300 sheets**
   - Primary source:
     - `Module8_MultiSpecies_InputErrorChecks.bas`

## 3) Build a function crosswalk (VBA -> R)

For each VBA function/sub:
- Record VBA name, purpose, inputs, outputs, side effects (worksheet reads/writes), and dependencies.
- Create an R counterpart with the same numeric behavior first (even if style is not ideal).
- Preserve constants exactly (including piecewise logic and hard bounds).

Suggested crosswalk groups:
- **Solvers**: `Bisection`, `Index300`, `CalcAge`, `CalcA200`
- **Height/DBH/BA/Vol**: `CalcMTH`, `DBHmodel`, `CalcDBH`, `CalcBAfromDBH`, `CalcVol`, `calcBAfromVol`
- **Stand process**: `stock`, `thinning`, `Ageshifts`, `Growth`, `earlyield`, `mortvol`, `density`

## 4) Enforce parity rules while porting

- Use fixed-iteration bisection where VBA does.
- Match model routing decisions from sheet flags (height, mortality, volume table).
- Match units and conversions exactly (cm/m, per-ha, log terms, clipping).
- Do not refactor formulas until parity tests pass.

## 5) Build parity test fixtures

Create a small fixed test set (5-10 representative plots):
- young, mid-rotation, old age
- low/high SI
- low/high stocking
- at least one thinning and one pruning scenario

For each fixture, compare VBA vs R outputs:
- `SI`, `I300`, `DBH`, `MTH`, `BA`, `Volume`, `Stocking`, `Wood density`
- Tolerance targets:
  - exact/near-exact for deterministic solver outputs
  - otherwise absolute diff threshold agreed per variable

## 6) Stabilize run pipeline

- Keep one command/script that runs:
  1. input load
  2. SI/I300 solve
  3. annual yield generation
  4. density + mortality outputs
  5. CSV comparison report

- Emit a comparison report each run:
  - columns: plot, age, VBA value, R value, diff, pass/fail

## 7) Only then refactor

After parity is achieved:
- split R into modules (`solvers`, `growth`, `density`, `mortality`, `io`)
- add input validation layer mirroring VBA checks
- improve naming/comments while preserving formulas

