# Step-by-step reproducibility workflow for the practical case

This guide documents the manuscript-specific workflow used to compare observed survivorship profiles with profiles reconstructed by the **StablePopulation** R package. It complements the ECOSISTEMAS note and the external script `make_figure_and_outputs.R`.

The original analysis described in the note was carried out with **StablePopulation 1.0.3**. This workflow uses only the core functions `find_alphas()` and `calculate_population()`, already available in that release, and is intended to remain compatible with later package versions that preserve those functions. It is separate from the fixed-`beta` Excel interface illustrated in [`examples/run_analysis/`](../examples/run_analysis/README.md).

## Aim of the practical case

For each species, the analysis identifies the discrete Weibull survivorship profile that satisfies the demographic condition `R0 = 1` and best approximates an observed survivorship profile. The selected profile is the one with the smallest mean squared error (MSE; `ECM` in the manuscript).

The three examples in this repository (*Castor canadensis*, *Ovis dalli*, and *Rupicapra rupicapra*) are methodological contrasts. They show how the fit can differ among profiles before applying the approach to fossil assemblages and paleoecological questions.

## Input data

The file `Inputs_3_examples.xlsx` contains one worksheet per species. Each worksheet has three required columns:

| Column | Meaning | Requirement |
|---|---|---|
| `age` | Age class, `x` | Consecutive values beginning at age 0 |
| `lx_obs` | Observed survivorship, `lₓ` | Values between 0 and 1; the first value should equal 1 |
| `mx` | Age-specific fertility, `mₓ` | Non-negative values; expressed as daughters per female |

For example, the first rows for *Castor canadensis* are:

| age | lx_obs | mx |
|---:|---:|---:|
| 0 | 1.000 | 0.000 |
| 1 | 0.481 | 0.000 |
| 2 | 0.460 | 0.315 |
| 3 | 0.275 | 0.400 |
| 4 | 0.178 | 0.895 |

The life-table framework models female demographic parameters. Therefore, when source fertility values refer to offspring of both sexes, they should be converted to daughters per female before running the analysis, under the chosen assumption about the sex ratio at birth.

## Workflow

1. **Read and validate the input table.**  
   The script reads `age`, `lx_obs`, and `mx`, removes incomplete rows, orders the data by age, and checks that ages are consecutive, survivorship lies between 0 and 1, and fertility is non-negative.

2. **Define the `beta` grid.**  
   In the manuscript workflow, `beta` ranges from 0.05 to 3.00 in increments of 0.05.

3. **Obtain `alpha` for each candidate `beta`.**  
   `find_alphas()` uses the fertility vector to obtain the value of `alpha` compatible with the condition `R0 = 1`:

   `sum(lx_pred * mx) = 1`

4. **Generate the reconstructed survivorship profile.**  
   `calculate_population()` uses the resulting `alpha`, the corresponding `beta`, and the fertility vector to generate the discrete reconstructed profile `lx_pred`.

5. **Compare reconstructed and observed profiles.**  
   The script calculates the mean squared error (`ECM`) and its square root (`RMSE`) between `lx_pred` and `lx_obs` for every valid value of `beta`.

6. **Select the best fit.**  
   The value of `beta` with the smallest `ECM` is selected. Its associated `alpha`, `RMSE`, and reconstructed profile are retained.

7. **Write outputs and generate Figure 1.**  
   `Outputs_3_examples.xlsx` contains a complete `beta` sweep and the selected profile for each species. The script also generates equivalent English and Spanish versions of Figure 1 in several formats.

## Core functions used

| Function | Role in this workflow |
|---|---|
| `find_alphas()` | Finds `alpha` compatible with `R0 = 1` for a specified `beta` and fertility vector. |
| `calculate_population()` | Generates the reconstructed discrete survivorship profile and the total number of births. |

The package also includes `weibull_survival()` for evaluating survivorship at a given age and `run_analysis()` for processing Excel-defined cases with a fixed value of `beta`. The latter interface is documented separately in [`examples/run_analysis/`](../examples/run_analysis/README.md).

## Results of the practical case

The selected values below are the rows with the minimum `ECM` in the beta sweep for each species. The full candidate sweeps and the corresponding reconstructed survivorship profiles are available in `Outputs_3_examples.xlsx`.

| Species | Selected beta | Estimated alpha | ECM | RMSE |
|---|---:|---:|---:|---:|
| *Castor canadensis* | 0.80 | 2.049302 | 0.000807 | 0.028409 |
| *Ovis dalli* | 0.55 | 3.371558 | 0.002671 | 0.051685 |
| *Rupicapra rupicapra* | 2.05 | 4.867964 | 0.004167 | 0.064549 |

Figure 1 displays the observed survivorship profile and the selected Weibull reconstruction for each species. Together, the input workbook, output workbook, external script, and figure provide a complete, reproducible path from the original demographic data to the graphical comparison presented in the note.

## Applying the workflow to another species

To adapt this example to another species:

1. Add a worksheet with the columns `age`, `lx_obs`, and `mx`.
2. Add the species and worksheet name to `species_sheets` in `make_figure_and_outputs.R`.
3. Add the short output name to `species_short` and, when needed, define a panel title in `panel_titles`.
4. Run `source("make_figure_and_outputs.R")` from the repository folder.

The script will create a new output workbook and regenerate Figure 1 according to the configured species.

## Files associated with this workflow

- `Inputs_3_examples.xlsx`: observed survivorship and fertility data.
- `make_figure_and_outputs.R`: external reproducibility script.
- `Outputs_3_examples.xlsx`: complete `beta` sweeps and selected profiles.
- `figures/en/Figure1_WeibullExamples.*`: final English figure formats.
- `figures/es/Figure1_WeibullExamples_ES.*`: equivalent final Spanish figure formats.
