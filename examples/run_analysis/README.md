# `run_analysis()` example

This folder provides an example of the Excel-based `run_analysis()` workflow.

## Files

- `Input_Data.xlsx` illustrates the Excel layout expected by `run_analysis()`.
- `outputs/` contains the result files generated from the worksheets in `Input_Data.xlsx`.

## Input structure

Each worksheet in `Input_Data.xlsx` represents one case. Fertility rates (`mₓ`) are provided in column B, and a fixed value of `β` is specified in cell C2.

## Scope of this example

This material documents the package interface for processing Excel-defined cases with a fixed value of `β`.

It is separate from the manuscript-specific reproducibility workflow, which performs a scan across values of `β`, compares reconstructed and observed survivorship profiles, calculates ECM and RMSE, and generates Figure 1.
