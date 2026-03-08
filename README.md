# make_figure_and_outputs_from_inputs_only.R
# ------------------------------------------------------------
# Input:
#   - Inputs_3_examples.xlsx
# Output only:
#   - Outputs_3_examples.xlsx
#   - Figure1_WeibullExamples.pdf
#   - Figure1_WeibullExamples_300dpi.png
#   - Figure1_WeibullExamples_300dpi.tiff
#   - Figure1_WeibullExamples_300dpi.jpg
#   - Figure1_WeibullExamples_300dpi.<other_format>  # optional extra journal file
#
# The input workbook must contain the sheets:
#   - Castor_canadensis_input
#   - Ovis_dalli_input
#   - Rangifer_tarandus_input
# with columns: age, lx_obs, mx
# ------------------------------------------------------------

# 0) Packages
pkgs <- c("readxl", "openxlsx", "StablePopulation")
to_install <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(to_install) > 0) install.packages(to_install)

library(readxl)
library(openxlsx)
library(StablePopulation)

# 1) File paths and export options
xlsx_path <- "Inputs_3_examples.xlsx"

# The script always creates TIFF and JPG copies at 300 dpi.
# You may optionally request one extra format among: "none", "jpeg", "eps".
extra_journal_format <- "none"

if (!file.exists(xlsx_path)) {
  stop(
    "No encuentro el Excel de entrada: ", xlsx_path,
    "\nArchivos .xlsx disponibles: ",
    paste(list.files(pattern = "\\.xlsx$", ignore.case = TRUE), collapse = ", ")
  )
}

# ------------------------------------------------------------
# Helper: robust numeric conversion
# ------------------------------------------------------------
to_num <- function(x) {
  x <- as.character(x)
  x <- gsub(",", ".", x)
  x <- gsub("[^0-9\\.\\-eE+]", "", x)
  suppressWarnings(as.numeric(x))
}

# ------------------------------------------------------------
# Helper: read one already-prepared input sheet
# ------------------------------------------------------------
read_input_sheet <- function(path, sheet_name) {
  df <- suppressMessages(read_excel(path, sheet = sheet_name, .name_repair = "unique_quiet"))

  needed <- c("age", "lx_obs", "mx")
  if (!all(needed %in% names(df))) {
    stop(
      "La hoja '", sheet_name, "' debe contener las columnas: ",
      paste(needed, collapse = ", "),
      ".\nColumnas encontradas: ",
      paste(names(df), collapse = ", ")
    )
  }

  out <- data.frame(
    age    = to_num(df$age),
    lx_obs = to_num(df$lx_obs),
    mx     = to_num(df$mx)
  )

  out <- out[complete.cases(out), ]
  out <- out[order(out$age), ]
  rownames(out) <- NULL
  out
}

# ------------------------------------------------------------
# Core routine: beta sweep -> alpha(beta) -> ECM/RMSE -> best
# ------------------------------------------------------------
fit_by_beta_sweep <- function(dat, beta_grid = seq(0.05, 3.00, by = 0.05)) {

  dat <- dat[order(dat$age), ]

  if (any(duplicated(dat$age))) {
    stop("Hay edades duplicadas en el dataset.")
  }
  if (!all(diff(dat$age) == 1)) {
    stop("Las edades no son consecutivas (faltan clases de edad).")
  }
  if (!all(c("age", "lx_obs", "mx") %in% names(dat))) {
    stop("El dataset debe tener columnas: age, lx_obs, mx.")
  }

  R0_obs <- sum(dat$lx_obs * dat$mx, na.rm = TRUE)
  mx_scaled <- dat$mx / R0_obs

  if (sum(mx_scaled, na.rm = TRUE) < 1) {
    stop("No existe raÃ­z porque sum(mx/R0) < 1. Revisa mx y/o R0.")
  }

  res <- lapply(beta_grid, function(b) {
    a <- StablePopulation::find_alphas(beta = b, fertility_rates = mx_scaled)

    pop <- StablePopulation::calculate_population(
      alpha = a,
      beta = b,
      fertility_rates = mx_scaled
    )$population

    lx_pred <- pop
    ecm  <- mean((lx_pred - dat$lx_obs)^2, na.rm = TRUE)
    rmse <- sqrt(ecm)
    R0_check <- sum(lx_pred * dat$mx, na.rm = TRUE)

    data.frame(
      beta = b,
      alpha = a,
      ECM = ecm,
      RMSE = rmse,
      R0_obs = R0_obs,
      R0_check = R0_check
    )
  })

  res <- do.call(rbind, res)
  best_i <- which.min(res$ECM)
  best <- res[best_i, ]

  mx_scaled_best <- dat$mx / best$R0_obs
  pop_best <- StablePopulation::calculate_population(
    alpha = best$alpha,
    beta = best$beta,
    fertility_rates = mx_scaled_best
  )$population

  prof_best <- data.frame(
    age = dat$age,
    lx_obs = dat$lx_obs,
    lx_pred = pop_best,
    mx = dat$mx
  )

  list(summary = res, best = best, profile_best = prof_best)
}

# ------------------------------------------------------------
# Load the 3 prepared example datasets
# ------------------------------------------------------------
castor <- read_input_sheet(xlsx_path, "Castor_canadensis_input")
ovis   <- read_input_sheet(xlsx_path, "Ovis_dalli_input")
rang   <- read_input_sheet(xlsx_path, "Rangifer_tarandus_input")

# ------------------------------------------------------------
# Fit
# ------------------------------------------------------------
beta_grid <- seq(0.05, 3.00, by = 0.05)

fit_castor <- fit_by_beta_sweep(castor, beta_grid)
fit_ovis   <- fit_by_beta_sweep(ovis, beta_grid)
fit_rang   <- fit_by_beta_sweep(rang, beta_grid)

# ------------------------------------------------------------
# Save OUTPUTS only (including compact summary inside workbook)
# ------------------------------------------------------------
summary3 <- rbind(
  data.frame(species = "Castor canadensis", fit_castor$best),
  data.frame(species = "Ovis dalli", fit_ovis$best),
  data.frame(species = "Rangifer tarandus", fit_rang$best)
)

wb_out <- createWorkbook()

addWorksheet(wb_out, "Summary")
writeData(wb_out, "Summary", summary3)

addWorksheet(wb_out, "Castor_beta_sweep")
writeData(wb_out, "Castor_beta_sweep", fit_castor$summary)
addWorksheet(wb_out, "Castor_best_profile")
writeData(wb_out, "Castor_best_profile", fit_castor$profile_best)

addWorksheet(wb_out, "Ovis_beta_sweep")
writeData(wb_out, "Ovis_beta_sweep", fit_ovis$summary)
addWorksheet(wb_out, "Ovis_best_profile")
writeData(wb_out, "Ovis_best_profile", fit_ovis$profile_best)

addWorksheet(wb_out, "Rangifer_beta_sweep")
writeData(wb_out, "Rangifer_beta_sweep", fit_rang$summary)
addWorksheet(wb_out, "Rangifer_best_profile")
writeData(wb_out, "Rangifer_best_profile", fit_rang$profile_best)

saveWorkbook(wb_out, "Outputs_3_examples.xlsx", overwrite = TRUE)

# ------------------------------------------------------------
# Plot helper
# ------------------------------------------------------------
plot_profile <- function(fit_obj, title_text) {
  plot(
    fit_obj$profile_best$age,
    fit_obj$profile_best$lx_obs,
    type = "l",
    lwd = 2,
    xlab = "Edad (x)",
    ylab = expression(l[x]),
    main = title_text,
    sub = sprintf(
      "beta*=%.2f  alpha*=%.3f  RMSE=%.3f",
      fit_obj$best$beta,
      fit_obj$best$alpha,
      fit_obj$best$RMSE
    ),
    cex.sub = 0.80
  )
  lines(fit_obj$profile_best$age, fit_obj$profile_best$lx_pred, lwd = 2, lty = 2)
  legend(
    "topright",
    legend = c("Observado", "Weibull"),
    lwd = 2,
    lty = c(1, 2),
    bty = "n",
    cex = 0.80
  )
}

make_three_panel_figure <- function(device = c("pdf", "png", "tiff", "jpeg", "jpg", "eps")) {
  device <- match.arg(device)

  filename <- switch(
    device,
    pdf  = "Figure1_WeibullExamples.pdf",
    png  = "Figure1_WeibullExamples_300dpi.png",
    tiff = "Figure1_WeibullExamples_300dpi.tiff",
    jpeg = "Figure1_WeibullExamples_300dpi.jpeg",
    jpg  = "Figure1_WeibullExamples_300dpi.jpg",
    eps  = "Figure1_WeibullExamples.eps"
  )

  if (device == "pdf") {
    pdf(filename, width = 11, height = 4.2)
  } else if (device == "png") {
    png(filename, width = 11, height = 4.2, units = "in", res = 300)
  } else if (device == "tiff") {
    tiff(filename, width = 11, height = 4.2, units = "in", res = 300, compression = "lzw")
  } else if (device %in% c("jpeg", "jpg")) {
    jpeg(filename, width = 11, height = 4.2, units = "in", res = 300, quality = 100)
  } else if (device == "eps") {
    postscript(filename, width = 11, height = 4.2, horizontal = FALSE, onefile = FALSE, paper = "special")
  }

  par(
    mfrow = c(1, 3),
    mar = c(4.2, 4.2, 2.8, 1.0),
    oma = c(0, 0, 2.0, 0)
  )

  plot_profile(fit_castor, "A) Castor canadensis (excelente)")
  plot_profile(fit_ovis,   "B) Ovis dalli (medio)")
  plot_profile(fit_rang,   "C) Rangifer tarandus (peor)")

  mtext(
    expression(paste("Figura 1. Ejemplos de ajuste (observado vs Weibull) seleccionados por barrido de ", beta)),
    outer = TRUE,
    cex = 0.95
  )

  dev.off()
}

# ------------------------------------------------------------
# Create figures
# ------------------------------------------------------------
make_three_panel_figure("pdf")
make_three_panel_figure("png")

make_three_panel_figure("tiff")
make_three_panel_figure("jpg")

if (!identical(extra_journal_format, "none")) {
  make_three_panel_figure(extra_journal_format)
}

message(
  "Done. Files created:\n",
  "  Outputs_3_examples.xlsx\n",
  "  Figure1_WeibullExamples.pdf\n",
  "  Figure1_WeibullExamples_300dpi.png\n",
  "  Figure1_WeibullExamples_300dpi.tiff\n",
  "  Figure1_WeibullExamples_300dpi.jpg",
  if (!identical(extra_journal_format, "none")) {
    paste0("\n  Figure1_WeibullExamples_300dpi.", extra_journal_format)
  } else {
    ""
  }
)
