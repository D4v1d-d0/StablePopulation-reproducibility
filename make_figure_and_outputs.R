# ============================================================
# make_figure_and_outputs.R
# ============================================================
#
# This script reads the three example sheets in Inputs_3_examples.xlsx,
# performs the beta sweep used in the reproducibility repository,
# writes Outputs_3_examples.xlsx, and generates the figure files in
# both English and Spanish.
# ============================================================

pkgs <- c("readxl", "openxlsx", "StablePopulation")
to_install <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(to_install) > 0) install.packages(to_install)

library(readxl)
library(openxlsx)
library(StablePopulation)

# ============================================================
# 1) GENERAL SETTINGS
# ============================================================

xlsx_path <- "Inputs_3_examples.xlsx"
beta_grid <- seq(0.05, 3.00, by = 0.05)
include_summary_sheet <- FALSE
include_metadata_sheet <- FALSE
extra_journal_format <- "none"
figure_languages <- c("en", "es")

species_sheets <- c(
  "Castor canadensis" = "Castor_Cannadensis_input",
  "Ovis dalli" = "Ovis_dalli_input",
  "Rupicapra rupicapra" = "Rupicapra_Rupicapra_input"
)

species_short <- c(
  "Castor canadensis" = "Castor",
  "Ovis dalli" = "Ovis",
  "Rupicapra rupicapra" = "Rupicapra"
)

panel_titles <- list(
  "Castor canadensis" = expression("A) " * italic(Castor) ~ italic(canadensis)),
  "Ovis dalli" = expression("B) " * italic(Ovis) ~ italic(dalli)),
  "Rupicapra rupicapra" = expression("C) " * italic(Rupicapra) ~ italic(rupicapra))
)

if (!file.exists(xlsx_path)) {
  stop(
    "Input workbook not found: ", xlsx_path,
    "\nAvailable .xlsx files in the current directory: ",
    paste(list.files(pattern = "\\.xlsx$", ignore.case = TRUE), collapse = ", ")
  )
}

to_num <- function(x) {
  x <- as.character(x)
  x <- gsub(",", ".", x)
  x <- gsub("[^0-9\\.\\-eE+]", "", x)
  suppressWarnings(as.numeric(x))
}

read_input_sheet <- function(path, sheet_name) {
  df <- suppressMessages(read_excel(path, sheet = sheet_name, .name_repair = "unique_quiet"))
  needed <- c("age", "lx_obs", "mx")
  if (!all(needed %in% names(df))) {
    stop(
      "Sheet '", sheet_name, "' must contain columns: ",
      paste(needed, collapse = ", "),
      ".\nColumns found: ", paste(names(df), collapse = ", ")
    )
  }
  out <- data.frame(age = to_num(df$age), lx_obs = to_num(df$lx_obs), mx = to_num(df$mx))
  out <- out[complete.cases(out), ]
  out <- out[order(out$age), ]
  rownames(out) <- NULL
  out
}

fit_by_beta_sweep <- function(dat, beta_grid) {
  mx_used <- dat$mx
  results <- lapply(beta_grid, function(b) {
    tryCatch({
      a <- StablePopulation::find_alphas(beta = b, fertility_rates = mx_used)
      pop <- StablePopulation::calculate_population(alpha = a, beta = b, fertility_rates = mx_used)$population
      lx_pred <- pop
      ecm <- mean((lx_pred - dat$lx_obs)^2, na.rm = TRUE)
      rmse <- sqrt(ecm)
      R0_check <- sum(lx_pred * dat$mx, na.rm = TRUE)
      list(
        row = data.frame(beta = b, alpha = a, ECM = ecm, RMSE = rmse, R0_check = R0_check, status = "ok"),
        profile = data.frame(age = dat$age, lx_obs = dat$lx_obs, lx_pred = lx_pred, mx = dat$mx)
      )
    }, error = function(e) {
      list(
        row = data.frame(beta = b, alpha = NA_real_, ECM = NA_real_, RMSE = NA_real_, R0_check = NA_real_, status = paste("error:", conditionMessage(e))),
        profile = NULL
      )
    })
  })
  res <- do.call(rbind, lapply(results, `[[`, "row"))
  profiles_list <- lapply(results, `[[`, "profile")
  valid <- which(is.finite(res$ECM))
  best_i <- valid[which.min(res$ECM[valid])]
  list(summary = res, best = res[best_i, ], profile_best = profiles_list[[best_i]])
}

inputs <- lapply(species_sheets, function(sh) read_input_sheet(xlsx_path, sh))
fits <- lapply(inputs, fit_by_beta_sweep, beta_grid = beta_grid)

# ============================================================
# 2) WRITE OUTPUT EXCEL
# ============================================================

wb_out <- createWorkbook()
if (include_summary_sheet) {
  summary3 <- do.call(rbind, lapply(names(fits), function(sp) data.frame(species = sp, fits[[sp]]$best)))
  addWorksheet(wb_out, "Summary")
  writeData(wb_out, "Summary", summary3)
}
if (include_metadata_sheet) {
  meta <- data.frame(
    item = c("R.version", "StablePopulation.version", "input_file", "beta_min", "beta_max", "beta_step", "constraint", "date"),
    value = c(R.version.string, as.character(packageVersion("StablePopulation")), xlsx_path, min(beta_grid), max(beta_grid), beta_grid[2] - beta_grid[1], "Forced R0 = 1 using observed mx", as.character(Sys.time()))
  )
  addWorksheet(wb_out, "Metadata")
  writeData(wb_out, "Metadata", meta)
}
for (sp in names(fits)) {
  short_name <- species_short[[sp]]
  addWorksheet(wb_out, paste0(short_name, "_beta_sweep"))
  writeData(wb_out, paste0(short_name, "_beta_sweep"), fits[[sp]]$summary)
  addWorksheet(wb_out, paste0(short_name, "_best_fit"))
  writeData(wb_out, paste0(short_name, "_best_fit"), fits[[sp]]$profile_best)
}
saveWorkbook(wb_out, "Outputs_3_examples.xlsx", overwrite = TRUE)

# ============================================================
# 3) PLOTTING HELPERS
# ============================================================

get_plot_labels <- function(language = c("en", "es")) {
  language <- match.arg(language)
  if (language == "es") {
    list(xlab = "Edad (x)", ylab = expression(l[x]), legend = c("Observado", "Weibull"))
  } else {
    list(xlab = "Age (x)", ylab = expression(l[x]), legend = c("Observed", "Weibull"))
  }
}

get_figure_filename <- function(device = c("pdf", "png", "tiff", "jpeg", "jpg", "eps"), language = c("en", "es")) {
  device <- match.arg(device)
  language <- match.arg(language)
  if (language == "en") {
    switch(device,
      pdf  = "Figure1_WeibullExamples.pdf",
      png  = "Figure1_WeibullExamples_300dpi.png",
      tiff = "Figure1_WeibullExamples_300dpi.tiff",
      jpeg = "Figure1_WeibullExamples_300dpi.jpeg",
      jpg  = "Figure1_WeibullExamples_300dpi.jpg",
      eps  = "Figure1_WeibullExamples.eps")
  } else {
    switch(device,
      pdf  = "Figure1_WeibullExamples_ES.pdf",
      png  = "Figure1_WeibullExamples_ES_300dpi.png",
      tiff = "Figure1_WeibullExamples_ES_300dpi.tiff",
      jpeg = "Figure1_WeibullExamples_ES_300dpi.jpeg",
      jpg  = "Figure1_WeibullExamples_ES_300dpi.jpg",
      eps  = "Figure1_WeibullExamples_ES.eps")
  }
}

plot_profile <- function(fit_obj, title_text, language = c("en", "es")) {
  language <- match.arg(language)
  lab <- get_plot_labels(language)
  plot(
    fit_obj$profile_best$age,
    fit_obj$profile_best$lx_obs,
    type = "l",
    lwd = 2,
    lty = 1,
    xlab = lab$xlab,
    ylab = lab$ylab,
    main = title_text,
    ylim = c(-0.05, 1.04)
  )
  lines(
    fit_obj$profile_best$age,
    fit_obj$profile_best$lx_pred,
    lwd = 2,
    lty = 2
  )
  legend(
    "topright",
    legend = lab$legend,
    lwd = 2,
    lty = c(1, 2),
    bty = "n",
    cex = 0.80,
    seg.len = 4.7
  )
}

make_three_panel_figure <- function(device = c("pdf", "png", "tiff", "jpeg", "jpg", "eps"), language = c("en", "es")) {
  device <- match.arg(device)
  language <- match.arg(language)
  filename <- get_figure_filename(device, language)
  if (device == "pdf") {
    pdf(filename, width = 14, height = 4.8)
  } else if (device == "png") {
    png(filename, width = 14, height = 4.8, units = "in", res = 300)
  } else if (device == "tiff") {
    tiff(filename, width = 14, height = 4.8, units = "in", res = 300, compression = "lzw")
  } else if (device %in% c("jpeg", "jpg")) {
    jpeg(filename, width = 14, height = 4.8, units = "in", res = 300, quality = 100)
  } else if (device == "eps") {
    postscript(filename, width = 14, height = 4.8, horizontal = FALSE, onefile = FALSE, paper = "special")
  }
  par(mfrow = c(1, 3), mar = c(4.5, 4.6, 3.0, 1.0), oma = c(0, 0, 0, 0))
  for (sp in names(fits)) {
    plot_profile(fits[[sp]], panel_titles[[sp]], language = language)
  }
  dev.off()
  invisible(filename)
}

# ============================================================
# 4) GENERATE FIGURES
# ============================================================

created_figures <- character(0)
for (lang in figure_languages) {
  created_figures <- c(
    created_figures,
    make_three_panel_figure("pdf", language = lang),
    make_three_panel_figure("png", language = lang),
    make_three_panel_figure("tiff", language = lang),
    make_three_panel_figure("jpg", language = lang)
  )
  if (!identical(extra_journal_format, "none")) {
    created_figures <- c(created_figures, make_three_panel_figure(extra_journal_format, language = lang))
  }
}

message(
  "Done. Files created:\n  ",
  paste(created_figures, collapse = "\n  "),
  "\n  Outputs_3_examples.xlsx"
)
