# ============================================================
# make_figure_and_outputs.R
# ============================================================
#
# OBJETIVO
# --------
# A partir de un Excel de entrada con 3 hojas (una por especie),
# este script:
#
#   1) Lee las tablas de entrada.
#   2) Usa directamente los mx ORIGINALES de cada especie.
#   3) Para cada beta del barrido:
#         - calcula el alpha que hace R0 = 1
#           usando esos mx originales
#         - genera el perfil predicho
#         - calcula ECM y RMSE frente a lx observado
#   4) Elige el mejor ajuste (mínimo ECM).
#   5) Exporta un Excel con resultados.
#   6) Genera una figura con tres paneles.
#
# IMPORTANTE
# ----------
# Este script NO calcula R0 observado y NO reescala mx.
#
# Es decir, aquí se fuerza artificialmente:
#
#       sum(lx_pred * mx) = 1
#
# usando los mx originales de cada especie.
#
# ENTRADA
# -------
#   - Inputs_3_examples.xlsx
#
# HOJAS ESPERADAS
# ---------------
#   - Castor_Cannadensis_input
#   - Ovis_dalli_input
#   - Rupicapra_Rupicapra_input
#
# COLUMNAS ESPERADAS EN CADA HOJA
# -------------------------------
#   - age
#   - lx_obs
#   - mx
#
# SALIDAS
# -------
#   - Outputs_3_examples.xlsx
#   - Figure1_WeibullExamples.pdf
#   - Figure1_WeibullExamples_300dpi.png
#   - Figure1_WeibullExamples_300dpi.tiff
#   - Figure1_WeibullExamples_300dpi.jpg
#   - opcionalmente Figure1_WeibullExamples.eps
# ============================================================


# ============================================================
# 0) PAQUETES NECESARIOS
# ============================================================
#
# Si falta algún paquete, se instala automáticamente.
# Esto incluye StablePopulation, que puede instalarse desde CRAN.
# ============================================================

pkgs <- c("readxl", "openxlsx", "StablePopulation")

to_install <- pkgs[!vapply(pkgs, requireNamespace, logical(1), quietly = TRUE)]

if (length(to_install) > 0) {
  install.packages(to_install)
}

library(readxl)
library(openxlsx)
library(StablePopulation)


# ============================================================
# 1) CONFIGURACIÓN GENERAL
# ============================================================

# Archivo Excel de entrada
xlsx_path <- "Inputs_3_examples.xlsx"

# Barrido de beta
beta_grid <- seq(0.05, 3.00, by = 0.05)

# Incluir o no la hoja resumen en el Excel de salida
include_summary_sheet <- FALSE

# Añadir o no una hoja de metadatos
include_metadata_sheet <- TRUE

# Formato gráfico adicional opcional: "none", "jpeg", "eps"
extra_journal_format <- "none"

# Especies y nombres de sus hojas de entrada
# ORDEN FIJADO: Castor -> Ovis -> Rupicapra
species_sheets <- c(
  "Castor canadensis" = "Castor_Cannadensis_input",
  "Ovis dalli" = "Ovis_dalli_input",
  "Rupicapra rupicapra" = "Rupicapra_Rupicapra_input"
)

# Nombres cortos para hojas Excel de salida
species_short <- c(
  "Castor canadensis" = "Castor",
  "Ovis dalli" = "Ovis",
  "Rupicapra rupicapra" = "Rupicapra"
)

# Títulos de los paneles de la figura
# Nombres científicos en cursiva, sin etiquetas valorativas.
panel_titles <- list(
  "Castor canadensis" = expression("A) " * italic(Castor) ~ italic(canadensis)),
  "Ovis dalli" = expression("B) " * italic(Ovis) ~ italic(dalli)),
  "Rupicapra rupicapra" = expression("C) " * italic(Rupicapra) ~ italic(rupicapra))
)


# ============================================================
# 1.1) COMPROBAR QUE EL EXCEL DE ENTRADA EXISTA
# ============================================================

if (!file.exists(xlsx_path)) {
  stop(
    "No encuentro el Excel de entrada: ", xlsx_path,
    "\nArchivos .xlsx disponibles en la carpeta actual: ",
    paste(list.files(pattern = "\\.xlsx$", ignore.case = TRUE), collapse = ", ")
  )
}


# ============================================================
# 2) FUNCIÓN AUXILIAR: CONVERSIÓN ROBUSTA A NUMÉRICO
# ============================================================

to_num <- function(x) {
  x <- as.character(x)
  x <- gsub(",", ".", x)
  x <- gsub("[^0-9\\.\\-eE+]", "", x)
  suppressWarnings(as.numeric(x))
}


# ============================================================
# 3) LEER Y LIMPIAR UNA HOJA DE ENTRADA
# ============================================================

read_input_sheet <- function(path, sheet_name) {
  
  df <- suppressMessages(
    read_excel(path, sheet = sheet_name, .name_repair = "unique_quiet")
  )
  
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
  
  # Eliminar filas con NA
  out <- out[complete.cases(out), ]
  
  # Ordenar por edad
  out <- out[order(out$age), ]
  rownames(out) <- NULL
  
  # Validaciones básicas
  if (nrow(out) == 0) {
    stop("La hoja '", sheet_name, "' no contiene datos válidos.")
  }
  if (any(duplicated(out$age))) {
    stop("La hoja '", sheet_name, "' contiene edades duplicadas.")
  }
  if (!all(diff(out$age) == 1)) {
    stop("La hoja '", sheet_name, "' tiene edades no consecutivas.")
  }
  if (any(out$mx < 0, na.rm = TRUE)) {
    stop("La hoja '", sheet_name, "' contiene mx negativas.")
  }
  if (any(out$lx_obs < 0 | out$lx_obs > 1, na.rm = TRUE)) {
    stop("La hoja '", sheet_name, "' contiene lx_obs fuera de [0,1].")
  }
  if (!isTRUE(all.equal(out$lx_obs[1], 1))) {
    warning(
      "En la hoja '", sheet_name,
      "', lx_obs en la primera edad no vale exactamente 1."
    )
  }
  
  out
}


# ============================================================
# 4) AJUSTE POR BARRIDO DE BETA FORZANDO R0 = 1 SIN REESCALAR
# ============================================================
#
# IDEA DEL CÁLCULO
# ----------------
# A diferencia del script anterior:
#
#   - aquí NO se calcula R0 observado
#   - aquí NO se reescala mx como mx / R0_obs
#
# Lo que se hace es usar directamente los mx originales y pedir
# a find_alphas() que halle el alpha que hace:
#
#       sum(lx_pred * mx) = 1
#
# para cada beta del barrido.
#
# En lugar de recalcular el mejor perfil al final, aquí se guarda
# el perfil lx_pred de cada beta durante el propio barrido.
# Después, cuando se identifica el beta óptimo, simplemente se
# recupera el perfil ya calculado.
# ============================================================

fit_by_beta_sweep <- function(dat, beta_grid) {
  
  # ----------------------------------------------------------
  # 4.1) FERTILIDADES USADAS EN EL AJUSTE
  # ----------------------------------------------------------
  #
  # Aquí usamos directamente los mx originales, SIN reescalado.
  # ----------------------------------------------------------
  mx_used <- dat$mx
  
  if (sum(mx_used, na.rm = TRUE) < 1) {
    stop("No existe raíz porque sum(mx) < 1. Revisa mx.")
  }
  
  # ----------------------------------------------------------
  # 4.2) BARRIDO DE BETA
  # ----------------------------------------------------------
  #
  # Para cada beta guardamos:
  #   - una fila resumen con alpha, ECM, RMSE, etc.
  #   - el perfil lx_pred correspondiente
  #
  # Si un beta falla, la fila resumen lleva NA y el perfil se
  # guarda como NULL.
  # ----------------------------------------------------------
  
  results <- lapply(beta_grid, function(b) {
    
    tryCatch({
      # Hallar alpha para este beta imponiendo R0 = 1
      a <- StablePopulation::find_alphas(
        beta = b,
        fertility_rates = mx_used
      )
      
      # Generar perfil poblacional / perfil predicho
      pop <- StablePopulation::calculate_population(
        alpha = a,
        beta = b,
        fertility_rates = mx_used
      )$population
      
      # Comprobación de longitud
      if (length(pop) != nrow(dat)) {
        stop("La longitud de population no coincide con lx_obs.")
      }
      
      lx_pred <- pop
      
      # Error cuadrático medio y su raíz
      ecm  <- mean((lx_pred - dat$lx_obs)^2, na.rm = TRUE)
      rmse <- sqrt(ecm)
      
      # Comprobación de consistencia con la restricción forzada
      R0_check <- sum(lx_pred * dat$mx, na.rm = TRUE)
      
      list(
        row = data.frame(
          beta = b,
          alpha = a,
          ECM = ecm,
          RMSE = rmse,
          R0_check = R0_check,
          status = "ok"
        ),
        profile = data.frame(
          age = dat$age,
          lx_obs = dat$lx_obs,
          lx_pred = lx_pred,
          mx = dat$mx
        )
      )
      
    }, error = function(e) {
      list(
        row = data.frame(
          beta = b,
          alpha = NA_real_,
          ECM = NA_real_,
          RMSE = NA_real_,
          R0_check = NA_real_,
          status = paste("error:", conditionMessage(e))
        ),
        profile = NULL
      )
    })
  })
  
  # Tabla resumen del barrido completo
  res <- do.call(rbind, lapply(results, `[[`, "row"))
  
  # Lista de perfiles (uno por beta; NULL si falló)
  profiles_list <- lapply(results, `[[`, "profile")
  
  # ----------------------------------------------------------
  # 4.3) ELEGIR EL MEJOR AJUSTE ENTRE LOS BETA VÁLIDOS
  # ----------------------------------------------------------
  valid <- which(is.finite(res$ECM))
  
  if (length(valid) == 0) {
    stop("Ningún beta produjo un ajuste válido.")
  }
  
  best_i <- valid[which.min(res$ECM[valid])]
  best <- res[best_i, ]
  
  # ----------------------------------------------------------
  # 4.4) RECUPERAR DIRECTAMENTE EL MEJOR PERFIL
  # ----------------------------------------------------------
  profile_best <- profiles_list[[best_i]]
  
  list(
    summary = res,
    best = best,
    profile_best = profile_best
  )
}


# ============================================================
# 5) LEER LOS DATOS DE ENTRADA DE TODAS LAS ESPECIES
# ============================================================

inputs <- lapply(species_sheets, function(sh) {
  read_input_sheet(xlsx_path, sh)
})


# ============================================================
# 6) AJUSTAR TODAS LAS ESPECIES
# ============================================================

fits <- lapply(inputs, fit_by_beta_sweep, beta_grid = beta_grid)


# ============================================================
# 7) CREAR EL EXCEL DE SALIDA
# ============================================================

wb_out <- createWorkbook()

# ------------------------------------------------------------
# 7.1) HOJA RESUMEN OPCIONAL
# ------------------------------------------------------------
if (include_summary_sheet) {
  summary3 <- do.call(
    rbind,
    lapply(names(fits), function(sp) {
      data.frame(species = sp, fits[[sp]]$best)
    })
  )
  
  addWorksheet(wb_out, "Summary")
  writeData(wb_out, "Summary", summary3)
}

# ------------------------------------------------------------
# 7.2) HOJA DE METADATOS OPCIONAL
# ------------------------------------------------------------
if (include_metadata_sheet) {
  meta <- data.frame(
    item = c(
      "R.version",
      "StablePopulation.version",
      "input_file",
      "beta_min",
      "beta_max",
      "beta_step",
      "constraint",
      "date"
    ),
    value = c(
      R.version.string,
      as.character(packageVersion("StablePopulation")),
      xlsx_path,
      min(beta_grid),
      max(beta_grid),
      beta_grid[2] - beta_grid[1],
      "Forced R0 = 1 using original mx (no rescaling)",
      as.character(Sys.time())
    )
  )
  
  addWorksheet(wb_out, "Metadata")
  writeData(wb_out, "Metadata", meta)
}

# ------------------------------------------------------------
# 7.3) HOJAS POR ESPECIE
# ------------------------------------------------------------
for (sp in names(fits)) {
  
  short_name <- species_short[[sp]]
  
  # Hoja con todo el barrido de beta
  addWorksheet(wb_out, paste0(short_name, "_beta_sweep"))
  writeData(wb_out, paste0(short_name, "_beta_sweep"), fits[[sp]]$summary)
  
  # Hoja con el mejor perfil
  addWorksheet(wb_out, paste0(short_name, "_best_fit"))
  writeData(wb_out, paste0(short_name, "_best_fit"), fits[[sp]]$profile_best)
}

# Guardar Excel
saveWorkbook(wb_out, "Outputs_3_examples.xlsx", overwrite = TRUE)


# ============================================================
# 8) FUNCIÓN PARA DIBUJAR UN PANEL
# ============================================================
#
# Figura completamente en inglés para mantener consistencia
# con el nombre de archivo y el título general.
# ============================================================

plot_profile <- function(fit_obj, title_text) {
  plot(
    fit_obj$profile_best$age,
    fit_obj$profile_best$lx_obs,
    type = "l",
    lwd = 2,
    xlab = "Age (x)",
    ylab = expression(l[x]),
    main = title_text,
    sub = sprintf(
      "beta*=%.2f   alpha*=%.3f   RMSE=%.3f",
      fit_obj$best$beta,
      fit_obj$best$alpha,
      fit_obj$best$RMSE
    ),
    cex.sub = 0.80,
    ylim = c(0, 1)
  )
  
  lines(
    fit_obj$profile_best$age,
    fit_obj$profile_best$lx_pred,
    lwd = 2,
    lty = 2
  )
  
  legend(
    "topright",
    legend = c("Observed", "Weibull"),
    lwd = 2,
    lty = c(1, 2),
    bty = "n",
    cex = 0.80
  )
}


# ============================================================
# 9) FUNCIÓN PARA CREAR LA FIGURA DE TRES PANELES
# ============================================================

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
  
  # Abrir el dispositivo gráfico adecuado
  if (device == "pdf") {
    pdf(filename, width = 11, height = 4.2)
  } else if (device == "png") {
    png(filename, width = 11, height = 4.2, units = "in", res = 300)
  } else if (device == "tiff") {
    tiff(filename, width = 11, height = 4.2, units = "in", res = 300, compression = "lzw")
  } else if (device %in% c("jpeg", "jpg")) {
    jpeg(filename, width = 11, height = 4.2, units = "in", res = 300, quality = 100)
  } else if (device == "eps") {
    postscript(
      filename,
      width = 11,
      height = 4.2,
      horizontal = FALSE,
      onefile = FALSE,
      paper = "special"
    )
  }
  
  # 1 fila x 3 columnas
  par(
    mfrow = c(1, 3),
    mar = c(4.2, 4.2, 2.8, 1.0),
    oma = c(0, 0, 2.0, 0)
  )
  
  # Dibujar especies en orden
  for (sp in names(fits)) {
    plot_profile(fits[[sp]], panel_titles[[sp]])
  }
  
  # Título general
  mtext(
    expression(
      paste("Figure 1. Observed and Weibull-predicted ", l[x], " profiles selected by ", beta, " sweep")
    ),
    outer = TRUE,
    cex = 0.95
  )
  
  dev.off()
  
  invisible(filename)
}


# ============================================================
# 10) CREAR LAS FIGURAS
# ============================================================

created_figures <- c(
  make_three_panel_figure("pdf"),
  make_three_panel_figure("png"),
  make_three_panel_figure("tiff"),
  make_three_panel_figure("jpg")
)

if (!identical(extra_journal_format, "none")) {
  created_figures <- c(created_figures, make_three_panel_figure(extra_journal_format))
}


# ============================================================
# 11) MENSAJE FINAL
# ============================================================

message(
  "Done. Files created:\n  ",
  paste(created_figures, collapse = "\n  "),
  "\n  Outputs_3_examples.xlsx"
)
