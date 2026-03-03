required_pkgs <- c(
  "shiny",
  "shinyFiles",
  "tidyr",
  "dplyr",
  "openxlsx",
  "stringr",
  "readr",
  "janitor"
)

for (pkg in required_pkgs) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
  }
  library(pkg, character.only = TRUE)
}

# ---------------------------
# Helpers
# ---------------------------

sanitize_sheet_name <- function(x) {
  x <- str_replace_all(x, "[:\\\\/\\?\\*\\[\\]]", "_")
  x <- str_trim(x)
  if (nchar(x) == 0) x <- "Sheet"
  substr(x, 1, 31)
}

fix_colnames_for_excel <- function(df) {
  df <- as.data.frame(df)
  nm <- names(df)
  if (is.null(nm)) stop("NULL column names detected")

  bad <- is.na(nm) | nm == ""
  if (any(bad)) nm[bad] <- paste0("col_", which(bad))

  nm <- gsub("[^A-Za-z0-9_]+", "_", nm)
  nm <- gsub("_+", "_", nm)
  nm <- gsub("^_|_$", "", nm)
  nm <- make.unique(nm, sep = "_")

  names(df) <- nm
  df
}

validate_schema <- function(df, required_cols, context = "data frame") {
  missing <- setdiff(required_cols, names(df))
  if (length(missing) > 0) {
    stop(
      "Missing required columns in ", context, ": ", paste(missing, collapse = ", "), "\n",
      "Columns present: ", paste(names(df), collapse = ", ")
    )
  }
  invisible(TRUE)
}

safeWriteTable <- function(wb, sheet, x, ...) {
  x <- fix_colnames_for_excel(x)
  tryCatch(
    openxlsx::writeDataTable(wb, sheet, x, ...),
    error = function(e) stop("writeDataTable failed on sheet: ", sheet, "\n", conditionMessage(e))
  )
}

preflight_table <- function(df, table_name) {
  nm <- names(df)
  if (is.null(nm)) stop("Preflight failed [", table_name, "]: NULL column names")
  if (anyDuplicated(nm) > 0) stop("Preflight failed [", table_name, "]: duplicated column names")
  if (any(is.na(nm) | nm == "")) stop("Preflight failed [", table_name, "]: NA/empty column names")

  data.frame(
    table = table_name,
    n_rows = nrow(df),
    n_cols = ncol(df),
    status = "OK",
    stringsAsFactors = FALSE
  )
}

GROUP_ORDER <- c(
  "PE", "PE-Cl", "PP", "PS", "PA", "PVC",
  "PET", "PEN",
  "PC", "PMMA", "PAN", "PLA", "PVA", "EVA",
  "POM", "POB", "IR", "PTFE", "PU",
  "ABS", "PES", "PPS", "CA", "CAB", "PAM/ACR"
)

FERET_BINS  <- c("< 100 µm", "100 – 200 µm", "200 – 300 µm", "> 300 µm")

standardize_group_code <- function(x) {
  x <- stringr::str_trim(as.character(x))
  x_up <- toupper(x)

  # If already a valid abbreviation, keep it
  ifelse(
    x_up %in% GROUP_ORDER,
    x_up,
    dplyr::case_when(

      # --- Chlorinated polyethylene (must be before generic PE)
      x_up %in% c("PE-CL", "PECL", "CPE") ~ "PE-Cl",
      stringr::str_detect(x, stringr::regex("polyethylene\\s*chlor", ignore_case = TRUE)) ~ "PE-Cl",
      stringr::str_detect(x, stringr::regex("chlorinated\\s*polyethylene", ignore_case = TRUE)) ~ "PE-Cl",

      # --- Polyethylene naphthalate (must be before generic PE)
      x_up == "PEN" ~ "PEN",
      stringr::str_detect(x, stringr::regex("polyethylene\\s*(napht|naphth)", ignore_case = TRUE)) ~ "PEN",
      stringr::str_detect(x, stringr::regex("\\bPEN\\b", ignore_case = TRUE)) ~ "PEN",

      # --- PET (must be before generic PE)
      x_up == "PET" ~ "PET",
      stringr::str_detect(x, stringr::regex("poly(ethylene\\s*)?tereph", ignore_case = TRUE)) ~ "PET",

      # --- Generic PE LAST so it doesn't capture PEN / PET / PE-Cl
      stringr::str_starts(x, stringr::regex("PolyEthylen", ignore_case = TRUE)) ~ "PE",
      stringr::str_detect(x, stringr::regex("polyethy", ignore_case = TRUE))    ~ "PE",

      # --- PP
      stringr::str_starts(x, stringr::regex("Polypro", ignore_case = TRUE))     ~ "PP",
      stringr::str_detect(x, stringr::regex("polyprop", ignore_case = TRUE))    ~ "PP",

      # --- PS
      x_up == "PS"  ~ "PS",
      stringr::str_detect(x, stringr::regex("polystyr", ignore_case = TRUE))    ~ "PS",

      # --- PA
      x_up == "PA"  ~ "PA",
      stringr::str_detect(x, stringr::regex("polyamid", ignore_case = TRUE))    ~ "PA",

      # --- PVC
      x_up == "PVC" ~ "PVC",
      stringr::str_detect(x, stringr::regex("polyvinyl\\s*chlor", ignore_case = TRUE)) ~ "PVC",

      # --- PMMA
      x_up == "PMMA" ~ "PMMA",
      stringr::str_detect(x, stringr::regex("polymethyl\\s*methacry", ignore_case = TRUE)) ~ "PMMA",

      # --- PAN
      x_up == "PAN" ~ "PAN",
      stringr::str_starts(x, stringr::regex("Polyacrilonitrile", ignore_case = TRUE)) ~ "PAN",
      stringr::str_detect(x, stringr::regex("acrylonitrile", ignore_case = TRUE)) ~ "PAN",

      # --- PLA
      x_up == "PLA" ~ "PLA",
      stringr::str_detect(x, stringr::regex("polylact", ignore_case = TRUE)) ~ "PLA",

      # --- PVA
      x_up == "PVA" ~ "PVA",
      stringr::str_detect(x, stringr::regex("polyvinyl\\s*alcohol", ignore_case = TRUE)) ~ "PVA",

      # --- EVA
      x_up == "EVA" ~ "EVA",
      stringr::str_detect(x, stringr::regex("ethylene\\s*vinyl\\s*acet", ignore_case = TRUE)) ~ "EVA",

      # --- POM
      x_up == "POM" ~ "POM",
      stringr::str_detect(x, stringr::regex("polyoxymeth", ignore_case = TRUE)) ~ "POM",

      # --- POB
      x_up == "POB" ~ "POB",

      # --- IR (polyisoprene)
      x_up == "IR"  ~ "IR",
      stringr::str_detect(x, stringr::regex("polyisoprene|isoprene", ignore_case = TRUE)) ~ "IR",

      # --- PTFE
      x_up == "PTFE" ~ "PTFE",
      stringr::str_detect(x, stringr::regex("polytetrafluoro|teflon", ignore_case = TRUE)) ~ "PTFE",

      # --- PU
      x_up == "PU"   ~ "PU",
      stringr::str_detect(x, stringr::regex("polyure", ignore_case = TRUE)) ~ "PU",

      # --- PC
      x_up == "PC" ~ "PC",
      stringr::str_detect(x, stringr::regex("polycarbonate", ignore_case = TRUE)) ~ "PC",

      # --- ABS
      x_up == "ABS" ~ "ABS",
      stringr::str_detect(x, stringr::regex("acrylonitrile\\s*butadiene\\s*styrene|\\bABS\\b", ignore_case = TRUE)) ~ "ABS",

      # --- PES (polyethersulfone / polyether sulphone)
      x_up == "PES" ~ "PES",
      stringr::str_detect(x, stringr::regex("polyether\\s*sul(phone|fone)|polyethersulfone|\\bPES\\b", ignore_case = TRUE)) ~ "PES",

      # --- PPS (polyphenylene sulfide)
      x_up == "PPS" ~ "PPS",
      stringr::str_detect(x, stringr::regex("polyphenylene\\s*sul(phide|fide)|\\bPPS\\b", ignore_case = TRUE)) ~ "PPS",

      # --- CA (cellulose acetate)
      x_up == "CA" ~ "CA",
      stringr::str_detect(x, stringr::regex("cellulose\\s*acetate|\\bCA\\b", ignore_case = TRUE)) ~ "CA",

      # --- CAB (cellulose acetate butyrate)
      x_up == "CAB" ~ "CAB",
      stringr::str_detect(x, stringr::regex("cellulose\\s*acetate\\s*butyrate|\\bCAB\\b", ignore_case = TRUE)) ~ "CAB",

      # --- PAM/ACR (polyacrylamide/acrylate)
      x_up == "PAM/ACR" ~ "PAM/ACR",
      stringr::str_detect(x, stringr::regex("polyacrylamide\\s*/\\s*acrylate|polyacrylamide\\s*acrylate", ignore_case = TRUE)) ~ "PAM/ACR",

      TRUE ~ NA_character_
    )
  )
}

add_total_row <- function(df, id_col = "source_file", label = "Total") {
  if (!id_col %in% names(df)) stop("add_total_row: id_col not found in df")

  num_cols <- names(df)[vapply(df, is.numeric, logical(1))]
  num_cols <- setdiff(num_cols, id_col)

  totals <- df %>%
    dplyr::summarise(dplyr::across(dplyr::all_of(num_cols), ~ sum(.x, na.rm = TRUE)))

  totals[[id_col]] <- label

  totals <- totals[, names(df), drop = FALSE]
  dplyr::bind_rows(df, totals)
}

add_mean_row <- function(df, id_col = "source_file", label = "Mean") {
  if (!id_col %in% names(df)) stop("add_mean_row: id_col not found in df")

  num_cols <- names(df)[vapply(df, is.numeric, logical(1))]
  num_cols <- setdiff(num_cols, id_col)

  means <- df %>%
    dplyr::summarise(dplyr::across(dplyr::all_of(num_cols), ~ mean(.x, na.rm = TRUE)))

  means[[id_col]] <- label
  means <- means[, names(df), drop = FALSE]
  dplyr::bind_rows(df, means)
}

total_row_style <- createStyle(
  textDecoration = "bold",
  border = "top",
  borderStyle = "thick"
)

bin_feret <- function(x_um) {
  cut(
    x_um,
    breaks = c(-Inf, 100, 200, 300, Inf),
    labels = FERET_BINS,
    right = FALSE
  )
}

ensure_cols_in_order <- function(df_wide, cols_order, id_col = "source_file") {
  # Ensure all expected columns exist (fill missing with 0),
  # BUT keep any additional columns that appear in the data (do not drop them).
  missing <- setdiff(cols_order, names(df_wide))
  if (length(missing) > 0) {
    for (m in missing) df_wide[[m]] <- 0
  }

  # Put known columns first (in the requested order), then append any extras.
  known <- intersect(cols_order, names(df_wide))
  extras <- setdiff(names(df_wide), c(id_col, cols_order))
  extras <- sort(extras)

  df_wide[, c(id_col, known, extras), drop = FALSE]
}



parse_num_any <- function(x) {
  x_chr <- as.character(x)
  # Handle European decimal comma (e.g., "0,82")
  use_comma <- any(grepl(",", x_chr, fixed = TRUE), na.rm = TRUE) && !any(grepl("\\.", x_chr), na.rm = TRUE)
  if (use_comma) {
    readr::parse_number(x_chr, locale = readr::locale(decimal_mark = ","))
  } else {
    readr::parse_number(x_chr)
  }
}


# ---------------------------
# FT-IR workflow (existing)
# ---------------------------
read_one_csv_ftir <- function(path) {
  
  try_read <- function(enc) {
    readr::read_csv(
      path,
      show_col_types = FALSE,
      locale = readr::locale(encoding = enc)
    )
  }
  
  # NOTE: drop UTF-8-BOM (readr/ICU usually handles BOM under UTF-8)
  candidates <- c("UTF-8", "UTF-16LE", "UTF-16BE", "Windows-1252", "Latin1")
  
  dfs <- list()
  scores <- numeric(length(candidates))
  
  for (i in seq_along(candidates)) {
    enc <- candidates[[i]]
    df_i <- tryCatch(try_read(enc), error = function(e) NULL)
    if (is.null(df_i)) {
      scores[i] <- -Inf
      next
    }
    
    # --- Long-term fix: normalize column names once
    cn <- names(df_i)
    cn_utf8 <- iconv(cn, from = "", to = "UTF-8", sub = "byte")
    names(df_i) <- cn_utf8
    cn <- cn_utf8
    
    suppressWarnings({
      has_micro <- any(grepl("µ", cn, fixed = TRUE))
      has_sq    <- any(grepl("²", cn, fixed = TRUE)) || any(grepl("³", cn, fixed = TRUE))
      has_repl  <- any(grepl("\uFFFD", cn))
    })
    
    scores[i] <- (has_micro * 10) + (has_sq * 3) - (has_repl * 100)
    dfs[[i]] <- df_i
  }
  
  best_i <- which.max(scores)
  df <- dfs[[best_i]]
  
  if (is.null(df) || !is.finite(scores[best_i]) || scores[best_i] < 0) {
    stop(
      "Could not read CSV headers without replacement characters. ",
      "This usually means the file encoding is unusual or the file is not plain CSV.\n",
      "Try opening the CSV in a text editor and checking encoding (UTF-16LE is common for instruments)."
    )
  }
  
  cn <- names(df)
  required_exact <- c("Identifier", "Group", "Feret min [µm]", "Mass [ng]")
  missing <- setdiff(required_exact, cn)
  if (length(missing) > 0) {
    stop(
      "Missing required columns in file: ", basename(path), "\n",
      "Missing headers: ", paste(missing, collapse = ", "), "\n",
      "Columns present: ", paste(cn, collapse = ", ")
    )
  }
  
  df <- df %>%
    dplyr::filter(stringr::str_starts(as.character(`Identifier`), "MP")) %>%
    dplyr::mutate(
      group_raw    = stringr::str_trim(as.character(`Group`)),
      group_mapped = standardize_group_code(group_raw),
      # Keep ALL unique materials: only abbreviate when we have a mapped code
      group_code   = dplyr::if_else(is.na(group_mapped) | group_mapped == "", group_raw, group_mapped),
      feret_um     = as.numeric(`Feret min [µm]`),
      mass_ng      = as.numeric(`Mass [ng]`),
      source_file  = basename(path)
    )
  
  df
}

make_pivots_ftir <- function(df_one) {
  src <- df_one$source_file[[1]]

  pA <- df_one %>%
    dplyr::filter(!is.na(group_code)) %>%
    dplyr::group_by(group_code) %>%
    dplyr::summarise(value = dplyr::n(), .groups = "drop") %>%
    tidyr::pivot_wider(
      names_from = group_code,
      values_from = value,
      values_fill = 0
    ) %>%
    dplyr::mutate(source_file = src, .before = 1)

  pA <- ensure_cols_in_order(pA, GROUP_ORDER, id_col = "source_file")

  pB <- df_one %>%
    dplyr::mutate(
      feret_bin = bin_feret(feret_um),
      feret_bin = factor(feret_bin, levels = FERET_BINS)
    ) %>%
    dplyr::group_by(feret_bin) %>%
    dplyr::summarise(value = dplyr::n(), .groups = "drop") %>%
    tidyr::complete(
      feret_bin = factor(FERET_BINS, levels = FERET_BINS),
      fill = list(value = 0)
    ) %>%
    tidyr::pivot_wider(
      names_from = feret_bin,
      values_from = value,
      values_fill = 0
    ) %>%
    dplyr::mutate(source_file = src, .before = 1)

  pB <- ensure_cols_in_order(pB, FERET_BINS, id_col = "source_file")

  pC <- df_one %>%
    dplyr::filter(!is.na(group_code)) %>%
    dplyr::group_by(group_code) %>%
    dplyr::summarise(value = round(sum(mass_ng, na.rm = TRUE), digits = 2), .groups = "drop") %>%
    tidyr::pivot_wider(
      names_from = group_code,
      values_from = value,
      values_fill = 0
    ) %>%
    dplyr::mutate(source_file = src, .before = 1)

  pC <- ensure_cols_in_order(pC, GROUP_ORDER, id_col = "source_file")

  list(
    count_by_group = pA,
    count_by_feret = pB,
    mass_by_group  = pC
  )
}

write_output_workbook_ftir <- function(input_paths, output_path) {
  dfs <- lapply(input_paths, read_one_csv_ftir)
  long_df <- dplyr::bind_rows(dfs)

  long_df <- long_df %>% dplyr::select(-c(group_raw, group_mapped, group_code, feret_um, mass_ng))

  all_A <- list()
  all_B <- list()
  all_C <- list()

  wb <- createWorkbook()

  addWorksheet(wb, "Long_Table")
  safeWriteTable(wb, "Long_Table", fix_colnames_for_excel(long_df))
  setColWidths(wb, "Long_Table", cols = 1:min(30, ncol(long_df)), widths = 10)

  for (i in seq_along(dfs)) {
    df_one <- dfs[[i]]
    file_name <- basename(input_paths[[i]])
    sheet_name <- sanitize_sheet_name(tools::file_path_sans_ext(file_name))

    original <- sheet_name
    k <- 2
    while (sheet_name %in% names(wb)) {
      sheet_name <- sanitize_sheet_name(paste0(original, "_", k))
      k <- k + 1
    }

    piv <- make_pivots_ftir(df_one)

    all_A[[i]] <- piv$count_by_group
    all_B[[i]] <- piv$count_by_feret
    all_C[[i]] <- piv$mass_by_group

    addWorksheet(wb, sheet_name)

    r <- 1

    writeData(wb, sheet_name, "A) Counts per plastic type", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, fix_colnames_for_excel(piv$count_by_group), startRow = r, startCol = 1)
    r <- r + nrow(piv$count_by_group) + 3

    writeData(wb, sheet_name, "B) Counts per size class", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, fix_colnames_for_excel(piv$count_by_feret), startRow = r, startCol = 1)
    r <- r + nrow(piv$count_by_feret) + 3

    writeData(wb, sheet_name, "C) Total Mass [ng] per plastic type", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, fix_colnames_for_excel(piv$mass_by_group), startRow = r, startCol = 1)

    setColWidths(wb, sheet_name, cols = 1:max(ncol(piv$count_by_group), ncol(piv$count_by_feret), ncol(piv$mass_by_group)), widths = 10)
  }

  summary_A <- dplyr::bind_rows(all_A)
  summary_B <- dplyr::bind_rows(all_B)
  summary_C <- dplyr::bind_rows(all_C)

  summary_A <- add_total_row(summary_A, id_col = "source_file", label = "Total")
  summary_B <- add_total_row(summary_B, id_col = "source_file", label = "Total")
  summary_C <- add_total_row(summary_C, id_col = "source_file", label = "Total")

  addWorksheet(wb, "Total")
  rr <- 1

  writeData(wb, "Total", "A) Counts per plastic type", startRow = rr, startCol = 1)
  rr <- rr + 1
  start_A <- rr
  safeWriteTable(wb, "Total", fix_colnames_for_excel(summary_A), startRow = rr, startCol = 1)
  row_A_total <- start_A + which(summary_A$source_file == "Total")
  addStyle(wb, "Total", total_row_style, rows = row_A_total, cols = 1:ncol(summary_A), gridExpand = TRUE, stack = TRUE)
  rr <- rr + nrow(summary_A) + 3

  writeData(wb, "Total", "B) Counts per size class", startRow = rr, startCol = 1)
  rr <- rr + 1
  start_B <- rr
  safeWriteTable(wb, "Total", fix_colnames_for_excel(summary_B), startRow = rr, startCol = 1)
  row_B_total <- start_B + which(summary_B$source_file == "Total")
  addStyle(wb, "Total", total_row_style, rows = row_B_total, cols = 1:ncol(summary_B), gridExpand = TRUE, stack = TRUE)
  rr <- rr + nrow(summary_B) + 3

  writeData(wb, "Total", "C) Total mass [ng] per plastic type", startRow = rr, startCol = 1)
  rr <- rr + 1
  start_C <- rr
  safeWriteTable(wb, "Total", fix_colnames_for_excel(summary_C), startRow = rr, startCol = 1)
  row_C_total <- start_C + which(summary_C$source_file == "Total")
  addStyle(wb, "Total", total_row_style, rows = row_C_total, cols = 1:ncol(summary_C), gridExpand = TRUE, stack = TRUE)

  setColWidths(wb, "Total", cols = 1:max(ncol(summary_A), ncol(summary_B), ncol(summary_C)), widths = 10)

  total_A <- summary_A %>% dplyr::filter(source_file == "Total")
  total_B <- summary_B %>% dplyr::filter(source_file == "Total")
  total_C <- summary_C %>% dplyr::filter(source_file == "Total")

  addWorksheet(wb, "Summary")
  r2 <- 1
  writeData(wb, "Summary", "A) Total counts per plastic type", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", fix_colnames_for_excel(total_A), startRow = r2, startCol = 1)
  r2 <- r2 + nrow(total_A) + 3

  writeData(wb, "Summary", "B) Total counts per size class", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", fix_colnames_for_excel(total_B), startRow = r2, startCol = 1)
  r2 <- r2 + nrow(total_B) + 3

  writeData(wb, "Summary", "C) Total mass [ng] per plastic type", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", fix_colnames_for_excel(total_C), startRow = r2, startCol = 1)

  setColWidths(wb, "Summary", cols = 1:max(ncol(total_A), ncol(total_B), ncol(total_C)), widths = 10)

  saveWorkbook(wb, output_path, overwrite = TRUE)
  output_path
}

# ---------------------------
# Raman workflow (new)
# ---------------------------

# Read Raman CSV with robust encoding attempts and flexible header matching.
# Requirements (conceptual):
# - Material column (plastic type)
# - Feret max [um] (size metric)
# - HQI (quality index) for filtering
read_one_csv_raman <- function(path, hqi_cutoff = 70) {

  first_line <- tryCatch(readLines(path, n = 1, warn = FALSE), error = function(e) "")
  delim <- if (length(first_line) && grepl(";", first_line, fixed = TRUE) && !grepl(",", first_line, fixed = TRUE)) {
    ";"
  } else if (length(first_line) && grepl(";", first_line, fixed = TRUE) && grepl(",", first_line, fixed = TRUE)) {
    if (stringr::str_count(first_line, fixed(";")) >= stringr::str_count(first_line, fixed(","))) ";" else ","
  } else {
    ","
  }

  raw_df <- readr::read_delim(
    file = path,
    delim = delim,
    show_col_types = FALSE,
    locale = readr::locale(encoding = "UTF-8"),
    progress = FALSE,
    trim_ws = TRUE
  )

  names(raw_df) <- janitor::clean_names(names(raw_df), replace = c("µ" = "u", "μ" = "u"))

  pick_col <- function(candidates, patterns) {
    exact_hit <- intersect(candidates, names(raw_df))
    if (length(exact_hit) > 0) return(exact_hit[[1]])

    for (pat in patterns) {
      idx <- which(grepl(pat, names(raw_df)))[1]
      if (!is.na(idx)) return(names(raw_df)[[idx]])
    }
    NA_character_
  }

  col_particle <- pick_col(
    c("particle_name", "particle", "identifier", "id", "name"),
    c("particle", "identifier", "^id$", "name")
  )
  col_material <- pick_col(
    c("material", "polymer", "plastic", "group"),
    c("(^|_)material($|_)", "(^|_)polymer($|_)", "(^|_)plastic($|_)", "(^|_)group($|_)")
  )
  col_hqi <- pick_col(
    c("hqi", "hq_index", "quality_index", "quality"),
    c("(^|_)hqi($|_)", "hq_index", "quality.*index", "quality")
  )
  col_feret <- pick_col(
    c("feret_max_um", "feret_max", "max_feret", "feret"),
    c("feret.*max", "max.*feret", "feret_max")
  )

  detected_required <- c(
    material_raw = col_material,
    hqi_value_raw = col_hqi,
    feret_max_um_raw = col_feret
  )
  if (any(is.na(detected_required))) {
    missing <- names(detected_required)[is.na(detected_required)]
    stop(
      "Raman CSV is missing required columns in file: ", basename(path), "\n",
      "Missing canonical fields: ", paste(missing, collapse = ", "), "\n",
      "Columns present: ", paste(names(raw_df), collapse = ", ")
    )
  }

  df <- raw_df %>%
    dplyr::mutate(
      particle_name = if (!is.na(col_particle)) as.character(.data[[col_particle]]) else NA_character_,
      material_raw = as.character(.data[[col_material]]),
      hqi_value_raw = readr::parse_number(as.character(.data[[col_hqi]])),
      feret_max_um_raw = readr::parse_number(as.character(.data[[col_feret]])),
      source_file = basename(path)
    )

  validate_schema(
    df,
    required_cols = c("particle_name", "material_raw", "hqi_value_raw", "feret_max_um_raw", "source_file"),
    context = paste0("Raman file ", basename(path))
  )

  df <- df %>%
    dplyr::mutate(
      material_raw = stringr::str_trim(as.character(material_raw)),
      material_mapped = standardize_group_code(material_raw),
      group_code = dplyr::if_else(is.na(material_mapped) | material_mapped == "", "OTHER", material_mapped),
      hqi_value = dplyr::if_else(!is.na(hqi_value_raw) & hqi_value_raw <= 1, hqi_value_raw * 100, hqi_value_raw),
      feret_max_um = feret_max_um_raw
    ) %>%
    dplyr::filter(!is.na(hqi_value) & hqi_value >= hqi_cutoff)

  df
}


make_pivots_raman <- function(df_one, src_override = NULL) {

  src <- src_override
  if (is.null(src)) {
    if ("source_file" %in% names(df_one) && nrow(df_one) > 0) {
      src <- df_one$source_file[[1]]
    } else {
      src <- "Unknown"
    }
  }

  group_levels <- c(GROUP_ORDER, "OTHER")

  count_by_group <- df_one %>%
    dplyr::mutate(group_code = dplyr::if_else(is.na(group_code) | group_code == "", "OTHER", group_code)) %>%
    dplyr::count(group_code, name = "value") %>%
    dplyr::mutate(group_code = factor(group_code, levels = group_levels)) %>%
    tidyr::complete(group_code = factor(group_levels, levels = group_levels), fill = list(value = 0)) %>%
    tidyr::pivot_wider(names_from = group_code, values_from = value, values_fill = 0) %>%
    dplyr::mutate(source_file = src, .before = 1)
  count_by_group <- ensure_cols_in_order(count_by_group, group_levels, id_col = "source_file")

  count_by_feret <- df_one %>%
    dplyr::mutate(
      feret_bin = bin_feret(feret_max_um),
      feret_bin = factor(feret_bin, levels = FERET_BINS)
    ) %>%
    dplyr::count(feret_bin, name = "value") %>%
    tidyr::complete(feret_bin = factor(FERET_BINS, levels = FERET_BINS), fill = list(value = 0)) %>%
    tidyr::pivot_wider(names_from = feret_bin, values_from = value, values_fill = 0) %>%
    dplyr::mutate(source_file = src, .before = 1)
  count_by_feret <- ensure_cols_in_order(count_by_feret, FERET_BINS, id_col = "source_file")

  unknown_materials_long <- df_one %>%
    dplyr::filter(is.na(material_mapped) | material_mapped == "") %>%
    dplyr::select(source_file, particle_name, material_raw, hqi_value, feret_max_um)

  list(
    count_by_group = count_by_group,
    count_by_feret = count_by_feret,
    unknown_materials_long = unknown_materials_long
  )
}


write_output_workbook_raman <- function(input_paths, output_path, hqi_cutoff = 70) {
  dfs <- lapply(input_paths, read_one_csv_raman, hqi_cutoff = hqi_cutoff)
  long_df <- dplyr::bind_rows(dfs)
  long_df <- fix_colnames_for_excel(long_df)

  all_A <- list()
  all_B <- list()
  all_unknown <- list()

  wb <- createWorkbook()

  addWorksheet(wb, "Long_Table")
  safeWriteTable(wb, "Long_Table", long_df)
  setColWidths(wb, "Long_Table", cols = 1:ncol(long_df), widths = 10)

  for (i in seq_along(dfs)) {
    df_one <- dfs[[i]]
    file_name <- basename(input_paths[[i]])
    sheet_name <- sanitize_sheet_name(tools::file_path_sans_ext(file_name))

    original <- sheet_name
    k <- 2
    while (sheet_name %in% names(wb)) {
      sheet_name <- sanitize_sheet_name(paste0(original, "_", k))
      k <- k + 1
    }

    piv <- make_pivots_raman(df_one, src_override = basename(input_paths[[i]]))
    piv$count_by_group <- fix_colnames_for_excel(piv$count_by_group)
    piv$count_by_feret <- fix_colnames_for_excel(piv$count_by_feret)
    piv$unknown_materials_long <- fix_colnames_for_excel(piv$unknown_materials_long)

    all_A[[i]] <- piv$count_by_group
    all_B[[i]] <- piv$count_by_feret
    all_unknown[[i]] <- piv$unknown_materials_long

    addWorksheet(wb, sheet_name)

    r <- 1
    writeData(wb, sheet_name, "A) Counts per plastic type", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, piv$count_by_group, startRow = r, startCol = 1)
    r <- r + nrow(piv$count_by_group) + 3

    writeData(wb, sheet_name, "B) Counts per size class", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, piv$count_by_feret, startRow = r, startCol = 1)

    setColWidths(wb, sheet_name, cols = 1:max(ncol(piv$count_by_group), ncol(piv$count_by_feret)), widths = 10)
  }

  summary_A <- dplyr::bind_rows(all_A)
  summary_B <- dplyr::bind_rows(all_B)
  unknown_all <- dplyr::bind_rows(all_unknown)

  summary_A <- fix_colnames_for_excel(add_total_row(summary_A, id_col = "source_file", label = "Total"))
  summary_B <- fix_colnames_for_excel(add_total_row(summary_B, id_col = "source_file", label = "Total"))
  unknown_all <- fix_colnames_for_excel(unknown_all)

  addWorksheet(wb, "Total")
  rr <- 1

  writeData(wb, "Total", "A) Counts per plastic type", startRow = rr, startCol = 1)
  rr <- rr + 1
  start_A <- rr
  safeWriteTable(wb, "Total", summary_A, startRow = rr, startCol = 1)
  row_A_total <- start_A + which(summary_A$source_file == "Total")
  addStyle(wb, "Total", total_row_style, rows = row_A_total, cols = 1:ncol(summary_A), gridExpand = TRUE, stack = TRUE)
  rr <- rr + nrow(summary_A) + 3

  writeData(wb, "Total", "B) Counts per size class", startRow = rr, startCol = 1)
  rr <- rr + 1
  start_B <- rr
  safeWriteTable(wb, "Total", summary_B, startRow = rr, startCol = 1)
  row_B_total <- start_B + which(summary_B$source_file == "Total")
  addStyle(wb, "Total", total_row_style, rows = row_B_total, cols = 1:ncol(summary_B), gridExpand = TRUE, stack = TRUE)
  setColWidths(wb, "Total", cols = 1:max(ncol(summary_A), ncol(summary_B)), widths = 10)

  mean_A <- summary_A %>% dplyr::filter(source_file != "Total")
  mean_B <- summary_B %>% dplyr::filter(source_file != "Total")

  mean_A <- fix_colnames_for_excel(add_mean_row(mean_A, id_col = "source_file", label = "Mean"))
  mean_B <- fix_colnames_for_excel(add_mean_row(mean_B, id_col = "source_file", label = "Mean"))

  mean_A <- fix_colnames_for_excel(dplyr::bind_rows(mean_A, summary_A %>% dplyr::filter(source_file == "Total")))
  mean_B <- fix_colnames_for_excel(dplyr::bind_rows(mean_B, summary_B %>% dplyr::filter(source_file == "Total")))

  addWorksheet(wb, "Summary")
  r2 <- 1
  writeData(wb, "Summary", "A) Mean counts per plastic type (across files)", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", mean_A, startRow = r2, startCol = 1)
  r2 <- r2 + nrow(mean_A) + 3

  writeData(wb, "Summary", "B) Mean counts per size class (across files)", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", mean_B, startRow = r2, startCol = 1)
  setColWidths(wb, "Summary", cols = 1:max(ncol(mean_A), ncol(mean_B)), widths = 10)

  addWorksheet(wb, "Unknown_Materials")
  safeWriteTable(wb, "Unknown_Materials", unknown_all)
  setColWidths(wb, "Unknown_Materials", cols = 1:ncol(unknown_all), widths = 10)

  preflight <- dplyr::bind_rows(
    preflight_table(long_df, "Long_Table"),
    preflight_table(summary_A, "Total_A"),
    preflight_table(summary_B, "Total_B"),
    preflight_table(mean_A, "Summary_A"),
    preflight_table(mean_B, "Summary_B"),
    preflight_table(unknown_all, "Unknown_Materials")
  )
  print(preflight)

  addWorksheet(wb, "Log")
  safeWriteTable(wb, "Log", preflight)
  setColWidths(wb, "Log", cols = 1:ncol(preflight), widths = 10)

  saveWorkbook(wb, output_path, overwrite = TRUE)
  output_path
}


# ---------------------------
# Shiny app
# ---------------------------

ui <- fluidPage(
  titlePanel("CSV → Processed Workbook"),

  sidebarLayout(
    sidebarPanel(
      selectInput(
        inputId = "workflow",
        label = "Data type",
        choices = c("FT-IR" = "ftir", "Raman" = "raman"),
        selected = "ftir"
      ),

      conditionalPanel(
        condition = "input.workflow == 'raman'",
        numericInput(
          inputId = "hqi_cutoff",
          label = "HQI cutoff (%)",
          value = 70,
          min = 0,
          max = 100,
          step = 1
        ),
        tags$small("Rows with HQI below the cutoff are discarded before processing.")
      ),

      br(),

      shinyFilesButton(
        id = "files",
        label = "Choose CSV files (multi-select)",
        title = "Select CSV files to process",
        multiple = TRUE
      ),
      tags$small("Output will be saved next to the selected files as processed_data.xlsx"),
      br(), br(),
      actionButton("process", "Process & Save", class = "btn-primary"),
      br(), br(),
      strong("Status:"),
      verbatimTextOutput("status"),
      br(),
      strong("Output path:"),
      verbatimTextOutput("out_path")
    ),
    mainPanel(
      h4("Selected files"),
      tableOutput("file_table")
    )
  )
)

server <- function(input, output, session) {

  roots <- c("Home" = normalizePath("~", winslash = "/", mustWork = TRUE))
  if (.Platform$OS.type == "windows") roots <- c(roots, "C:" = "C:/")

  shinyFileChoose(input, "files", roots = roots, filetypes = c("csv"))

  selected_paths <- reactive({
    req(input$files)
    paths <- shinyFiles::parseFilePaths(roots, input$files)
    as.character(paths$datapath)
  })

  output$file_table <- renderTable({
    req(selected_paths())
    data.frame(
      File = basename(selected_paths()),
      Folder = dirname(selected_paths()),
      stringsAsFactors = FALSE
    )
  })

  status_val <- reactiveVal("Select CSV files, choose FT-IR or Raman, then click Process & Save.")
  out_path_val <- reactiveVal("")

  output$status <- renderText(status_val())
  output$out_path <- renderText(out_path_val())

  observeEvent(input$process, {
    req(selected_paths())
    paths <- selected_paths()

    out_dir <- dirname(paths[[1]])
    out_file <- file.path(
      out_dir,
      paste0("processed_data_", input$workflow, "_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")
    )
    
    status_val("Processing...")
    out_path_val("")

    tryCatch({
      if (input$workflow == "ftir") {
        res <- write_output_workbook_ftir(paths, out_file)
      } else {
        res <- write_output_workbook_raman(paths, out_file, hqi_cutoff = input$hqi_cutoff)
      }
      status_val("Done.")
      out_path_val(res)
    }, error = function(e) {
      status_val(paste("ERROR:", conditionMessage(e)))
      out_path_val("")
    })
  })
}

shinyApp(ui, server)
