required_pkgs <- c(
  "shiny",
  "shinyFiles",
  "tidyr",
  "dplyr",
  "openxlsx",
  "stringr",
  "readr",
  "janitor",
  "readxl"
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

      # --- Non-polymer guard: obvious non-plastic materials → NA
      # Fires when a non-polymer keyword is present and no polymer cue overrides it.
      stringr::str_detect(x_up, "\\bOIL\\b|\\bGUM\\b|\\bRESIN\\b|\\bOXIDE\\b|\\bPHOSPHATE\\b|\\bSTEARATE\\b|\\bSALT\\b|\\bMICA\\b|\\bQUARTZ\\b|\\bTITANIUM\\b|\\bPIGMENT\\b|\\bPARAFFIN\\b") &
        !stringr::str_detect(x_up, "\\bPOLY|\\bEVOH\\b|\\bPVAC\\b|\\bEPDM\\b|\\bNYLON\\b") ~ NA_character_,

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

      # --- EVOH → PVA (ethylene-vinyl alcohol copolymer; before generic PE)
      x_up == "EVOH" ~ "PVA",
      stringr::str_detect(x, stringr::regex("\\bEVOH\\b", ignore_case = TRUE)) ~ "PVA",
      stringr::str_detect(x, stringr::regex("ethylene\\s*vinyl\\s*alcohol", ignore_case = TRUE)) ~ "PVA",
      stringr::str_detect(x, stringr::regex("vinyl\\s*alcohol\\s*copolymer", ignore_case = TRUE)) ~ "PVA",

      # --- EPDM / ethylene-propylene-diene rubber → IR (elastomer bucket; before generic PE/PP)
      x_up %in% c("EPDM", "EPR") ~ "IR",
      stringr::str_detect(x, stringr::regex("\\bEPDM\\b", ignore_case = TRUE)) ~ "IR",
      stringr::str_detect(x, stringr::regex("ethylene\\s*[:/-]\\s*propylene\\s*[:/-]\\s*diene", ignore_case = TRUE)) ~ "IR",
      stringr::str_detect(x, stringr::regex("ethylene\\s*propylene\\s*diene", ignore_case = TRUE)) ~ "IR",

      # --- Masterbatch / Film: classify only when an explicit polymer is identifiable
      # These rules must precede the broad poly* patterns to prevent mis-classification.
      stringr::str_detect(x_up, "MASTERBATCH|\\bFILM\\b") &
        stringr::str_detect(x, stringr::regex("ethylene.*vinyl.*acet|poly\\[ethylene.*vinyl\\s*acetate", ignore_case = TRUE)) ~ "EVA",
      stringr::str_detect(x_up, "MASTERBATCH|\\bFILM\\b") &
        stringr::str_detect(x, stringr::regex("polyethy", ignore_case = TRUE)) ~ "PE",
      stringr::str_detect(x_up, "MASTERBATCH|\\bFILM\\b") &
        stringr::str_detect(x, stringr::regex("polyprop", ignore_case = TRUE)) ~ "PP",
      stringr::str_detect(x_up, "MASTERBATCH|\\bFILM\\b") &
        stringr::str_detect(x, stringr::regex("polystyr", ignore_case = TRUE)) ~ "PS",

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
      stringr::str_detect(x, stringr::regex("nylon\\s*[-_]?\\s*6[,\\.]?6|\\bnylon\\b", ignore_case = TRUE)) ~ "PA",

      # --- PET / copolyesters
      x_up == "PETG" ~ "PET",
      stringr::str_detect(x, stringr::regex("\\bpetg\\b", ignore_case = TRUE)) ~ "PET",

      # --- EVA / related notation
      stringr::str_detect(x, stringr::regex("ethylene\\s*[/:-]\\s*vinyl\\s*acet", ignore_case = TRUE)) ~ "EVA",

      # --- PE copolymer shorthand
      x_up == "EMAA" ~ "PE",
      stringr::str_detect(x, stringr::regex("ethylene\\s*methacrylic\\s*acid", ignore_case = TRUE)) ~ "PE",

      # --- Vinyl chloride copolymers
      stringr::str_detect(x, stringr::regex("vinyl\\s*chloride\\s*[/:-]\\s*vinyl\\s*acet", ignore_case = TRUE)) ~ "PVC",

      # --- Methacrylate polymers
      stringr::str_detect(x, stringr::regex("poly\\s*\\(\\s*n[- ]?butyl\\s*methacrylate\\s*\\)", ignore_case = TRUE)) ~ "PMMA",

      # --- Polyolefin variants
      stringr::str_detect(x, stringr::regex("poly\\s*\\(\\s*4[- ]?methyl[- ]?1[- ]?pentene\\s*\\)", ignore_case = TRUE)) ~ "PP",
      stringr::str_detect(x, stringr::regex("polypropylene\\s*isotactic\\s*chlorinated", ignore_case = TRUE)) ~ "PP",

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

      # --- PCL → PLA (polycaprolactone; biodegradable polyester, closest available group)
      x_up == "PCL" ~ "PLA",
      stringr::str_detect(x, stringr::regex("\\bpolycaprolactone\\b", ignore_case = TRUE)) ~ "PLA",
      stringr::str_detect(x, stringr::regex("\\bPCL\\b", ignore_case = TRUE)) ~ "PLA",

      # --- PLA
      x_up == "PLA" ~ "PLA",
      stringr::str_detect(x, stringr::regex("polylact", ignore_case = TRUE)) ~ "PLA",

      # --- PVAc → PVA (poly(vinyl acetate); same vinyl backbone, closest available group)
      x_up == "PVAC" ~ "PVA",
      stringr::str_detect(x, stringr::regex("\\bPVAC\\b", ignore_case = TRUE)) ~ "PVA",
      stringr::str_detect(x, stringr::regex("poly\\s*\\(\\s*vinyl\\s*acetate\\s*\\)", ignore_case = TRUE)) ~ "PVA",
      stringr::str_detect(x, stringr::regex("polyvinyl\\s*acetate", ignore_case = TRUE)) ~ "PVA",

      # --- PVA
      x_up == "PVA" ~ "PVA",
      stringr::str_detect(x, stringr::regex("poly\\s*\\(\\s*vinyl\\s*alcohol\\s*\\)", ignore_case = TRUE)) ~ "PVA",
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

make_material_mapping_table <- function(df, raw_col, mapped_col) {
  mapping <- unique(df[, c(raw_col, mapped_col)])
  names(mapping) <- c("raw_material_name", "standardized_group")
  mapping <- mapping[order(mapping$raw_material_name, na.last = TRUE), ]
  rownames(mapping) <- NULL
  mapping
}

# ---------------------------
# Instrument auto-detection
# ---------------------------

# Peek at column names of the first uploaded file and return one of:
# "ldir", "raman", "ftir", or NA_character_ (unknown).
detect_instrument_type <- function(paths) {
  if (length(paths) == 0) return(NA_character_)
  path <- paths[[1]]
  ext  <- tolower(tools::file_ext(path))

  cols <- NULL

  # xlsx: read sheet 2 headers (LDIR particle data sheet)
  if (ext %in% c("xlsx", "xls")) {
    df_peek <- tryCatch(
      readxl::read_excel(path, sheet = 2, n_max = 0),
      error = function(e) NULL
    )
    if (!is.null(df_peek)) cols <- names(df_peek)
  }

  # CSV (or failed xlsx path): try multiple encodings and delimiters
  if (is.null(cols)) {
    first_line <- tryCatch(readLines(path, n = 1, warn = FALSE), error = function(e) "")
    delim <- if (length(first_line) > 0 && grepl(";", first_line[1], fixed = TRUE)) ";" else ","

    for (enc in c("UTF-8", "Latin1", "Windows-1252", "UTF-16LE")) {
      df_peek <- tryCatch(
        readr::read_delim(
          path, delim = delim, n_max = 0,
          locale = readr::locale(encoding = enc),
          show_col_types = FALSE, progress = FALSE
        ),
        error = function(e) NULL
      )
      if (!is.null(df_peek) && length(names(df_peek)) > 0) {
        candidate <- names(df_peek)
        # Prefer the encoding that gives us a recognisable signature column
        if (any(c("Eccentricity", "RamanSignal") %in% candidate) ||
            any(grepl("Area on map", candidate, fixed = TRUE))) {
          cols <- candidate
          break
        }
        if (is.null(cols)) cols <- candidate  # keep first successful parse as fallback
      }
    }
  }

  if (is.null(cols) || length(cols) == 0) return(NA_character_)

  # Detection order per spec
  if ("Eccentricity" %in% cols)                      return("ldir")
  if ("RamanSignal"  %in% cols)                      return("raman")
  if (any(grepl("Area on map", cols, fixed = TRUE))) return("ftir")

  NA_character_
}

# For FTIR files, try to disambiguate Spotlight vs Lumos from filenames.
# Returns "ftir_spotlight", "ftir_lumos", or NA (needs modal).
detect_ftir_subtype <- function(paths) {
  # Check both the filename and the full path (folders often contain "Lumos" / "SpotLight")
  haystack <- tolower(c(basename(paths), paths))
  if (any(grepl("spotlight", haystack, fixed = TRUE))) return("ftir_spotlight")
  if (any(grepl("lumos",     haystack, fixed = TRUE))) return("ftir_lumos")
  NA_character_
}

# ---------------------------
# FT-IR workflow
# ---------------------------
read_one_csv_ftir <- function(path) {

  try_read <- function(enc) {
    readr::read_csv(
      path,
      show_col_types = FALSE,
      locale = readr::locale(encoding = enc)
    )
  }

  candidates <- c("UTF-8", "UTF-16LE", "UTF-16BE", "Windows-1252", "Latin1")

  dfs    <- list()
  scores <- numeric(length(candidates))

  for (i in seq_along(candidates)) {
    enc  <- candidates[[i]]
    df_i <- tryCatch(try_read(enc), error = function(e) NULL)
    if (is.null(df_i)) { scores[i] <- -Inf; next }

    cn      <- names(df_i)
    cn_utf8 <- iconv(cn, from = "", to = "UTF-8", sub = "byte")
    names(df_i) <- cn_utf8
    cn <- cn_utf8

    suppressWarnings({
      has_micro <- any(grepl("µ", cn, fixed = TRUE))
      has_sq    <- any(grepl("²", cn, fixed = TRUE)) || any(grepl("³", cn, fixed = TRUE))
      has_repl  <- any(grepl("�", cn))
    })

    scores[i] <- (has_micro * 10) + (has_sq * 3) - (has_repl * 100)
    dfs[[i]]  <- df_i
  }

  best_i <- which.max(scores)
  df     <- dfs[[best_i]]

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

  # Locate "Area on map [µm²]" flexibly to handle minor µ encoding differences
  area_col <- cn[grepl("^Area on map", cn)]
  if (length(area_col) == 0) {
    stop("Missing 'Area on map [µm²]' column in FTIR file: ", basename(path))
  }
  area_col <- area_col[1]

  df <- df %>%
    dplyr::filter(stringr::str_starts(as.character(`Identifier`), "MP")) %>%
    dplyr::mutate(
      group_raw          = stringr::str_trim(as.character(`Group`)),
      group_mapped       = standardize_group_code(group_raw),
      group_code         = dplyr::if_else(is.na(group_mapped) | group_mapped == "", group_raw, group_mapped),
      feret_um           = as.numeric(`Feret min [µm]`),
      mass_ng            = as.numeric(`Mass [ng]`),
      `CE Diameter [µm]` = sqrt(4 * as.numeric(.data[[area_col]]) / pi),
      source_file        = basename(path)
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
  dfs     <- lapply(input_paths, read_one_csv_ftir)
  long_df <- dplyr::bind_rows(dfs)

  mapping_df <- make_material_mapping_table(long_df, "group_raw", "group_mapped")
  long_df    <- long_df %>% dplyr::select(-c(group_raw, group_mapped, group_code, feret_um, mass_ng))

  all_A <- list()
  all_B <- list()
  all_C <- list()

  wb <- createWorkbook()

  addWorksheet(wb, "Long_Table")
  safeWriteTable(wb, "Long_Table", fix_colnames_for_excel(long_df))
  setColWidths(wb, "Long_Table", cols = 1:min(30, ncol(long_df)), widths = 10)

  for (i in seq_along(dfs)) {
    df_one    <- dfs[[i]]
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

  addWorksheet(wb, "Material_Mapping")
  safeWriteTable(wb, "Material_Mapping", fix_colnames_for_excel(mapping_df))
  setColWidths(wb, "Material_Mapping", cols = 1, widths = 50)
  setColWidths(wb, "Material_Mapping", cols = 2, widths = 20)

  saveWorkbook(wb, output_path, overwrite = TRUE)
  output_path
}

# ---------------------------
# Raman workflow
# ---------------------------

read_one_csv_raman <- function(path, hqi_cutoff = 70) {

  first_line <- tryCatch(readLines(path, n = 1, warn = FALSE), error = function(e) "")
  delim <- if (length(first_line) && grepl(";", first_line, fixed = TRUE) && !grepl(",", first_line, fixed = TRUE)) {
    ";"
  } else if (length(first_line) && grepl(";", first_line, fixed = TRUE) && grepl(",", first_line, fixed = TRUE)) {
    if (stringr::str_count(first_line, fixed(";")) >= stringr::str_count(first_line, fixed(","))) ";" else ","
  } else {
    ","
  }

  try_read <- function(enc) {
    readr::read_delim(
      file = path,
      delim = delim,
      show_col_types = FALSE,
      locale = readr::locale(encoding = enc),
      progress = FALSE,
      trim_ws = TRUE
    )
  }

  candidates <- c("UTF-8", "UTF-16LE", "UTF-16BE", "Windows-1252", "Latin1")
  dfs    <- list()
  scores <- numeric(length(candidates))

  for (i in seq_along(candidates)) {
    enc  <- candidates[[i]]
    df_i <- tryCatch(try_read(enc), error = function(e) NULL)
    if (is.null(df_i)) { scores[i] <- -Inf; next }

    cn <- names(df_i)
    if (is.null(cn)) { scores[i] <- -Inf; next }

    cn_utf8 <- iconv(cn, from = "", to = "UTF-8", sub = "byte")
    names(df_i) <- cn_utf8

    suppressWarnings({
      has_micro <- any(grepl("µ", cn_utf8, fixed = TRUE)) || any(grepl("μ", cn_utf8, fixed = TRUE))
      has_repl  <- any(grepl("�", cn_utf8))
      has_bytes <- any(grepl("<[0-9A-Fa-f]{2}>", cn_utf8, perl = TRUE))
    })

    scores[i] <- (has_micro * 5) - (has_repl * 100) - (has_bytes * 20)
    dfs[[i]]  <- df_i
  }

  best_i <- which.max(scores)
  raw_df  <- dfs[[best_i]]
  if (is.null(raw_df) || !is.finite(scores[best_i])) {
    stop(
      "Could not read Raman CSV reliably (encoding/header issue) for file: ", basename(path), "\n",
      "Tried encodings: ", paste(candidates, collapse = ", "), "\n",
      "Tip: export as UTF-8/UTF-16 CSV with headers."
    )
  }

  raw_names <- names(raw_df)
  if (is.null(raw_names)) {
    stop(
      "Raman CSV has no header row (NULL column names) in file: ", basename(path), "\n",
      "Please re-export with column headers enabled."
    )
  }

  # Detect area column from raw names BEFORE janitor cleans them.
  # Match "Area [...]" but NOT "Convex Area [...]".
  area_raw_idx <- which(
    grepl("^Area[[:space:]]*\\[", raw_names) &
      !grepl("Convex", raw_names, ignore.case = TRUE)
  )
  cleaned_raw <- janitor::make_clean_names(raw_names, replace = c("µ" = "u", "μ" = "u"))
  col_area    <- if (length(area_raw_idx) > 0) cleaned_raw[[area_raw_idx[1]]] else NA_character_

  names(raw_df) <- cleaned_raw

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
    material_raw     = col_material,
    hqi_value_raw    = col_hqi,
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
      particle_name    = if (!is.na(col_particle)) as.character(.data[[col_particle]]) else NA_character_,
      material_raw     = as.character(.data[[col_material]]),
      hqi_value_raw    = parse_num_any(.data[[col_hqi]]),
      feret_max_um_raw = parse_num_any(.data[[col_feret]]),
      source_file      = basename(path)
    )

  validate_schema(
    df,
    required_cols = c("particle_name", "material_raw", "hqi_value_raw", "feret_max_um_raw", "source_file"),
    context = paste0("Raman file ", basename(path))
  )

  df <- df %>%
    dplyr::mutate(
      material_raw     = stringr::str_trim(as.character(material_raw)),
      material_mapped  = standardize_group_code(material_raw),
      group_code       = dplyr::if_else(is.na(material_mapped) | material_mapped == "", "OTHER", material_mapped),
      hqi_value        = dplyr::if_else(!is.na(hqi_value_raw) & hqi_value_raw <= 1, hqi_value_raw * 100, hqi_value_raw),
      feret_max_um     = feret_max_um_raw,
      `CE Diameter [µm]` = if (!is.na(col_area)) {
        sqrt(4 * parse_num_any(.data[[col_area]]) / pi)
      } else {
        NA_real_
      }
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
    count_by_group        = count_by_group,
    count_by_feret        = count_by_feret,
    unknown_materials_long = unknown_materials_long
  )
}


write_output_workbook_raman <- function(input_paths, output_path, hqi_cutoff = 70) {
  dfs     <- lapply(input_paths, read_one_csv_raman, hqi_cutoff = hqi_cutoff)
  long_df <- dplyr::bind_rows(dfs)
  mapping_df <- make_material_mapping_table(long_df, "material_raw", "material_mapped")
  long_df <- fix_colnames_for_excel(long_df)

  all_A     <- list()
  all_B     <- list()
  all_unknown <- list()

  wb <- createWorkbook()

  addWorksheet(wb, "Long_Table")
  safeWriteTable(wb, "Long_Table", long_df)
  setColWidths(wb, "Long_Table", cols = 1:ncol(long_df), widths = 10)

  for (i in seq_along(dfs)) {
    df_one    <- dfs[[i]]
    file_name  <- basename(input_paths[[i]])
    sheet_name <- sanitize_sheet_name(tools::file_path_sans_ext(file_name))

    original <- sheet_name
    k <- 2
    while (sheet_name %in% names(wb)) {
      sheet_name <- sanitize_sheet_name(paste0(original, "_", k))
      k <- k + 1
    }

    piv <- make_pivots_raman(df_one, src_override = basename(input_paths[[i]]))
    piv$count_by_group         <- fix_colnames_for_excel(piv$count_by_group)
    piv$count_by_feret         <- fix_colnames_for_excel(piv$count_by_feret)
    piv$unknown_materials_long <- fix_colnames_for_excel(piv$unknown_materials_long)

    all_A[[i]]       <- piv$count_by_group
    all_B[[i]]       <- piv$count_by_feret
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

  summary_A   <- dplyr::bind_rows(all_A)
  summary_B   <- dplyr::bind_rows(all_B)
  unknown_all <- dplyr::bind_rows(all_unknown)

  summary_A   <- fix_colnames_for_excel(add_total_row(summary_A, id_col = "source_file", label = "Total"))
  summary_B   <- fix_colnames_for_excel(add_total_row(summary_B, id_col = "source_file", label = "Total"))
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

  addWorksheet(wb, "Material_Mapping")
  safeWriteTable(wb, "Material_Mapping", fix_colnames_for_excel(mapping_df))
  setColWidths(wb, "Material_Mapping", cols = 1, widths = 50)
  setColWidths(wb, "Material_Mapping", cols = 2, widths = 20)

  addWorksheet(wb, "Unknown_Materials")
  safeWriteTable(wb, "Unknown_Materials", unknown_all)
  setColWidths(wb, "Unknown_Materials", cols = 1:ncol(unknown_all), widths = 10)

  preflight <- dplyr::bind_rows(
    preflight_table(long_df,    "Long_Table"),
    preflight_table(summary_A,  "Total_A"),
    preflight_table(summary_B,  "Total_B"),
    preflight_table(mean_A,     "Summary_A"),
    preflight_table(mean_B,     "Summary_B"),
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
# LDIR workflow
# ---------------------------

read_one_xlsx_ldir <- function(path) {
  df <- readxl::read_excel(path, sheet = 2)

  cn <- names(df)

  required_cols <- c("Eccentricity", "Is Valid", "Match Type", "Id")
  missing <- setdiff(required_cols, cn)
  if (length(missing) > 0) {
    stop(
      "Missing required columns in LDIR file: ", basename(path), "\n",
      "Missing: ", paste(missing, collapse = ", "), "\n",
      "Columns present: ", paste(cn, collapse = ", ")
    )
  }

  # Locate area column: "Area (µm²)" — parentheses notation, flexible µ encoding
  area_col <- cn[grepl("^Area[[:space:]]*\\(", cn)]
  if (length(area_col) == 0) {
    stop("Missing 'Area (µm²)' column in LDIR file: ", basename(path))
  }
  area_col <- area_col[1]

  df <- df %>%
    dplyr::mutate(
      material_raw       = stringr::str_trim(as.character(`Match Type`)),
      material_mapped    = standardize_group_code(material_raw),
      group_code         = dplyr::if_else(is.na(material_mapped) | material_mapped == "", "OTHER", material_mapped),
      `CE Diameter [µm]` = sqrt(4 * as.numeric(.data[[area_col]]) / pi),
      source_file        = basename(path)
    )

  df
}

make_pivots_ldir <- function(df_one, src_override = NULL) {
  src <- src_override
  if (is.null(src)) {
    src <- if ("source_file" %in% names(df_one) && nrow(df_one) > 0)
      df_one$source_file[[1]] else "Unknown"
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

  count_by_size <- df_one %>%
    dplyr::mutate(
      size_bin = bin_feret(`CE Diameter [µm]`),
      size_bin = factor(size_bin, levels = FERET_BINS)
    ) %>%
    dplyr::count(size_bin, name = "value") %>%
    tidyr::complete(size_bin = factor(FERET_BINS, levels = FERET_BINS), fill = list(value = 0)) %>%
    tidyr::pivot_wider(names_from = size_bin, values_from = value, values_fill = 0) %>%
    dplyr::mutate(source_file = src, .before = 1)
  count_by_size <- ensure_cols_in_order(count_by_size, FERET_BINS, id_col = "source_file")

  list(
    count_by_group = count_by_group,
    count_by_size  = count_by_size
  )
}

write_output_workbook_ldir <- function(input_paths, output_path) {
  dfs     <- lapply(input_paths, read_one_xlsx_ldir)
  long_df <- dplyr::bind_rows(dfs)

  mapping_df  <- make_material_mapping_table(long_df, "material_raw", "material_mapped")
  long_df_out <- long_df %>%
    dplyr::select(-c(material_raw, material_mapped, group_code)) %>%
    fix_colnames_for_excel()

  all_A <- list()
  all_B <- list()

  wb <- createWorkbook()

  addWorksheet(wb, "Long_Table")
  safeWriteTable(wb, "Long_Table", long_df_out)
  setColWidths(wb, "Long_Table", cols = 1:ncol(long_df_out), widths = 10)

  for (i in seq_along(dfs)) {
    df_one    <- dfs[[i]]
    file_name  <- basename(input_paths[[i]])
    sheet_name <- sanitize_sheet_name(tools::file_path_sans_ext(file_name))

    original <- sheet_name
    k <- 2
    while (sheet_name %in% names(wb)) {
      sheet_name <- sanitize_sheet_name(paste0(original, "_", k))
      k <- k + 1
    }

    piv <- make_pivots_ldir(df_one, src_override = basename(input_paths[[i]]))
    piv$count_by_group <- fix_colnames_for_excel(piv$count_by_group)
    piv$count_by_size  <- fix_colnames_for_excel(piv$count_by_size)

    all_A[[i]] <- piv$count_by_group
    all_B[[i]] <- piv$count_by_size

    addWorksheet(wb, sheet_name)
    r <- 1

    writeData(wb, sheet_name, "A) Counts per polymer type", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, piv$count_by_group, startRow = r, startCol = 1)
    r <- r + nrow(piv$count_by_group) + 3

    writeData(wb, sheet_name, "B) Counts per size class (CE Diameter)", startRow = r, startCol = 1)
    r <- r + 1
    safeWriteTable(wb, sheet_name, piv$count_by_size, startRow = r, startCol = 1)

    setColWidths(wb, sheet_name, cols = 1:max(ncol(piv$count_by_group), ncol(piv$count_by_size)), widths = 10)
  }

  summary_A <- dplyr::bind_rows(all_A)
  summary_B <- dplyr::bind_rows(all_B)

  summary_A <- fix_colnames_for_excel(add_total_row(summary_A, id_col = "source_file", label = "Total"))
  summary_B <- fix_colnames_for_excel(add_total_row(summary_B, id_col = "source_file", label = "Total"))

  addWorksheet(wb, "Total")
  rr <- 1

  writeData(wb, "Total", "A) Counts per polymer type", startRow = rr, startCol = 1)
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
  writeData(wb, "Summary", "A) Mean counts per polymer type (across files)", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", mean_A, startRow = r2, startCol = 1)
  r2 <- r2 + nrow(mean_A) + 3

  writeData(wb, "Summary", "B) Mean counts per size class (across files)", startRow = r2, startCol = 1)
  r2 <- r2 + 1
  safeWriteTable(wb, "Summary", mean_B, startRow = r2, startCol = 1)
  setColWidths(wb, "Summary", cols = 1:max(ncol(mean_A), ncol(mean_B)), widths = 10)

  addWorksheet(wb, "Material_Mapping")
  safeWriteTable(wb, "Material_Mapping", fix_colnames_for_excel(mapping_df))
  setColWidths(wb, "Material_Mapping", cols = 1, widths = 50)
  setColWidths(wb, "Material_Mapping", cols = 2, widths = 20)

  saveWorkbook(wb, output_path, overwrite = TRUE)
  output_path
}


# ---------------------------
# Shiny app
# ---------------------------

ui <- fluidPage(
  titlePanel("Microplastic Data Processor"),

  sidebarLayout(
    sidebarPanel(
      shinyFilesButton(
        id    = "files",
        label = "Choose files (CSV or XLSX)",
        title = "Select data files to process",
        multiple = TRUE
      ),
      tags$small("Raman: .csv • FTIR: .csv • LDIR: .xlsx"),
      br(), br(),

      uiOutput("instrument_label"),
      br(),

      uiOutput("hqi_ui"),

      tags$small("Output will be saved next to the selected files."),
      br(), br(),

      actionButton("process", "Process & Save", class = "btn-primary"),
      br(), br(),
      strong("Status:"),
      verbatimTextOutput("status"),
      br(),
      strong("Output path:"),
      verbatimTextOutput("out_path"),
      uiOutput("open_folder_btn")
    ),
    mainPanel(
      h4("Selected files"),
      tableOutput("file_table")
    )
  )
)

server <- function(input, output, session) {

  roots <- c("Home" = normalizePath("~", winslash = "/", mustWork = TRUE))
  if (.Platform$OS.type == "windows") {
    # Scan every drive letter so mapped network drives appear automatically
    for (letter in LETTERS) {
      path <- paste0(letter, ":/")
      if (file.exists(path)) roots[[letter]] <- path
    }
    # Also expose the IAEA network share directly via UNC if reachable
    unc <- "//vs-monaco1.iaea.org/REL"
    if (file.exists(unc)) roots[["REL (vs-monaco1)"]] <- unc
  }

  shinyFileChoose(input, "files", roots = roots, filetypes = c("csv", "xlsx", "xls"))

  selected_paths <- reactive({
    req(input$files)
    paths <- shinyFiles::parseFilePaths(roots, input$files)
    as.character(paths$datapath)
  })

  output$file_table <- renderTable({
    req(selected_paths())
    data.frame(
      File   = basename(selected_paths()),
      Folder = dirname(selected_paths()),
      stringsAsFactors = FALSE
    )
  })

  # Holds the confirmed instrument type: "raman", "ldir", "ftir_spotlight",
  # "ftir_lumos", "ftir" (unconfirmed), or NA.
  detected_type <- reactiveVal(NA_character_)

  # Detect instrument whenever the file selection changes
  observeEvent(selected_paths(), {
    detected_type(NA_character_)
    paths <- selected_paths()
    if (length(paths) == 0) return()

    type <- tryCatch(
      detect_instrument_type(paths),
      error = function(e) NA_character_
    )

    if (is.na(type)) {
      detected_type(NA_character_)
      return()
    }

    if (type == "ftir") {
      sub <- detect_ftir_subtype(paths)
      if (!is.na(sub)) {
        detected_type(sub)
      } else {
        detected_type("ftir")  # waiting for user confirmation
        showModal(modalDialog(
          title = "FTIR Instrument Type",
          p("The filename does not indicate whether this is FTIR Spotlight or FTIR Lumos."),
          p("Please select the correct instrument before processing:"),
          radioButtons(
            inputId  = "ftir_subtype_choice",
            label    = NULL,
            choices  = c("FTIR Spotlight" = "ftir_spotlight",
                         "FTIR Lumos"     = "ftir_lumos"),
            selected = "ftir_spotlight"
          ),
          footer    = actionButton("ftir_confirm", "Confirm", class = "btn-primary"),
          easyClose = FALSE
        ))
      }
    } else {
      detected_type(type)
    }
  })

  observeEvent(input$ftir_confirm, {
    detected_type(input$ftir_subtype_choice)
    removeModal()
  })

  output$instrument_label <- renderUI({
    type <- detected_type()
    paths <- if (length(selected_paths()) > 0) selected_paths() else character(0)

    if (is.null(type) || is.na(type)) {
      if (length(paths) > 0)
        return(tags$p(tags$em("Could not detect instrument type."), style = "color:#c0392b;"))
      return(tags$p(tags$em("No files selected."), style = "color:gray;"))
    }

    label <- switch(type,
      "raman"          = "Raman",
      "ldir"           = "LDIR (Agilent 8700)",
      "ftir_spotlight" = "FTIR Spotlight",
      "ftir_lumos"     = "FTIR Lumos",
      "ftir"           = "FTIR (awaiting confirmation…)",
      paste("Unknown:", type)
    )
    tags$p(
      tags$strong("Detected instrument: "),
      tags$span(label, style = "color:#2c7bb6; font-weight:bold;")
    )
  })

  # Show HQI cutoff only for Raman
  output$hqi_ui <- renderUI({
    type <- detected_type()
    if (!is.null(type) && !is.na(type) && type == "raman") {
      tagList(
        numericInput(
          inputId = "hqi_cutoff",
          label   = "HQI cutoff (%)",
          value   = 70, min = 0, max = 100, step = 1
        ),
        tags$small("Rows with HQI below the cutoff are discarded before processing."),
        br(), br()
      )
    }
  })

  status_val   <- reactiveVal("Select files, then click Process & Save.")
  out_path_val <- reactiveVal("")

  output$status   <- renderText(status_val())
  output$out_path <- renderText(out_path_val())

  observeEvent(input$process, {
    req(selected_paths())
    paths <- selected_paths()
    type  <- detected_type()

    if (is.null(type) || is.na(type)) {
      status_val("ERROR: Instrument type could not be detected. Check your file.")
      return()
    }
    if (type == "ftir") {
      status_val("Please confirm the FTIR instrument type in the dialog first.")
      return()
    }

    type_label <- switch(type,
      "raman"          = "raman",
      "ldir"           = "ldir",
      "ftir_spotlight" = "ftir_spotlight",
      "ftir_lumos"     = "ftir_lumos",
      type
    )

    out_dir  <- dirname(paths[[1]])
    out_file <- file.path(
      out_dir,
      paste0("processed_data_", type_label, "_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")
    )

    status_val("Processing…")
    out_path_val("")

    tryCatch({
      res <- if (type == "raman") {
        hqi <- if (!is.null(input$hqi_cutoff)) input$hqi_cutoff else 70
        write_output_workbook_raman(paths, out_file, hqi_cutoff = hqi)
      } else if (type == "ldir") {
        write_output_workbook_ldir(paths, out_file)
      } else if (type %in% c("ftir_spotlight", "ftir_lumos")) {
        write_output_workbook_ftir(paths, out_file)
      } else {
        stop("Unknown instrument type: ", type)
      }
      status_val("Done.")
      out_path_val(res)
    }, error = function(e) {
      status_val(paste("ERROR:", conditionMessage(e)))
      out_path_val("")
    })
  })

  output$open_folder_btn <- renderUI({
    req(nchar(out_path_val()) > 0)
    actionButton("open_folder", "Open output folder",
                 icon = icon("folder-open"), class = "btn-default btn-sm")
  })

  observeEvent(input$open_folder, {
    folder <- normalizePath(dirname(out_path_val()), winslash = "\\", mustWork = FALSE)
    if (.Platform$OS.type == "windows") {
      shell.exec(folder)
    } else if (Sys.info()[["sysname"]] == "Darwin") {
      system2("open", shQuote(folder))
    } else {
      system2("xdg-open", shQuote(folder))
    }
  })
}

shinyApp(ui, server)
