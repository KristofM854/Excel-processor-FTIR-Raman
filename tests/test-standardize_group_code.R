# tests/test-standardize_group_code.R
#
# Run from the project root:
#   source("tests/test-standardize_group_code.R")
#
# Purpose:
#   Validate standardize_group_code() mappings, print diagnostics, and list
#   any inputs that remain NA for manual review.

message("=== Loading app.R ===")
suppressMessages(source("app.R"))

# ---------------------------------------------------------------------------
# Test vectors
# ---------------------------------------------------------------------------

# Known-good assertions (input → expected output)
assert_cases <- list(
  # New mappings: EVOH → PVA
  list(input = "BIAXIALLY-ORIENTED EVOH FILM",             expected = "PVA"),
  list(input = "EVOH",                                     expected = "PVA"),
  list(input = "Ethylene vinyl alcohol copolymer",         expected = "PVA"),
  list(input = "vinyl alcohol copolymer (EVOH type)",      expected = "PVA"),

  # New mappings: PVAc → PVA
  list(input = "Poly(Vinyl Acetate) 15000",                expected = "PVA"),
  list(input = "PVAC",                                     expected = "PVA"),
  list(input = "polyvinyl acetate resin",                  expected = "PVA"),  # has "resin" but also POLY cue

  # PVA extended pattern
  list(input = "POLY(VINYL ALCOHOL) 500",                  expected = "PVA"),
  list(input = "polyvinyl alcohol",                        expected = "PVA"),

  # New mappings: PCL → PLA
  list(input = "POLYCAPROLACTONE",                         expected = "PLA"),
  list(input = "PCL",                                      expected = "PLA"),
  list(input = "Polycaprolactone diol 2000",               expected = "PLA"),

  # New mappings: EPDM → IR
  list(input = "POLY(ETHYLENE:PROPYLENE:DIENE) M-CLASS",   expected = "IR"),
  list(input = "EPDM",                                     expected = "IR"),
  list(input = "EPR",                                      expected = "IR"),
  list(input = "ethylene-propylene-diene rubber",          expected = "IR"),
  list(input = "ethylene propylene diene monomer",         expected = "IR"),

  # Masterbatch: explicit polymer → correct group
  list(input = "COLOR MASTERBATCH POLY[ETHYLENE-CO-(VINYL ACETATE)] + 45%",
                                                           expected = "EVA"),
  list(input = "Masterbatch polyethylene base",            expected = "PE"),
  list(input = "Polypropylene film",                       expected = "PP"),

  # Existing mappings still work
  list(input = "Polyethylene",                             expected = "PE"),
  list(input = "Polypropylene",                            expected = "PP"),
  list(input = "Polystyrene",                              expected = "PS"),
  list(input = "Polyamide 6",                              expected = "PA"),
  list(input = "Nylon 6,6",                                expected = "PA"),
  list(input = "Polyethylene terephthalate",               expected = "PET"),
  list(input = "PETG",                                     expected = "PET"),
  list(input = "Polycarbonate",                            expected = "PC"),
  list(input = "Polymethyl methacrylate",                  expected = "PMMA"),
  list(input = "Polyvinyl chloride",                       expected = "PVC"),
  list(input = "Polylactic acid",                          expected = "PLA"),
  list(input = "Polyurethane",                             expected = "PU"),
  list(input = "Polytetrafluoroethylene",                  expected = "PTFE"),
  list(input = "ABS",                                      expected = "ABS"),

  # Non-polymers → NA (guard clause)
  list(input = "ARABIC GUM",                               expected = NA_character_),
  list(input = "SILICON(IV) OXIDE",                        expected = NA_character_),
  list(input = "LEMON OIL",                                expected = NA_character_),
  list(input = "PARAFFIN",                                 expected = NA_character_),
  list(input = "TITANIUM DIOXIDE",                         expected = NA_character_),
  list(input = "ZINC STEARATE",                            expected = NA_character_),
  list(input = "SODIUM PHOSPHATE",                         expected = NA_character_),
  list(input = "CARBON BLACK PIGMENT",                     expected = NA_character_),
  list(input = "QUARTZ SAND",                              expected = NA_character_)
)

# ---------------------------------------------------------------------------
# Run assertions
# ---------------------------------------------------------------------------
inputs   <- vapply(assert_cases, `[[`, character(1), "input")
expected <- vapply(assert_cases, function(x) {
  v <- x[["expected"]]
  if (is.null(v)) NA_character_ else as.character(v)
}, character(1))

results <- standardize_group_code(inputs)

pass  <- mapply(function(r, e) identical(r, e), results, expected)
fails <- which(!pass)

# ---------------------------------------------------------------------------
# Diagnostics: table of returned groups
# ---------------------------------------------------------------------------
message("\n=== Result table ===")
result_df <- data.frame(
  input    = inputs,
  expected = expected,
  got      = results,
  pass     = pass,
  stringsAsFactors = FALSE
)
print(result_df, row.names = FALSE)

# Inputs that remain NA
message("\n=== Inputs returning NA ===")
na_idx <- which(is.na(results))
if (length(na_idx) == 0) {
  message("  (none)")
} else {
  for (i in na_idx) message("  [", i, "] ", inputs[i])
}

# Inputs matched by new rules
new_rule_labels <- c(
  "EVOH"  = "\\bEVOH\\b|ethylene\\s*vinyl\\s*alcohol|vinyl\\s*alcohol\\s*copolymer",
  "PVAc"  = "poly\\s*\\(\\s*vinyl\\s*acetate\\s*\\)|polyvinyl\\s*acetate|\\bPVAC\\b",
  "PCL"   = "\\bpolycaprolactone\\b|\\bPCL\\b",
  "EPDM"  = "\\bEPDM\\b|\\bEPR\\b|ethylene.{0,3}propylene.{0,3}diene"
)
message("\n=== Inputs matched by new rules ===")
for (rule_name in names(new_rule_labels)) {
  pat <- new_rule_labels[[rule_name]]
  matched <- inputs[grepl(pat, inputs, ignore.case = TRUE) & !is.na(results)]
  if (length(matched)) {
    message("  ", rule_name, " → ", unique(results[grepl(pat, inputs, ignore.case = TRUE) & !is.na(results)]))
    for (m in matched) message("    ", m)
  }
}

# ---------------------------------------------------------------------------
# Broad diagnostic: run on extended realistic vector and list NA remainders
# ---------------------------------------------------------------------------
extended_vector <- c(
  inputs,
  "EVOH FILM", "Ethylene-Vinyl Alcohol (EVOH)", "Poly[Ethylene-co-Vinyl Alcohol]",
  "Polycaprolactone triol", "poly-epsilon-caprolactone",
  "EPDM rubber", "Ethylene/Propylene/Diene copolymer",
  "Polyethylene wax", "Polypropylene wax",
  "Cellulose acetate", "Cellulose acetate butyrate",
  "Polyacrylamide", "Polyacrylate",
  "Gellan gum", "Gum arabic", "Carrageenan",
  "Mineral oil", "Silicone oil", "Paraffin wax",
  "Iron oxide", "Zinc oxide", "Magnesium stearate",
  "Sodium chloride", "Calcium carbonate",
  "PDMS", "Silicone", "Dimethyl siloxane"
)

ext_results <- standardize_group_code(extended_vector)
message("\n=== Extended vector: group distribution ===")
print(sort(table(ext_results, useNA = "ifany")))

polymer_na <- extended_vector[is.na(ext_results) &
  grepl("^poly|\\bpoly|\\bPCL\\b|\\bEPDM\\b|\\bEVOH\\b|\\bPVAC\\b|rubber|elastomer",
        extended_vector, ignore.case = TRUE)]
message("\n=== Polymer-like strings still NA (candidates for manual review) ===")
if (length(polymer_na) == 0) {
  message("  (none)")
} else {
  for (s in polymer_na) message("  ", s)
}

# ---------------------------------------------------------------------------
# Final pass / fail summary
# ---------------------------------------------------------------------------
message("\n=== Assertion summary ===")
message("  Total : ", length(assert_cases))
message("  Passed: ", sum(pass))
message("  Failed: ", length(fails))

if (length(fails) > 0) {
  message("\nFAILED cases:")
  for (i in fails) {
    message("  [", i, "] '", inputs[i], "'")
    message("       expected: ", expected[i])
    message("       got     : ", results[i])
  }
  stop("standardize_group_code() has ", length(fails), " failing assertion(s). See output above.")
} else {
  message("\nAll assertions PASSED.")
}
