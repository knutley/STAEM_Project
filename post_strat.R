# =============================================================================
# Holyrood 2026 — Marginal Poststratification Script
# =============================================================================
#
# METHOD: Marginal-based synthetic estimation
#
#   True MRP needs individual-level microdata + a multilevel regression model.
#   This script does something honest and defensible with what you *do* have:
#   your demographic crosstabs (marginal distributions) + NRS Census 2022 data.
#
#   For each constituency it asks: given this constituency's age/sex mix,
#   what vote share does each party get if we apply the crosstab estimates?
#   The result is demographically-adjusted, constituency-level vote shares.
#
#   This is sometimes called "synthetic estimation" or "MRP-lite".
#   It is NOT the same as running a multilevel regression + full poststrat,
#   but it's a legitimate improvement over uniform national swing.
#
# DATA YOU NEED TO PROVIDE:
#   data/census_constituencies.csv  — NRS Census 2022, age × sex by constituency
#   data/results_2021.csv           — 2021 Holyrood constituency results
#   (Crosstab URLs come from your Google Sheets — see SOURCES below)
#
# OUTPUT:
#   output/constituency_estimates.csv   — wide format, one row per constituency
#   output/constituency_estimates.json  — dashboard-ready JSON
#
# INSTALL REQUIRED PACKAGES (run once):
#   install.packages(c("tidyverse", "readr", "jsonlite", "janitor", "lubridate", "readxl"))
# =============================================================================

library(tidyverse)
library(readr)
library(jsonlite)
library(janitor)
library(lubridate)
library(readxl)

# =============================================================================
# 0. CONFIGURATION
# =============================================================================

# ── Crosstab sheet URLs ──────────────────────────────────────────────────────
# For each pollster sheet: File → Share → Publish to web → pick the
# "Crosstabs" sheet tab → CSV → copy the URL and paste below.
CROSSTAB_URLS <- list(
  survation_const = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVHoDYDQBczBB68faWVZ6fGRNqPSjlJj9-yTKLzJc6h7P1OT6HQiFvz8aXVSG-JH9fZrPoq3X1Bw1K/pubhtml?gid=1374442542&single=true",
  survation_list  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVHoDYDQBczBB68faWVZ6fGRNqPSjlJj9-yTKLzJc6h7P1OT6HQiFvz8aXVSG-JH9fZrPoq3X1Bw1K/pubhtml?gid=1374442542&single=true",
  norstat_const   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQvwlOJGWbsgeU-1El-ruiXk0bo6_6Cv231HFmABtu7y1yukxmI8lrFDTveP3XSUXy6lfZi-RQ5_mKb/pubhtml?gid=31364899&single=true",
  norstat_list    = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQvwlOJGWbsgeU-1El-ruiXk0bo6_6Cv231HFmABtu7y1yukxmI8lrFDTveP3XSUXy6lfZi-RQ5_mKb/pubhtml?gid=31364899&single=true",
  ipsos_const     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRRKZ0skhiQDkOWJplehISh4YktsynvLoSHYvqizkckRD2E8_AzXE-x0cO_WbhZ0HlmMzRgVHCaZJUD/pubhtml?gid=1369530937&single=true",
  ipsos_list      = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSkkfZ99xen3Z3FVOwVnxoqY2JZhxkEzejwagwS4p1n1f8awbC453uv5yb-ydAsagXz108FC-efiuQ_/pubhtml?gid=1393854765&single=true"
)

# ── Party name → key mapping ─────────────────────────────────────────────────
# Left-hand side: exactly as the "Party" column appears in your crosstab sheets
PARTY_MAP <- c(
  "SNP"                = "snp",
  "Labour"             = "lab",
  "Conservative"       = "con",
  "Liberal Democrat"   = "lib",
  "Lib Dem"            = "lib",
  "Green"              = "grn",
  "Reform"             = "rfm",
  "Alba"               = "alba",
  "Other Party"        = "other"
)

# ── Approximate Scottish 16+ population share by age band ───────────────────
# Used to compute national-level weighted average from age marginals.
# Source: NRS 2022 mid-year population estimates. Replace if you have better.
AGE_WEIGHTS <- c(
  "16-24" = 0.13,
  "25-34" = 0.16,
  "35-49" = 0.26,
  "50-64" = 0.25,
  "65+"   = 0.20
)

# ── Headline sheet URLs (for national averages including Reform, Green, Alba) ──
# These are the same sheets the dashboard uses — first tab of each spreadsheet.
HEADLINE_CONST_URLS <- list(
  survation = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVHoDYDQBczBB68faWVZ6fGRNqPSjlJj9-yTKLzJc6h7P1OT6HQiFvz8aXVSG-JH9fZrPoq3X1Bw1K/pub?output=csv",
  norstat   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQvwlOJGWbsgeU-1El-ruiXk0bo6_6Cv231HFmABtu7y1yukxmI8lrFDTveP3XSUXy6lfZi-RQ5_mKb/pub?output=csv",
  ipsos     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSkkfZ99xen3Z3FVOwVnxoqY2JZhxkEzejwagwS4p1n1f8awbC453uv5yb-ydAsagXz108FC-efiuQ_/pub?output=csv"
)

# ── Manual polling override (optional) ────────────────────────────────────────
# If headline auto-loading gives wrong results, paste the averages from the
# dashboard's "Latest Constituency Vote" strip here and set to a named vector.
# Set to NULL to use auto-loaded averages.
# Example: MANUAL_NATIONAL_POLL <- c(snp=38, lab=22, con=18, lib=9, rfm=16, grn=6, alba=1)
MANUAL_NATIONAL_POLL <- c(
  snp  = 35.9,
  lab  = 16.8,
  con  = 11.0,
  lib  = 9.6,
  rfm  = 15.9,
  grn  = 5.2,
  alba = 0
)
OUTPUT_DIR <- "output"
DATA_DIR   <- "data"


# =============================================================================
# 1. LOAD CROSSTABS
# =============================================================================
# Expected columns (from your sheets):
#   Poll ID | Date (start) | Date (end) | Client | Party |
#   Demographic variable | Category | N | Percentage | URL

load_crosstab <- function(url, label) {
  if (startsWith(url, "PASTE")) {
    message("⚠  Skipping ", label, " — URL not yet set")
    return(NULL)
  }
  
  # ── Fix URL format ───────────────────────────────────────────────────────────
  # Google Sheets publish URLs come in two forms:
  #   pubhtml?gid=...&single=true   → serves HTML (what read_csv() gets here)
  #   pub?output=csv&gid=...        → serves CSV  (what we need)
  # Convert either form to the CSV endpoint automatically.
  url <- url |>
    str_replace("pubhtml", "pub") |>
    (\(u) if (!str_detect(u, "output=csv")) str_replace(u, "\\?", "?output=csv&") else u)()
  
  message("Loading ", label, "…  (", str_extract(url, "gid=[0-9]+"), ")")
  
  df <- tryCatch(
    read_csv(url, show_col_types = FALSE, col_types = cols(.default = "c")),
    error = function(e) { warning("Failed: ", label, " — ", e$message); NULL }
  )
  if (is.null(df) || nrow(df) == 0) {
    warning("Skipping ", label, " — empty response (check the gid is the Crosstabs tab)")
    return(NULL)
  }
  
  # Robust column renaming — handles slight header variations across sheets
  df <- df |> clean_names()
  col_renames <- c(
    poll_id       = names(df)[1],
    date_start    = names(df)[grepl("start",    names(df))][1],
    date_end      = names(df)[grepl("end",      names(df))][1],
    client        = names(df)[grepl("client",   names(df))][1],
    party         = names(df)[grepl("^party",   names(df))][1],
    demo_var      = names(df)[grepl("demo|variable", names(df))][1],
    category      = names(df)[grepl("categ",    names(df))][1],
    n_respondents = names(df)[grepl("^n$|^n_",  names(df))][1],
    pct           = names(df)[grepl("pct|percent", names(df))][1]
  )
  col_renames <- col_renames[!is.na(col_renames)]
  df <- df |> rename(any_of(setNames(col_renames, names(col_renames))))
  
  required <- c("party", "demo_var", "category", "pct")
  missing  <- setdiff(required, names(df))
  if (length(missing) > 0) {
    warning("Skipping ", label, " — missing columns: ", paste(missing, collapse = ", "),
            "\n  Available columns: ", paste(names(df), collapse = ", "))
    return(NULL)
  }
  
  df |>
    filter(!is.na(party), !is.na(category), !is.na(pct)) |>
    mutate(
      source    = label,
      vote_type = if_else(grepl("list|region", label, ignore.case = TRUE),
                          "list", "constituency"),
      pct_num   = as.numeric(str_remove_all(pct, "[%,\\s]")),
      party_key = recode(str_trim(party), !!!PARTY_MAP, .default = NA_character_),
      # Fix: check column exists before parsing — avoids the exists() scope bug
      date_end  = if ("date_end" %in% names(df)) dmy(date_end) else as.Date(NA)
    ) |>
    filter(!is.na(pct_num), !is.na(party_key), pct_num > 0)
}

crosstabs_raw <- map2_dfr(CROSSTAB_URLS, names(CROSSTAB_URLS), load_crosstab)

if (nrow(crosstabs_raw) == 0) {
  stop(
    "\nNo crosstab data loaded.\n",
    "Check: (1) URLs are pasted in correctly above\n",
    "       (2) The sheets are published as CSV (File → Publish to web)\n",
    "       (3) The Crosstabs tab — not the Headline tab — is selected"
  )
}

message("\n✓ Loaded ", nrow(crosstabs_raw), " crosstab rows from ",
        n_distinct(crosstabs_raw$source), " sources")


# =============================================================================
# 2. LATEST POLL PER POLLSTER + NORMALISE WITHIN CELLS
# =============================================================================

# For each source, keep only the most recent poll
latest <- crosstabs_raw |>
  group_by(source) |>
  filter(date_end == max(date_end, na.rm = TRUE) | all(is.na(date_end))) |>
  ungroup()

# Normalise percentages within each (source, demo_var, category) cell
# so that party shares sum to 100% — handles any "other" residuals cleanly
normalised <- latest |>
  group_by(source, vote_type, demo_var, category) |>
  mutate(
    cell_total = sum(pct_num, na.rm = TRUE),
    pct_norm   = if_else(cell_total > 0, pct_num / cell_total * 100, 0)
  ) |>
  ungroup()

# Equal-weight average across pollsters for each cell
crosstab_avg <- normalised |>
  group_by(vote_type, demo_var, category, party_key) |>
  summarise(pct_avg = mean(pct_norm, na.rm = TRUE), .groups = "drop")

# =============================================================================
# 2b. LOAD HEADLINE POLLS + CORRECT CROSSTAB INFLATION
# =============================================================================
# Green and Alba are missing from crosstab cells, so when normalised to 100%
# every present party (including Reform) is inflated. We fix this by:
#   1. Loading the headline sheets to get true national shares for ALL parties
#   2. Scaling crosstab cell shares back down so they sum to (100 - missing share)
#
# We also switch to ADDITIVE swing in step 6 — this avoids the ×10 multiplier
# problem for parties with near-zero 2021 constituency bases (Green, Reform).

# ── Load headline sheets ──────────────────────────────────────────────────────
# Column layout (0-based, dashboard spec):
# 0=Poll ID, 1=Date start, 2=Date end, 3=Pollster, 4=Client,
# 5=Sample, 6=LV only, 7=Undecideds excl, 8=Q num, 9=Q text,
# 10=Con, 11=Lab, 12=LibDem, 13=SNP, 14=Alba, 15=Reform, 16=Green, 17=Other
# After read_csv with no skip, row 1 is the header → columns are 1-indexed as above +1

load_headline_robust <- function(url, pollster_name, filter_undecideds = FALSE) {
  tryCatch({
    # Read with actual column names (no skip) so we don't mess up column indices
    raw <- read_csv(url, show_col_types = FALSE,
                    col_types = cols(.default = "c"),
                    col_names = FALSE)  # no header — we'll find it
    
    # Find the row where column 1 looks like a Poll ID (e.g., starts with letters)
    # Skip any truly blank rows at top
    data_start <- which(!is.na(raw[[1]]) & nchar(str_trim(raw[[1]])) > 0)[1]
    if (is.na(data_start)) return(NULL)
    
    # Row data_start is the header; data rows are data_start+1 onward
    header_row <- as.character(raw[data_start, ])
    df <- raw[(data_start + 1):nrow(raw), ]
    names(df) <- paste0("V", seq_len(ncol(df)))
    
    # Column positions (1-indexed in R):
    # V8 = Undecideds excl, V11=Con, V12=Lab, V13=LibDem, V14=SNP,
    # V15=Alba, V16=Reform, V17=Green
    parse_pct <- function(v) suppressWarnings(as.numeric(str_remove_all(v, "[%,\\s]")))
    
    if (filter_undecideds && "V8" %in% names(df)) {
      df <- df |> filter(str_trim(V8) == "Y")
    }
    if (nrow(df) == 0) return(NULL)
    df <- df |> slice_tail(n = 1)
    
    tibble(
      pollster = pollster_name,
      con  = parse_pct(df$V11),
      lab  = parse_pct(df$V12),
      lib  = parse_pct(df$V13),
      snp  = parse_pct(df$V14),
      alba = parse_pct(df$V15),
      rfm  = parse_pct(df$V16),
      grn  = parse_pct(df$V17)
    ) |> filter(if_any(c(snp, lab, con), \(x) !is.na(x) & x > 0))
    
  }, error = function(e) {
    message("  ⚠ Headline load failed for ", pollster_name, ": ", conditionMessage(e))
    NULL
  })
}

message("\nLoading headline constituency polls…")
headline_polls <- bind_rows(
  load_headline_robust(HEADLINE_CONST_URLS$survation, "Survation", filter_undecideds = TRUE),
  load_headline_robust(HEADLINE_CONST_URLS$norstat,   "Norstat",   filter_undecideds = FALSE),
  load_headline_robust(HEADLINE_CONST_URLS$ipsos,     "Ipsos",     filter_undecideds = FALSE)
)

# ── Diagnostic: show raw values per pollster before averaging ─────────────────
cat("\n=== Headline raw values per pollster (check these look right) ===\n")
if (nrow(headline_polls) > 0) {
  print(headline_polls)
} else {
  cat("  No headline polls loaded!\n")
}

message("  Loaded headline data from: ",
        if (nrow(headline_polls) > 0) paste(headline_polls$pollster, collapse = ", ") else "NONE")

# ── Compute national averages ─────────────────────────────────────────────────
if (!is.null(MANUAL_NATIONAL_POLL)) {
  nat_avg <- tibble(
    party_key  = names(MANUAL_NATIONAL_POLL),
    p_national = as.numeric(MANUAL_NATIONAL_POLL)
  ) |> filter(p_national > 0) |>
    mutate(p_national = p_national / sum(p_national) * 100)
  message("✓ Using MANUAL_NATIONAL_POLL override")
  cat("\n=== National averages (manual override) ===\n")
  print(nat_avg |> arrange(desc(p_national)))
} else if (nrow(headline_polls) > 0) {
  nat_avg <- headline_polls |>
    summarise(across(c(con, lab, lib, snp, alba, rfm, grn),
                     \(x) mean(x, na.rm = TRUE))) |>
    pivot_longer(everything(), names_to = "party_key", values_to = "p_national") |>
    filter(!is.na(p_national), p_national > 0) |>
    mutate(p_national = p_national / sum(p_national) * 100)
  cat("\n=== National headline averages (constituency) ===\n")
  print(nat_avg |> arrange(desc(p_national)))
  message("\n⚠ If the above numbers look wrong, set MANUAL_NATIONAL_POLL in config (section 0)")
  message("  and paste the values from the dashboard's 'Latest Constituency Vote' strip.")
} else {
  message("⚠ No headline polls loaded — using crosstab-derived averages as fallback")
  nat_avg <- NULL
}

# ── Correct crosstab inflation ────────────────────────────────────────────────
# Identify parties present in headline polls but ABSENT from ALL crosstab cells
# (not just some cells — a party that appears in any cell counts as "present")
ct_parties_present <- unique(crosstab_avg$party_key)

if (!is.null(nat_avg)) {
  # Parties in headline polls but with essentially no crosstab representation
  # (either entirely absent, or present in <5 cells nationally)
  ct_coverage <- crosstab_avg |>
    filter(vote_type == "constituency") |>
    count(party_key, name = "n_cells")
  
  fully_absent <- nat_avg |>
    left_join(ct_coverage, by = "party_key") |>
    filter(is.na(n_cells) | n_cells < 5) |>   # fewer than 5 cells = effectively absent
    pull(party_key)
  
  missing_share <- nat_avg |>
    filter(party_key %in% fully_absent) |>
    summarise(total = sum(p_national)) |>
    pull(total)
  
  if (length(fully_absent) > 0 && missing_share > 0 && missing_share < 40) {
    scale_factor <- (100 - missing_share) / 100
    crosstab_avg <- crosstab_avg |>
      mutate(pct_avg = pct_avg * scale_factor)
    message("\n✓ Crosstab correction: absent parties = ", paste(fully_absent, collapse = ", "),
            " (", round(missing_share, 1), "pp) → scaled remaining by ", round(scale_factor, 4))
  } else {
    message("  No crosstab correction needed or possible (missing share = ",
            round(missing_share, 1), "%)")
  }
}

# Quick sense check after correction
cat("\n=== Constituency crosstab averages — Age (after correction) ===\n")
crosstab_avg |>
  filter(vote_type == "constituency", demo_var == "Age") |>
  arrange(category, desc(pct_avg)) |>
  select(category, party_key, pct_avg) |>
  print(n = Inf)


# =============================================================================
# 3. LOAD NRS CENSUS DATA
# =============================================================================
# File: data/spc_population_estimates.xlsx (auto-downloaded if missing)
#
# What we know about this file from inspection:
#   - One sheet per year: General_notes, Table_of_contents, 2001, 2002, … 2021
#   - NO header row. Data starts at row 1.
#   - 95 columns: constituency | code | sex | all_ages | age_0 | age_1 | … | age_90
#   - 222 rows: 73 constituencies × 3 sex rows (Persons/Male/Female) + Scotland total × 3
#   - We keep only Male and Female rows; drop Persons and Scotland total.

NRS_URL <- paste0(
  "https://www.nrscotland.gov.uk/media/xpefjc4s/",
  "estimated-population-by-sex-single-year-of-age-and-scottish-parliamentary-constituency-area.xlsx"
)
NRS_PATH <- file.path(DATA_DIR, "spc_population_estimates.xlsx")

if (!file.exists(NRS_PATH)) {
  message("Downloading NRS SPC population estimates…")
  download.file(NRS_URL, destfile = NRS_PATH, mode = "wb", quiet = FALSE)
  message("✓ Saved to ", NRS_PATH)
}

# Pick the most recent year sheet (last of the numeric-named sheets)
all_sheets  <- excel_sheets(NRS_PATH)
year_sheets <- all_sheets[str_detect(all_sheets, "^[0-9]{4}$")]
nrs_sheet   <- year_sheets[length(year_sheets)]   # e.g. "2021"
message("Loading NRS sheet: '", nrs_sheet, "'")

# Read with no column names — the file has no header row
nrs_raw <- read_excel(NRS_PATH, sheet = nrs_sheet,
                      col_names = FALSE, col_types = "text")

# Columns: 1=constituency, 2=code, 3=sex, 4=all_ages, 5..95 = ages 0..90
n_age_cols <- ncol(nrs_raw) - 4
names(nrs_raw) <- c("constituency", "code", "sex", "all_ages",
                    as.character(seq(0, n_age_cols - 1)))
age_cols <- as.character(0:min(90, n_age_cols - 1))

census <- nrs_raw |>
  select(constituency, sex, all_of(age_cols)) |>
  filter(
    !is.na(constituency),
    nchar(str_trim(constituency)) > 0,
    !str_detect(tolower(constituency), "^scotland")
  ) |>
  mutate(sex = case_when(
    str_detect(tolower(sex), "^f") ~ "Female",
    str_detect(tolower(sex), "^m") ~ "Male",
    TRUE ~ NA_character_
  )) |>
  filter(!is.na(sex)) |>                         # drop "Persons" rows
  pivot_longer(all_of(age_cols), names_to = "age_single", values_to = "population") |>
  mutate(age_single = as.integer(age_single),
         population = as.numeric(population)) |>
  filter(age_single >= 16, !is.na(population), population > 0) |>
  mutate(age_band = case_when(
    age_single <= 34 ~ "16-34",
    age_single <= 54 ~ "35-54",
    TRUE             ~ "55+"
  )) |>
  group_by(constituency, sex, age_band) |>
  summarise(population = sum(population, na.rm = TRUE), .groups = "drop")

message("✓ Census: ", n_distinct(census$constituency), " constituencies loaded")

# Proportions within each constituency (used for poststratification weights)
census_props <- census |>
  group_by(constituency) |>
  mutate(prop = population / sum(population)) |>
  ungroup()


# =============================================================================
# 4. LOAD 2021 RESULTS
# =============================================================================
# data/results_2021.csv — extracted from CBP9230 (House of Commons Library).
# Columns: constituency | snp | con | lab | lib | grn | oth | rfm | alba |
#          total_valid | electorate | turnout
# Party columns are vote SHARES in % (e.g. 47.7, not 0.477).

results_2021 <- read_csv(
  file.path(DATA_DIR, "results_2021.csv"),
  show_col_types = FALSE
) |> clean_names()

message("✓ 2021 results: ", nrow(results_2021), " constituencies loaded")

# National 2021 constituency vote shares (official figures — used for swing calc)
NATIONAL_2021 <- c(snp = 47.7, con = 22.2, lab = 21.6, lib = 7.9,
                   grn = 0.6,  rfm = 0.0,  alba = 0.2, oth = 0.0)


# =============================================================================
# 5. POSTSTRATIFICATION
# =============================================================================
# We only apply demographic adjustment for parties IN the crosstabs (SNP, Con,
# Lab, LibDem). For parties not in crosstabs (Reform, Green, Alba) we apply
# pure national swing (demo_ratio = 1) — adjusting by a demographic profile
# we don't have data for would just be noise.
#
# We also cap the demographic ratio at 2.5 to prevent the multiplicative model
# from producing impossible constituency-level estimates (a constituency can't
# be 14× more LibDem than average — it just has a different age profile).

# ── Age marginals ─────────────────────────────────────────────────────────────
age_probs <- crosstab_avg |>
  filter(vote_type == "constituency",
         str_detect(tolower(demo_var), "^age")) |>
  # Deduplicate: if multiple demo_var labels match "age", average them
  group_by(party_key, age_band = category) |>
  summarise(p_age = mean(pct_avg, na.rm = TRUE), .groups = "drop")

# ── Sex marginals — top-level Male/Female only ────────────────────────────────
sex_probs <- crosstab_avg |>
  filter(vote_type == "constituency",
         str_detect(tolower(demo_var), "sex|gender")) |>
  filter(str_detect(tolower(category), "^(female|male|women|men)$")) |>
  mutate(sex = if_else(str_detect(tolower(category), "^(f|w)"), "Female", "Male")) |>
  # Deduplicate: if multiple demo_var labels match (e.g. "Sex" vs "Gender"),
  # take the mean — this is what caused the many-to-many join explosion
  group_by(party_key, sex) |>
  summarise(p_sex = mean(pct_avg, na.rm = TRUE), .groups = "drop")

# Parties with usable crosstab data
ct_parties <- intersect(unique(age_probs$party_key), unique(sex_probs$party_key))
message("Parties with crosstab demographic data: ", paste(ct_parties, collapse = ", "))
missing_from_ct <- setdiff(c("snp","con","lab","lib","rfm","grn","alba"), ct_parties)
if (length(missing_from_ct) > 0)
  message("No crosstab data for (will use national swing only): ",
          paste(missing_from_ct, collapse = ", "))

if (length(ct_parties) == 0) stop("No parties with both age AND sex crosstabs — check demo_var values.")

# ── Compute demographic estimate for crosstab parties only ────────────────────
# Use explicit join with relationship = "many-to-many" acknowledged but
# prevented by the deduplication above
postrat_raw <- census_props |>
  select(constituency, sex, age_band, prop) |>
  crossing(party_key = ct_parties) |>
  left_join(age_probs, by = c("party_key", "age_band")) |>
  left_join(sex_probs, by = c("party_key", "sex")) |>
  mutate(p_cell = (replace_na(p_age, 0) + replace_na(p_sex, 0)) / 2) |>
  group_by(constituency, party_key) |>
  summarise(p_demo = sum(prop * p_cell, na.rm = TRUE), .groups = "drop") |>
  group_by(constituency) |>
  mutate(p_demo = p_demo / sum(p_demo) * 100) |>
  ungroup()

# ── Demographic ratio — capped at 2.5 ─────────────────────────────────────────
# Ratio > 1 = constituency leans toward this party relative to national average.
# Cap prevents extreme values from dominating (e.g. LibDem 14× in rural seats).
MAX_DEMO_RATIO <- 2.5

national_demo <- postrat_raw |>
  group_by(party_key) |>
  summarise(p_national_demo = mean(p_demo), .groups = "drop")

demo_ratios <- postrat_raw |>
  left_join(national_demo, by = "party_key") |>
  mutate(
    raw_ratio  = if_else(p_national_demo > 0, p_demo / p_national_demo, 1),
    demo_ratio = pmin(raw_ratio, MAX_DEMO_RATIO)  # cap at 2.5
  ) |>
  select(constituency, party_key, demo_ratio)

# Check how many constituencies had ratios above cap (should be small)
n_capped <- postrat_raw |>
  left_join(national_demo, by = "party_key") |>
  mutate(raw_ratio = if_else(p_national_demo > 0, p_demo / p_national_demo, 1)) |>
  filter(raw_ratio > MAX_DEMO_RATIO) |>
  nrow()
if (n_capped > 0)
  message("  Capped ", n_capped, " ratio values above ", MAX_DEMO_RATIO,
          " (these were caused by small national average denominators)")


# =============================================================================
# 6. APPLY SWING + DEMOGRAPHIC ADJUSTMENT
# =============================================================================
# Crosstab parties (SNP, Con, Lab, LibDem): 2021 × swing_mult × demo_ratio
# Non-crosstab parties (Reform, Green, Alba, Other): 2021 × swing_mult only
# (demo_ratio = 1, equivalent to uniform national swing for those parties)
#
# National swing is derived from the headline polling averages, NOT the
# crosstab averages, so Reform and Green get their real headline numbers.

# ── National polling averages ─────────────────────────────────────────────────
if (!is.null(nat_avg) && nrow(nat_avg) > 0) {
  national_poll_avg <- nat_avg
  message("✓ Using headline poll averages for national swing")
} else {
  # Fallback: derive from crosstabs for parties we have, use NATIONAL_2021 for others
  message("⚠ Falling back to crosstab-derived averages")
  national_poll_avg <- age_probs |>
    mutate(weight = AGE_WEIGHTS[age_band]) |>
    group_by(party_key) |>
    summarise(p_poll = weighted.mean(p_age, weight, na.rm = TRUE), .groups = "drop") |>
    mutate(p_poll = p_poll / sum(p_poll) * 100) |>
    rename(p_national = p_poll)
}

# ── Additive swing in percentage points ───────────────────────────────────────
additive_swings <- national_poll_avg |>
  mutate(
    p_2021     = NATIONAL_2021[party_key],
    p_2021     = replace_na(p_2021, 0),
    swing_pp   = p_national - p_2021      # positive = party is up, negative = down
  ) |>
  select(party_key, swing_pp, p_national)

cat("\n=== Additive swings (pp vs 2021) ===\n")
print(additive_swings |> arrange(desc(swing_pp)))

# ── Pivot 2021 results to long format ─────────────────────────────────────────
results_long <- results_2021 |>
  pivot_longer(c(snp, con, lab, lib, grn, oth, rfm, alba),
               names_to = "party_key", values_to = "share_2021") |>
  select(constituency, party_key, share_2021)

# ── Combine: additive swing + demographic adjustment ──────────────────────────
# For parties in crosstabs: swing is scaled by demo_ratio (e.g. more LibDem
# swing in older constituencies). For others: pure national swing.
constituency_estimates <- results_long |>
  left_join(additive_swings, by = "party_key") |>
  left_join(demo_ratios, by = c("constituency", "party_key")) |>
  mutate(
    swing_pp   = replace_na(swing_pp, 0),
    # demo_ratio shifts swing magnitude: ratio>1 means more swing here than average
    demo_ratio = if_else(party_key %in% ct_parties, replace_na(demo_ratio, 1), 1),
    # Adjusted share: 2021 baseline + demographically-scaled swing
    share_adj  = share_2021 + swing_pp * demo_ratio,
    share_adj  = pmax(0, share_adj)   # floor at 0
  ) |>
  group_by(constituency) |>
  mutate(vote_share = share_adj / sum(share_adj, na.rm = TRUE) * 100) |>
  ungroup() |>
  select(constituency, party_key, vote_share)

# ── Sanity check: flag implausible estimates ──────────────────────────────────
cat("\n=== Sanity check: parties with >40% in constituencies where they were <15% in 2021 ===\n")
sanity_issues <- constituency_estimates |>
  left_join(results_long, by = c("constituency","party_key")) |>
  filter(share_2021 < 15, vote_share > 40) |>
  arrange(desc(vote_share))
if (nrow(sanity_issues) > 0) {
  print(sanity_issues)
  message("\n⚠  The above rows look implausible. Check demographic ratios and headline poll URLs.")
} else {
  message("✓ No implausible estimates detected.")
}


# =============================================================================
# 7. SEAT WINNERS + VALIDATION
# =============================================================================

seat_winners <- constituency_estimates |>
  group_by(constituency) |>
  slice_max(vote_share, n = 1, with_ties = FALSE) |>
  rename(winner = party_key, winner_share = vote_share) |>
  ungroup()

seat_summary <- seat_winners |>
  count(winner, name = "projected_seats") |>
  arrange(desc(projected_seats))

cat("\n=== Projected 2026 constituency seats ===\n")
print(seat_summary)
cat("Total:", sum(seat_summary$projected_seats), "of 73 constituencies\n")

# Sanity check: does the model correctly call 2021 winners?
actual_winners_2021 <- results_long |>
  group_by(constituency) |>
  slice_max(share_2021, n = 1, with_ties = FALSE) |>
  rename(actual_2021 = party_key) |>
  ungroup()

check <- seat_winners |>
  left_join(actual_winners_2021 |> select(constituency, actual_2021),
            by = "constituency") |>
  mutate(swing_changed = winner != actual_2021)

n_flipped <- sum(check$swing_changed, na.rm = TRUE)
cat("\nSeats projected to change hands vs 2021:", n_flipped, "\n")
cat("(If this looks too high or too low, adjust DEMOGRAPHIC_WEIGHT below)\n")

# Seats projected to flip — useful for review
if (n_flipped > 0) {
  cat("\nProjected flips:\n")
  check |>
    filter(swing_changed) |>
    select(constituency, actual_2021, projected_2026 = winner, winner_share) |>
    mutate(winner_share = round(winner_share, 1)) |>
    print(n = Inf)
}


# =============================================================================
# 8. EXPORT
# =============================================================================

# ── Wide CSV ──────────────────────────────────────────────────────────────────
results_wide <- constituency_estimates |>
  pivot_wider(names_from = party_key, values_from = vote_share, values_fill = 0) |>
  left_join(seat_winners |> select(constituency, projected_winner = winner, winner_share),
            by = "constituency") |>
  mutate(across(where(is.numeric), \(x) round(x, 2)))

write_csv(results_wide, file.path(OUTPUT_DIR, "constituency_estimates.csv"))
message("\n✓ Wrote: ", file.path(OUTPUT_DIR, "constituency_estimates.csv"))

# ── JSON for dashboard ────────────────────────────────────────────────────────
const_list <- constituency_estimates |>
  mutate(vote_share = round(vote_share, 2)) |>
  group_by(constituency) |>
  group_map(\(df, key) {
    shares     <- setNames(as.list(df$vote_share), df$party_key)
    winner_row <- df |> slice_max(vote_share, n = 1, with_ties = FALSE)
    c(shares, list(projected_winner = winner_row$party_key,
                   winner_share     = round(winner_row$vote_share, 2)))
  }) |>
  setNames(results_wide$constituency)

json_out <- list(
  meta = list(
    generated    = format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ", tz = "UTC"),
    method       = "Additive swing with demographic adjustment",
    method_note  = paste0(
      "NOT full MRP. Starts from 2021 constituency results (CBP9230), applies ",
      "additive national polling swing, then adjusts each constituency by its ",
      "NRS Census 2022 age×sex demographic profile via marginal poststratification. ",
      "National polling averages from dashboard headline strip (MANUAL_NATIONAL_POLL)."
    ),
    nrs_year       = nrs_sheet,
    constituencies = nrow(results_wide),
    seats_summary  = as.list(setNames(seat_summary$projected_seats, seat_summary$winner))
  ),
  constituencies = const_list
)

json_path <- file.path(OUTPUT_DIR, "constituency_estimates.json")
write_json(json_out, json_path, auto_unbox = TRUE, pretty = TRUE)
message("✓ Wrote: ", json_path)

cat("\n", strrep("=", 60), "\n")
cat("NEXT: upload constituency_estimates.json to GitHub Gist,\n")
cat("      click Raw, paste the URL into SOURCES.postrat.url\n")
cat("      in holyrood-2026.html, then refresh the dashboard.\n")
cat(strrep("=", 60), "\n")
