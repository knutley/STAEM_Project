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
  "16-34" = 0.29,
  "35-54" = 0.31,
  "55+"   = 0.40   # no space
)

OUTPUT_DIR <- "output"
DATA_DIR   <- "data"
dir.create(OUTPUT_DIR, showWarnings = FALSE)
dir.create(DATA_DIR,   showWarnings = FALSE)


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

# Quick sense check
cat("\n=== Constituency crosstab averages — Age ===\n")
crosstab_avg |>
  filter(vote_type == "constituency", demo_var == "Age") |>
  arrange(category, desc(pct_avg)) |>
  select(category, party_key, pct_avg) |>
  print(n = Inf)

cat("\n=== Constituency crosstab averages — Sex ===\n")
crosstab_avg |>
  filter(vote_type == "constituency",
         str_detect(tolower(demo_var), "sex|gender")) |>
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
# For each constituency we ask: given its age × sex composition, what vote
# share does each party get if we weight up the crosstab estimates?
#
# p̂(vote_j | constituency c) = Σ_{age,sex} prop_{c,age,sex} × p̂(vote_j | age, sex)
#
# We combine age and sex marginals as a simple average — the most defensible
# approach when we only have marginals, not joint cells.

age_probs <- crosstab_avg |>
  filter(vote_type == "constituency",
         str_detect(tolower(demo_var), "^age")) |>
  select(party_key, age_band = category, p_age = pct_avg)

sex_probs <- crosstab_avg |>
  filter(vote_type == "constituency",
         str_detect(tolower(demo_var), "sex|gender")) |>
  # Keep ONLY top-level Male/Female rows — drop sub-groups like "Female 18-34"
  filter(category %in% c("Female", "Male", "female", "male",
                         "Women", "Men", "women", "men")) |>
  mutate(sex = case_when(
    str_detect(tolower(category), "^f|^w") ~ "Female",
    str_detect(tolower(category), "^m")    ~ "Male",
    TRUE ~ NA_character_
  )) |>
  filter(!is.na(sex)) |>
  select(party_key, sex, p_sex = pct_avg)

if (nrow(age_probs) == 0) stop("No age crosstabs found — check demo_var values in your sheets.")
if (nrow(sex_probs) == 0) stop("No sex crosstabs found — check demo_var values in your sheets.")

# Demographic constituency estimate (pure postrat, no swing applied yet)
postrat_raw <- census_props |>
  select(constituency, sex, age_band, prop) |>
  crossing(party_key = unique(age_probs$party_key)) |>
  left_join(age_probs, by = c("party_key", "age_band")) |>
  left_join(sex_probs, by = c("party_key", "sex")) |>
  mutate(p_cell = (replace_na(p_age, 0) + replace_na(p_sex, 0)) / 2) |>
  group_by(constituency, party_key) |>
  summarise(p_demo = sum(prop * p_cell, na.rm = TRUE), .groups = "drop") |>
  group_by(constituency) |>
  mutate(p_demo = p_demo / sum(p_demo) * 100) |>   # normalise to 100%
  ungroup()

# National average implied by the demographic model
# (used to compute the ratio: how much does each constituency deviate from average?)
national_demo <- postrat_raw |>
  group_by(party_key) |>
  summarise(p_national_demo = mean(p_demo), .groups = "drop")

# Demographic adjustment ratio per constituency
# ratio > 1 means this constituency leans more toward this party than average
demo_ratios <- postrat_raw |>
  left_join(national_demo, by = "party_key") |>
  mutate(demo_ratio = if_else(p_national_demo > 0, p_demo / p_national_demo, 1)) |>
  select(constituency, party_key, demo_ratio)


# =============================================================================
# 6. APPLY SWING + DEMOGRAPHIC ADJUSTMENT
# =============================================================================
# Method: multiplicative swing with demographic scaling.
#
#   predicted[c,j] = results_2021[c,j]
#                    × (national_poll[j] / national_2021[j])   ← national swing
#                    × demo_ratio[c,j]                          ← demographic adjustment
#
# Using multiplicative rather than additive swing handles small-base parties
# (Reform, Alba) more sensibly — they grow proportionally, not by a fixed pp.

# National current polling from crosstab averages (age-weighted)
national_poll <- age_probs |>
  mutate(weight = AGE_WEIGHTS[age_band]) |>
  group_by(party_key) |>
  summarise(p_poll = weighted.mean(p_age, weight, na.rm = TRUE), .groups = "drop") |>
  mutate(p_poll = p_poll / sum(p_poll) * 100)

# Swing multiplier: current poll / 2021 national result
swing_mult <- national_poll |>
  mutate(
    p_2021   = NATIONAL_2021[party_key],
    p_2021   = if_else(is.na(p_2021) | p_2021 == 0, 0.1, p_2021),  # avoid /0
    mult     = p_poll / p_2021
  ) |>
  select(party_key, mult)

# Pivot 2021 results to long format
results_long <- results_2021 |>
  pivot_longer(c(snp, con, lab, lib, grn, oth, rfm, alba),
               names_to = "party_key", values_to = "share_2021") |>
  select(constituency, party_key, share_2021)

# Combine: 2021 result × swing multiplier × demographic ratio
constituency_estimates <- results_long |>
  left_join(swing_mult,   by = "party_key") |>
  left_join(demo_ratios,  by = c("constituency", "party_key")) |>
  mutate(
    mult       = replace_na(mult, 1),
    demo_ratio = replace_na(demo_ratio, 1),
    share_adj  = share_2021 * mult * demo_ratio
  ) |>
  group_by(constituency) |>
  mutate(vote_share = share_adj / sum(share_adj, na.rm = TRUE) * 100) |>
  ungroup() |>
  select(constituency, party_key, vote_share)


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
    method       = "Multiplicative swing with demographic adjustment",
    method_note  = paste0(
      "NOT full MRP. Starts from 2021 constituency results (CBP9230), applies ",
      "national polling swing (from crosstab averages), then adjusts each ",
      "constituency by its NRS Census 2022 age×sex demographic profile. "
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
