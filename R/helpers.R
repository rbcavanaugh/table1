library(flextable)
library(officer)
library(table1)

# ── APA Flextable Theme ────────────────────────────────────────────────────────

theme_apa2 <- function(.data, pad = 0, spacing = 1) {
  apa.border <- list("width" = 1, color = "black", style = "solid")
  font(.data, part = "all", fontname = "Times New Roman") |>
    line_spacing(space = spacing, part = "all") |>
    padding(padding = pad) |>
    hline_top(part = "head", border = apa.border) |>
    hline_bottom(part = "head", border = apa.border) |>
    hline_top(part = "body", border = apa.border) |>
    hline_bottom(part = "body", border = apa.border) |>
    align(align = "center", part = "all", j = -1) |>
    valign(valign = "center", part = "all") |>
    colformat_double(digits = 2) |>
    italic(part = "footer") |>
    hline_bottom(part = "footer", border = fp_border(width = 0, color = "transparent")) |>
    fix_border_issues()
}

apa_caption <- function(text) {
  as_paragraph(as_chunk(text, props = fp_text_default(font.family = "Times New Roman")))
}

# ── table1 Render Function Factories ──────────────────────────────────────────

# Returns a render.continuous function for table1.
#
# For single-stat modes (mean_sd, median_iqr) the function returns a plain
# string with no leading "" element. A leading "" creates an extra blank row
# above the value in the table; a bare string renders on the variable's own row.
# For "both", named elements create labelled sub-rows (requires a leading "").
render_continuous <- function(type) {
  switch(type,
    mean_sd = function(x, ...) {
      sprintf("%.2f (%.2f)", mean(x, na.rm = TRUE), sd(x, na.rm = TRUE))
    },
    median_iqr = function(x, ...) {
      sprintf("%.2f [%.2f, %.2f]",
        median(x, na.rm = TRUE),
        quantile(x, 0.25, na.rm = TRUE),
        quantile(x, 0.75, na.rm = TRUE))
    },
    both = function(x, ...) {
      c(
        "",
        `Mean (SD)`    = sprintf("%.2f (%.2f)", mean(x, na.rm = TRUE), sd(x, na.rm = TRUE)),
        `Median [IQR]` = sprintf("%.2f [%.2f, %.2f]",
          median(x, na.rm = TRUE),
          quantile(x, 0.25, na.rm = TRUE),
          quantile(x, 0.75, na.rm = TRUE))
      )
    }
  )
}

# Returns a render.categorical function for table1.
#
# IMPORTANT: do NOT call factor(x) unconditionally. R preserves factor levels
# when subsetting a data frame, so x arrives here as a factor with ALL levels
# from the full dataset (including levels with 0 observations in this group).
# Calling factor(x) would coerce through as.character() first, silently
# dropping zero-count levels and producing vectors of inconsistent length
# across groups — the root cause of the dimnames error.
render_categorical <- function(type) {
  switch(type,
    n_pct = function(x, ...) {
      if (!is.factor(x)) x <- factor(x)  # only re-factor if not already a factor
      f <- table(x)                        # includes 0-count levels for factors
      n <- sum(!is.na(x))
      vals <- if (n > 0) {
        sprintf("%d (%0.1f%%)", as.integer(f), 100 * as.numeric(f) / n)
      } else {
        rep("0 (0.0%)", length(f))
      }
      c("", setNames(vals, names(f)))
    },
    pct = function(x, ...) {
      if (!is.factor(x)) x <- factor(x)
      f <- table(x)
      n <- sum(!is.na(x))
      vals <- if (n > 0) {
        sprintf("%0.1f%%", 100 * as.numeric(f) / n)
      } else {
        rep("0.0%", length(f))
      }
      c("", setNames(vals, names(f)))
    },
    n = function(x, ...) {
      if (!is.factor(x)) x <- factor(x)
      f <- table(x)
      c("", setNames(sprintf("%d", as.integer(f)), names(f)))
    }
  )
}

# ── Data Quality Checks ───────────────────────────────────────────────────────

# Returns a character vector of plain-English warning messages (length 0 if
# the data looks clean). Each message is self-contained and actionable.
check_data <- function(df) {
  msgs <- character(0)

  nms <- names(df)

  # 1. Completely empty rows (every column NA)
  empty_rows <- which(rowSums(!is.na(df)) == 0)
  if (length(empty_rows) > 0) {
    shown <- if (length(empty_rows) <= 10) {
      paste(empty_rows, collapse = ", ")
    } else {
      paste0(paste(head(empty_rows, 10), collapse = ", "), ", \u2026")
    }
    msgs <- c(msgs, sprintf(
      "Row%s %s completely empty (all values missing): %s. Remove %s before uploading.",
      if (length(empty_rows) == 1) "" else "s",
      if (length(empty_rows) == 1) "is" else "are",
      shown,
      if (length(empty_rows) == 1) "it" else "them"
    ))
  }

  # 2. Column names starting with a digit
  num_start <- nms[grepl("^[0-9]", nms)]
  if (length(num_start) > 0) {
    msgs <- c(msgs, sprintf(
      "Column name%s starting with a number: %s. R requires names to begin with a letter \u2014 rename them (e.g. \u201cvisit1\u201d instead of \u201c1visit\u201d).",
      if (length(num_start) == 1) "" else "s",
      paste(sprintf('"%s"', num_start), collapse = ", ")
    ))
  }

  # 3. Column names with spaces or special characters (beyond letters/digits/._-)
  bad_chars <- nms[grepl("[^A-Za-z0-9_.\\-]", nms) & !grepl("^[0-9]", nms)]
  if (length(bad_chars) > 0) {
    msgs <- c(msgs, sprintf(
      "Column name%s with spaces or special characters: %s. Rename using only letters, numbers, and underscores to avoid errors.",
      if (length(bad_chars) == 1) "" else "s",
      paste(sprintf('"%s"', bad_chars), collapse = ", ")
    ))
  }

  # 4. Duplicate rows
  n_dup <- sum(duplicated(df))
  if (n_dup > 0) {
    msgs <- c(msgs, sprintf(
      "%d duplicate row%s detected. Each row should represent a unique participant \u2014 double-check for accidental copy-pastes.",
      n_dup, if (n_dup == 1) "" else "s"
    ))
  }

  # 5. Columns with no data at all
  all_na <- nms[sapply(df, function(x) all(is.na(x)))]
  if (length(all_na) > 0) {
    msgs <- c(msgs, sprintf(
      "Column%s with no data at all: %s. %s either empty or failed to import correctly.",
      if (length(all_na) == 1) "" else "s",
      paste(sprintf('"%s"', all_na), collapse = ", "),
      if (length(all_na) == 1) "It is" else "They are"
    ))
  }

  # 6. Columns with more than 50 % missing (but not 100 %)
  high_miss <- nms[sapply(df, function(x) {
    p <- mean(is.na(x)); p > 0.5 && p < 1.0
  })]
  if (length(high_miss) > 0) {
    pcts <- sapply(df[high_miss], function(x) round(100 * mean(is.na(x))))
    detail <- paste(sprintf('"%s" (%d%%)', high_miss, pcts), collapse = ", ")
    msgs <- c(msgs, sprintf(
      "Column%s with more than half of values missing: %s. Confirm this is expected.",
      if (length(high_miss) == 1) "" else "s", detail
    ))
  }

  # 7. Likely ID columns: non-numeric, every value unique, more than 10 rows
  likely_id <- nms[sapply(nms, function(v) {
    x    <- df[[v]]
    vals <- na.omit(x)
    !is.numeric(x) && length(vals) > 10 && length(unique(vals)) == length(vals)
  })]
  if (length(likely_id) > 0) {
    msgs <- c(msgs, sprintf(
      "Column%s where every value is unique: %s. %s probably an ID or name column that should be removed before generating Table 1.",
      if (length(likely_id) == 1) "" else "s",
      paste(sprintf('"%s"', likely_id), collapse = ", "),
      if (length(likely_id) == 1) "It is" else "They are"
    ))
  }

  # 8. Numeric columns with very few unique values (likely 0/1 codes)
  coded_num <- nms[sapply(nms, function(v) {
    x <- df[[v]]
    is.numeric(x) && length(unique(na.omit(x))) <= 5
  })]
  if (length(coded_num) > 0) {
    msgs <- c(msgs, sprintf(
      "Numeric column%s with very few unique values: %s. If %s a coded categorical variable (e.g. 0/1 for No/Yes), set its type to \u201cCategorical\u201d on the Configure tab.",
      if (length(coded_num) == 1) "" else "s",
      paste(sprintf('"%s"', coded_num), collapse = ", "),
      if (length(coded_num) == 1) "it is" else "they are"
    ))
  }

  # 9. Leading / trailing whitespace in character columns
  ws_cols <- nms[sapply(df, function(x) {
    is.character(x) && any(!is.na(x) & x != trimws(x))
  })]
  if (length(ws_cols) > 0) {
    msgs <- c(msgs, sprintf(
      "Column%s with extra spaces around values: %s. This often happens when pasting from Excel and can cause categories to be split (e.g. \u201cYes\u201d \u2260 \u201cYes \u201d).",
      if (length(ws_cols) == 1) "" else "s",
      paste(sprintf('"%s"', ws_cols), collapse = ", ")
    ))
  }

  # 10. Mixed capitalisation within a column (e.g. "Yes" and "yes" treated as different)
  mixed_case <- nms[sapply(df, function(x) {
    if (is.numeric(x)) return(FALSE)
    vals <- unique(na.omit(as.character(x)))
    anyDuplicated(tolower(vals)) > 0
  })]
  if (length(mixed_case) > 0) {
    msgs <- c(msgs, sprintf(
      "Column%s with inconsistent capitalisation: %s. Values like \u201cYes\u201d and \u201cyes\u201d will be treated as different categories \u2014 standardise them first.",
      if (length(mixed_case) == 1) "" else "s",
      paste(sprintf('"%s"', mixed_case), collapse = ", ")
    ))
  }

  msgs
}

# ── Suggestion Helpers ─────────────────────────────────────────────────────────

# Returns column names suitable for use as a stratification variable:
# non-numeric, between 2 and 10 unique non-NA values.
suggest_strata <- function(df) {
  is_candidate <- sapply(names(df), function(col) {
    x <- df[[col]]
    n_unique <- length(unique(na.omit(x)))
    !is.numeric(x) && n_unique >= 2 && n_unique <= 10
  })
  names(df)[is_candidate]
}

# Returns "mean_sd" or "median_iqr" for a numeric vector based on a Shapiro-Wilk
# normality test. Returns NULL for non-numeric input.
suggest_stats_for_var <- function(x) {
  if (!is.numeric(x)) return(NULL)
  x_clean <- na.omit(x)
  if (length(x_clean) < 3) return("mean_sd")
  n <- min(5000, length(x_clean))
  result <- tryCatch(
    shapiro.test(sample(x_clean, n)),
    error = function(e) NULL
  )
  if (is.null(result)) return("mean_sd")
  if (result$p.value <= 0.05) "median_iqr" else "mean_sd"
}

# ── Utilities ──────────────────────────────────────────────────────────────────

# Converts a column name to a valid Shiny input ID (letters, digits, underscores
# and hyphens only).
safe_id <- function(v) gsub("[^A-Za-z0-9_-]", "_", v)

# Null-coalescing operator.
`%||%` <- function(x, y) if (!is.null(x)) x else y
