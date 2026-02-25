# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This is an R Shiny application for generating "Table 1" — a demographics/baseline characteristics table commonly used in clinical manuscripts. The goal is to take clean data as input and produce a formatted demographics table.

## Running the App

```r
# From R or RStudio
shiny::runApp()

# Or run directly
Rscript -e "shiny::runApp('.')"
```

The app can also be launched in RStudio via the "Run App" button when `app.R` is open.

## Project Structure

- `app.R` — thin entry point; sources `ui.R` and `server.R`, calls `shinyApp()`
- `ui.R` — defines the `ui` object (3-tab layout)
- `server.R` — defines the `server` function (data loading, table generation, Word export)
- `R/helpers.R` — auto-sourced by Shiny; contains `theme_apa2`, `apa_caption`, `render_continuous`, `render_categorical`, `suggest_strata`, `suggest_stats_for_var`, `safe_id`

The `.Rproj` is configured to use 2-space indentation and UTF-8 encoding.

## R/Shiny Conventions

- 2 spaces for indentation (per `.Rproj` settings)
- Separate `ui.R` / `server.R` files

## Necessary features
- user input to provide labels that are formatted nicely using the table1 package (e.g., so `age` can be labelled as `Age (years)`)
- suggestions on when to stratify or include a single column. 
- suggestions on which statistics to use
- ability to choose which statistics are used for categorical variables and which summary statistics are used for continuous variables
- ability to exclude or include missing data
- ability to include or exclude the overall column
- guidance on what format to upload the data in (we will expect clean and ready to go data, but we need to provide an example of what it should look like)

## Relevant R Packages

The app is likely to use packages such as:
- `shiny` — web framework
- `table1`, for generating demographics tables
- `dplyr` / `tidyverse` / `readr`/ `readxl` — reading in data and data manipulation
- `DT` — interactive tables in Shiny for examining data that has been uploaded
= `flextable` and `officer` for allowing a download of the table in MS word

## Existing helpful code

```

# theming for apa
theme_apa2 <- function(.data, pad = 0, spacing = 2) {
  apa.border <- list("width" = 1, color = "black", style = "solid")
  font(.data, part = "all", fontname = "Times New Roman") |>
    line_spacing(space = spacing, part = "all") |>
    padding(padding=pad) |>
    hline_top(part = "head", border = apa.border) |>
    hline_bottom(part = "head", border = apa.border) |>
    hline_top(part = "body", border = apa.border) |>
    hline_bottom(part = "body", border = apa.border) |>
    align(align = "center", part = "all", j=-1) |>
    valign(valign = "center", part = "all") |>
    colformat_double(digits = 2) |>
    # Add footer formatting
    italic(part = "footer") |>
    hline_bottom(part = "footer", border = fp_border(width = 0, color = "transparent")) |>
    fix_border_issues()
}

# Set global defaults for all flextables
apa_caption <- function(text) {
  as_paragraph(as_chunk(text, props = fp_text_default(font.family = "Times New Roman")))
}

# example from another project
manuscript_table = icc_table |> 
  select(Stimuli = stimuli, Group = dx, Dependability = icc, `95% CI` = `ci_95`) |> 
  flextable() |> 
  set_table_properties(align = "left") |>   # Make sure table is left-aligned
  flextable::set_caption(apa_caption("Table X. Dependability Coefficients by Stimuli and Group"), align_with_table = TRUE,
                         word_stylename = "Normal",
                         fp_p = fp_par(text.align = "left")) |>
  flextable::autofit(add_w = 0, add_h = 1) |>
  flextable::add_footer_lines(values = "Note: Dependability coefficients are analogous to intraclass correlations for agreement") |>
  theme_apa2()

# To save it as a word document:
doc <- officer::read_docx()
# add the table to the document
doc <- body_add_flextable(doc, value = manuscript_table)
# filename within the path we want
print(doc, target = "output/icc_by_stimuli_and_group_revision.docx")

```
