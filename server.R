library(shiny)
library(dplyr)
library(readr)
library(readxl)
library(DT)
library(table1)
library(flextable)
library(officer)

server <- function(input, output, session) {

  # ── Data Loading ─────────────────────────────────────────────────────────────

  data_raw <- reactive({
    req(input$file)
    ext <- tolower(tools::file_ext(input$file$name))
    tryCatch(
      switch(ext,
        csv  = read_csv(input$file$datapath, show_col_types = FALSE),
        tsv  = read_tsv(input$file$datapath, show_col_types = FALSE),
        xlsx = read_excel(input$file$datapath),
        xls  = read_excel(input$file$datapath),
        {
          showNotification("Unsupported file type. Please upload a CSV, TSV, or Excel file.",
            type = "error")
          NULL
        }
      ),
      error = function(e) {
        showNotification(paste("Could not read file:", e$message), type = "error")
        NULL
      }
    )
  })

  output$upload_status <- renderUI({
    df <- data_raw()
    if (is.null(df)) return(NULL)
    div(class = "hint-box",
      strong(nrow(df)), " rows \u00d7 ", strong(ncol(df)), " columns loaded successfully. ",
      actionLink("go_to_config", "Next: Configure Table \u2192")
    )
  })

  observeEvent(input$go_to_config, {
    updateTabsetPanel(session, "main_tabs", selected = "2. Configure Table")
  })

  output$data_warnings <- renderUI({
    req(data_raw())
    msgs <- check_data(data_raw())
    if (length(msgs) == 0) return(NULL)
    tagList(lapply(msgs, function(msg) {
      div(class = "warn-box",
        tags$b("\u26a0 Check your data: "), msg
      )
    }))
  })

  output$data_preview <- renderDT({
    req(data_raw())
    datatable(data_raw(),
      options = list(scrollX = TRUE, pageLength = 8, dom = "tip"),
      rownames = FALSE
    )
  })

  observeEvent(input$show_tidy_help, {
    showModal(modalDialog(
      tags$img(
        src   = "https://cdn.myportfolio.com/45214904-6a61-4e23-98d6-b140f8654a40/85520b8f-4629-4763-8a2a-9ceff27458bf.jpg?h=27c3e74642d8602799d180ff30995d2a",
        style = "width:100%; height:auto;"
      ),
      tags$p(
        style = "color:#999; font-style:italic; text-align:center; font-size:0.8em; margin-top:10px;",
        "Artwork by Allison Horst. ",
        tags$a("allisonhorst.com/data-science-art",
          href   = "https://allisonhorst.com/data-science-art",
          target = "_blank",
          style  = "color:#999;"
        )
      ),
      title     = "Tidy Data",
      easyClose = TRUE,
      footer    = modalButton("Close"),
      size      = "l"
    ))
  })

  output$download_example <- downloadHandler(
    filename = "example_table1_data.csv",
    content = function(file) {
      set.seed(42)
      n <- 120
      df <- data.frame(
        age         = round(rnorm(n, 45, 12)),
        sex         = sample(c("Male", "Female"), n, replace = TRUE),
        group       = sample(c("Control", "Treatment"), n, replace = TRUE),
        bmi         = round(rnorm(n, 26.5, 4.5), 1),
        smoker      = sample(c("Yes", "No"), n, replace = TRUE, prob = c(0.25, 0.75)),
        systolic_bp = round(rnorm(n, 125, 18)),
        education   = sample(
          c("High School", "College", "Graduate"), n,
          replace = TRUE, prob = c(0.3, 0.45, 0.25)
        )
      )
      df$bmi[sample(n, 8)]         <- NA
      df$systolic_bp[sample(n, 5)] <- NA
      write_csv(df, file)
    }
  )

  # ── Configure: Strata + Options ───────────────────────────────────────────────

  output$strata_select_ui <- renderUI({
    req(data_raw())
    # Use a sentinel string rather than "" so "None" is always a concrete,
    # selectable value. Shiny can fail to fire change events when "" is
    # re-selected after a non-empty value was chosen.
    choices <- c("None" = ".none.", names(data_raw()))
    tagList(
      selectInput("strata_var", "Stratification variable:",
        choices = choices, selected = ".none."),
      p(style = "font-size:0.82em; color:#666; margin-top:-6px;",
        "Researchers often stratify by a grouping variable (e.g., treatment vs. control,",
        "or case vs. control) when the central research question compares those groups.",
        "If you simply want to describe the full sample, leave this as None."
      )
    )
  })

  output$overall_ui <- renderUI({
    if (!is.null(input$strata_var) && input$strata_var != ".none.") {
      checkboxInput("show_overall", "Include 'Overall' column", value = TRUE)
    }
  })

  output$suggestions_ui <- renderUI({
    req(data_raw())
    df <- data_raw()
    hints <- list()

    # Normality-based stats suggestion
    num_vars <- names(df)[sapply(df, is.numeric)]
    non_normal <- Filter(function(v) {
      s <- suggest_stats_for_var(df[[v]])
      !is.null(s) && s == "median_iqr"
    }, num_vars)

    if (length(non_normal) > 0) {
      hints[[length(hints) + 1]] <- div(class = "warn-box",
        tags$b("Statistics: "),
        "Possible non-normal distribution detected in ",
        tags$b(paste(non_normal, collapse = ", ")),
        ". Consider Median [IQR]."
      )
    }

    if (length(hints) == 0) return(NULL)
    tagList(h5("Suggestions"), hints)
  })

  # ── Configure: Variable Selection & Labeling ──────────────────────────────────

  # Variables available for the table (excludes the strata column when selected)
  available_vars <- reactive({
    req(data_raw())
    strata <- input$strata_var %||% ".none."
    setdiff(names(data_raw()), if (strata != ".none.") strata else character(0))
  })

  output$variable_config_ui <- renderUI({
    vars <- available_vars()
    df   <- data_raw()

    lapply(vars, function(v) {
      col     <- df[[v]]
      n_miss  <- sum(is.na(col))
      auto_type <- if (is.numeric(col)) "Continuous" else "Categorical"
      sid     <- safe_id(v)

      div(class = "var-row",
        fluidRow(
          column(1, style = "padding-top:6px;",
            checkboxInput(paste0("inc_", sid), label = NULL, value = TRUE)
          ),
          column(3,
            tags$b(v), tags$br(),
            tags$small(style = "color:#888;",
              auto_type,
              if (n_miss > 0) paste0(", ", n_miss, " missing") else NULL
            )
          ),
          column(5,
            textInput(paste0("lbl_", sid), label = NULL,
              value = v, placeholder = "Display label...")
          ),
          column(3,
            selectInput(paste0("typ_", sid), label = NULL,
              choices  = c("Auto", "Continuous", "Categorical"),
              selected = "Auto",
              width    = "100%"
            )
          )
        )
      )
    })
  })

  # ── Table Generation ──────────────────────────────────────────────────────────

  observeEvent(input$go_preview, {
    updateTabsetPanel(session, "main_tabs", selected = "3. Preview & Export")
  })

  table_built <- eventReactive(input$go_preview, {
    df         <- data_raw()
    vars       <- isolate(available_vars())
    strata_raw <- isolate(input$strata_var) %||% ".none."
    strata_var <- if (strata_raw == ".none.") "" else strata_raw

    # Filter to included variables
    included <- Filter(function(v) {
      isTRUE(isolate(input[[paste0("inc_", safe_id(v))]]))
    }, vars)

    if (length(included) == 0) {
      showNotification("Select at least one variable to include.", type = "warning")
      return(NULL)
    }

    # Use a plain data.frame (not tibble) so that column assignment preserves
    # custom attributes like the table1 label.
    df_lab <- as.data.frame(df)

    for (v in included) {
      sid     <- safe_id(v)
      lbl     <- trimws(isolate(input[[paste0("lbl_", sid)]]) %||% v)
      type_ov <- isolate(input[[paste0("typ_", sid)]]) %||% "Auto"

      col_data <- df_lab[[v]]

      # Coerce type FIRST — factor() and as.numeric() strip all attributes
      # including the table1 label, so the label must be applied afterwards.
      if (type_ov == "Continuous") {
        col_data <- suppressWarnings(as.numeric(col_data))
      } else if (type_ov == "Categorical" || (type_ov == "Auto" && !is.numeric(col_data))) {
        # Convert to factor using the full dataset so every level is represented.
        # R preserves factor levels when subsetting a data frame, which means
        # the render functions will receive a factor with ALL levels (even those
        # with 0 observations in a given strata group), keeping rendered vector
        # lengths consistent across groups.
        col_data <- factor(col_data)
      }

      # Apply the display label after coercion.
      if (nchar(lbl) > 0) label(col_data) <- lbl

      df_lab[[v]] <- col_data
    }

    # Build formula — backtick-quote names so columns with dots/dashes/spaces
    # are not misread as arithmetic operators by the formula parser.
    rhs <- paste(sprintf("`%s`", included), collapse = " + ")
    fml <- if (nchar(strata_var) > 0) {
      as.formula(paste("~", rhs, "|", sprintf("`%s`", strata_var)))
    } else {
      as.formula(paste("~", rhs))
    }

    # Overall column label.
    # When there is no strata, table1 produces exactly one column. Setting
    # overall = FALSE in that case removes it entirely → "Table has no columns".
    # Only suppress the Overall column when there IS stratification and the
    # user has unchecked the box.
    overall_label <- if (nchar(strata_var) == 0) {
      "Overall"
    } else if (isTRUE(isolate(input$show_overall))) {
      "Overall"
    } else {
      FALSE
    }

    # Render functions
    cont_fn <- render_continuous(isolate(input$cont_stats) %||% "both")
    cat_fn  <- render_categorical(isolate(input$cat_stats) %||% "n_pct")
    miss_fn <- if (isTRUE(isolate(input$show_missing))) render.missing.default else NULL

    tryCatch(
      table1(fml,
        data               = df_lab,
        render.continuous  = cont_fn,
        render.categorical = cat_fn,
        render.missing     = miss_fn,
        overall            = overall_label
      ),
      error = function(e) {
        showNotification(paste("Table generation failed:", e$message), type = "error")
        NULL
      }
    )
  })

  output$table1_output <- renderUI({
    tbl <- tryCatch(table_built(), error = function(e) NULL)
    if (is.null(tbl)) {
      return(p(class = "text-muted",
        "Configure your table on the previous tab and click 'Generate Table' to see a preview."
      ))
    }
    HTML(as.character(tbl))
  })

  # ── Word Export ───────────────────────────────────────────────────────────────

  output$download_word <- downloadHandler(
    filename = function() {
      paste0("table1_", format(Sys.Date(), "%Y%m%d"), ".docx")
    },
    content = function(file) {
      tbl <- tryCatch(table_built(), error = function(e) NULL)
      if (is.null(tbl)) {
        showNotification("Generate the table before downloading.", type = "warning")
        return()
      }

      caption <- trimws(input$table_caption %||% "")
      note    <- trimws(input$table_note    %||% "")

      ft <- t1flex(tbl) |>
        set_table_properties(align = "left") |>
        autofit(add_w = 0, add_h = 1)

      if (nchar(caption) > 0) {
        ft <- set_caption(ft,
          caption           = apa_caption(caption),
          align_with_table  = TRUE,
          word_stylename    = "Normal",
          fp_p              = fp_par(text.align = "left")
        )
      }

      if (nchar(note) > 0) {
        ft <- add_footer_lines(ft, values = note)
      }

      ft <- theme_apa2(ft)

      doc <- read_docx()
      doc <- body_add_flextable(doc, value = ft)
      print(doc, target = file)
    }
  )
}
