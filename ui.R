library(shiny)
library(DT)

ui <- fluidPage(
  titlePanel("Table 1 Generator"),

  tags$script(HTML("
    Shiny.addCustomMessageHandler('download_docx', function(data) {
      var byteChars = atob(data.base64);
      var byteNums  = new Uint8Array(byteChars.length);
      for (var i = 0; i < byteChars.length; i++) {
        byteNums[i] = byteChars.charCodeAt(i);
      }
      var blob = new Blob([byteNums], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      var url = URL.createObjectURL(blob);
      var a   = document.createElement('a');
      a.href  = url;
      a.download = data.filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });
  ")),

  tags$style(HTML("
    .var-row {
      border: 1px solid #ddd;
      border-radius: 4px;
      padding: 8px 10px;
      margin-bottom: 6px;
      background: #fafafa;
    }
    .hint-box {
      background: #e8f4fd;
      border-left: 3px solid #3498db;
      padding: 8px 12px;
      margin: 8px 0;
      font-size: 0.85em;
      border-radius: 2px;
    }
    .warn-box {
      background: #fef9e7;
      border-left: 3px solid #f39c12;
      padding: 8px 12px;
      margin: 8px 0;
      font-size: 0.85em;
      border-radius: 2px;
    }
    .config-header { font-weight: bold; color: #555; font-size: 0.8em; text-transform: uppercase; margin-bottom: 4px; }
  ")),

  tabsetPanel(
    id = "main_tabs",

    # ── Tab 1: Upload ────────────────────────────────────────────────────────
    tabPanel("1. Upload Data",
      br(),
      fluidRow(
        column(4,
          wellPanel(
            h4("Upload your data"),
            fileInput("file", NULL,
              accept = c(".csv", ".tsv", ".xlsx", ".xls"),
              placeholder = "No file selected"
            ),
            hr(),
            h5("Expected format: ", actionLink("show_tidy_help", "Tidy Data")),
            tags$ul(
              tags$li("One row per participant/observation"),
              tags$li("One column per variable"),
              tags$li("Only include columns for variables that you want to include in the table"),
              tags$li("Column names without spaces or special characters"),
              tags$li("A grouping column if you plan to stratify (e.g., ", tags$code("group"), ")")
            ),
            p(class = "text-muted", style = "font-size:0.85em;",
              "Clean, analysis-ready data only. No formulas, merged cells, or multi-row headers."
            ),
            hr(),
            downloadButton("download_example", "Download Example Dataset",
              style = "width:100%"
            )
          )
        ),
        column(8,
          uiOutput("table1_explainer"),
          uiOutput("upload_status"),
          uiOutput("data_warnings"),
          DTOutput("data_preview")
        )
      )
    ),

    # ── Tab 2: Configure ─────────────────────────────────────────────────────
    tabPanel("2. Configure Table",
      br(),
      fluidRow(
        column(4,
          wellPanel(
            h4("Table Structure"),
            uiOutput("strata_select_ui"),
            uiOutput("overall_ui"),
            hr(),
            h4("Statistics"),
            radioButtons("cont_stats", "Continuous variables:",
              choices = c(
                "Mean (SD)"    = "mean_sd",
                "Median [IQR]" = "median_iqr",
                "Both"         = "both"
              ),
              selected = "both"
            ),
            radioButtons("cat_stats", "Categorical variables:",
              choices = c(
                "N (%)"  = "n_pct",
                "% only" = "pct",
                "N only" = "n"
              )
            ),
            hr(),
            checkboxInput("show_missing", "Show missing data count", value = TRUE),
            hr(),
            uiOutput("suggestions_ui")
          )
        ),
        column(8,
          h4("Variables to Include"),
          p(class = "text-muted",
            "Check each variable to include. Edit the label and override the detected type if needed."
          ),
          fluidRow(
            column(1),
            column(3, p(class = "config-header", "Column")),
            column(5, p(class = "config-header", "Display Label")),
            column(3, p(class = "config-header", "Type"))
          ),
          uiOutput("variable_config_ui"),
          br(),
          actionButton("go_preview", "Generate Table \u2192",
            class = "btn-primary btn-lg"
          )
        )
      )
    ),

    # ── Tab 3: Preview & Export ───────────────────────────────────────────────
    tabPanel("3. Preview & Export",
      br(),
      fluidRow(
        column(4,
          wellPanel(
            h4("Export Options"),
            textInput("table_caption", "Caption:",
              value = "Table 1. Participant Characteristics"
            ),
            textAreaInput("table_note", "Note (optional):",
              placeholder = "Note: Data are presented as...",
              rows = 3
            ),
            hr(),
            h5("Pre-Download Checklist"),
            tags$ul(style = "font-size:0.85em; padding-left:18px;",
              tags$li(
                actionLink("title_tips_link", "Title is specific and in Title Case"),
                " \u2014 mention what is described and, if stratified, the grouping variable."
              ),
              tags$li(
                actionLink("note_tips_link", "Note explains how statistics are presented"),
                " and defines all abbreviations (SD, IQR, etc.)."
              ),
              tags$li(tags$b("Labels are in Title Case and proofread."),
                ' Use \u201cAge (Years)\u201d not \u201cage\u201d.'
              ),
              tags$li(tags$b("Units are in labels where needed"),
                ' (e.g., \u201cWeight (kg)\u201d, \u201cSystolic BP (mmHg)\u201d).'
              ),
              tags$li(tags$b("Outcome variables are excluded."),
                " Table 1 describes participants, not results."
              )
            ),
            hr(),
            actionButton("download_word_btn", "Download as Word (.docx)",
              class = "btn-success", style = "width:100%", icon = icon("download")
            ),
            br(), br(),
            p(class = "text-muted", style = "font-size:0.85em;",
              "The Word file is formatted in APA style (Times New Roman, horizontal rules only)."
            )
          )
        ),
        column(8,
          h4("Table Preview"),
          p(class = "text-muted",
            "HTML preview below. The downloaded .docx will be APA-formatted."
          ),
          uiOutput("table1_output")
        )
      )
    )
  )
)
