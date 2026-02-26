library(shiny)
library(DT)

ui <- fluidPage(

  tags$head(
    tags$link(
      href = "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap",
      rel  = "stylesheet"
    ),
    tags$style(paste(
      readLines(system.file("table1_defaults_1.0/table1_defaults.css", package = "table1")),
      collapse = "\n"
    )),
    tags$style(HTML("
      .Rtable1 table, table.Rtable1 {
        font-family: 'Inter', sans-serif !important;
        font-size: 0.9em !important;
      }
      .Rtable1 th, .Rtable1 td {
        border-top: none !important;
        border-left: none !important;
        border-right: none !important;
      }
      .Rtable1 thead > tr:first-child > th {
        border-top: 2pt solid black !important;
      }
      .Rtable1 thead > tr:last-child > th {
        border-bottom: 1pt solid black !important;
      }
      .Rtable1 tbody > tr:last-child > td {
        border-bottom: 2pt solid black !important;
      }
      .Rtable1 .firstrow.rowlabel {
        font-weight: bold !important;
      }
      .Rtable1 th.grouplabel > div {
        border-bottom: 1pt solid black !important;
      }
    "))
  ),

  tags$script(src = "copy.js"),

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
    body, .form-control, .selectize-input, .selectize-dropdown, button, label, p, li {
      font-family: 'Inter', sans-serif;
    }
    body {
      background-color: #f8f9fa;
      line-height: 1.6;
    }
    .app-header {
      background: #2c3e50;
      margin: -15px -15px 24px -15px;
      padding: 28px 30px 20px 30px;
    }
    .app-header h2 {
      color: #ffffff;
      font-size: 1.45em;
      font-weight: 600;
      margin: 0 0 2px 0;
    }
    .app-header p {
      color: rgba(255,255,255,0.65);
      font-size: 0.88em;
      margin: 0;
    }
    .well {
      border: none;
      border-radius: 8px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
      background: #ffffff;
    }
    .nav-tabs {
      border-bottom: 2px solid #dee2e6;
    }
    .nav-tabs > li > a {
      border: none;
      border-bottom: 3px solid transparent;
      margin-bottom: -2px;
      color: #666;
      font-weight: 500;
    }
    .nav-tabs > li > a:hover {
      border: none;
      border-bottom: 3px solid #ced4da;
      background: transparent;
      color: #333;
    }
    .nav-tabs > li.active > a,
    .nav-tabs > li.active > a:hover,
    .nav-tabs > li.active > a:focus {
      border: none;
      border-bottom: 3px solid #2c3e50;
      background: transparent;
      color: #2c3e50;
      font-weight: 600;
    }
    .var-row {
      border: 1px solid #e9ecef;
      border-radius: 6px;
      padding: 8px 10px;
      margin-bottom: 6px;
      background: #f8f9fa;
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
    .config-header { font-weight: 600; color: #555; font-size: 0.8em; text-transform: uppercase; margin-bottom: 4px; }
  ")),

  div(class = "app-header",
    tags$h2("Table 1 Builder"),
    tags$p("Publication-ready demographic tables")
  ),

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
            p(tags$b("Missing data: "),
              "Leave missing cells ", tags$b("empty"), " or enter ",
              tags$code("NA"), ". Do not use placeholder values like ",
              tags$code("999"), ", ", tags$code("-99"), ", ",
              tags$code("N/A"), ", or ", tags$code("."),
              " — these will be treated as real data values."
            ),
            hr(),
            downloadButton("download_example", "Download Example Dataset",
              style = "width:100%"
            ),
            hr(),
            p(style = "font-size:0.82em; color:#666; margin-bottom:0;",
              icon("lock"), " Your data never leaves your computer. ",
              actionLink("go_to_about", "Privacy info \u2192")
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
            uiOutput("suggestions_ui"),
            hr(),
            actionButton("go_preview", "Generate Table \u2192",
              class = "btn-primary", style = "width:100%"
            )
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
          uiOutput("variable_config_ui")
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
            ),
            tags$button(
              id = "copy_table_btn",
              "Copy and Paste into Excel/Sheets",
              class = "btn btn-default",
              style = "width:100%;",
              onclick = "copyTable1()"
            ),
            tags$p(id = "copy_table_status",
              style = "font-size:0.85em; margin-top:6px; min-height:1.2em;"
            )
          )
        ),
        column(8,
          h4("Table Preview"),
          div(style = "background:#fff; border-radius:8px; box-shadow:0 1px 4px rgba(0,0,0,0.08); padding:24px; overflow-x:auto;",
            uiOutput("table1_output")
          )
        )
      )
    ),

    # ── Tab 4: About ─────────────────────────────────────────────────────────
    tabPanel("About",
      br(),
      fluidRow(
        column(8, offset = 2,

          h3("About This Site"),
          p("Table 1 Generator is a browser-based tool for creating publication-ready
            demographic tables from research data."),

          hr(),
          h4(icon("lock"), " Privacy"),
          p("This app runs entirely in your browser using WebAssembly \u2014 a technology that
            allows R code to execute locally on your device. When you upload a file, it is
            loaded directly into your browser\u2019s memory and is never transmitted to any
            server. No data is stored, logged, or accessible to anyone else. Closing the tab
            clears everything."),
          p("No tracking scripts, analytics, cookies, or backend server are added by this app.
            It is hosted as a static page on GitHub Pages, which collects basic traffic data
            (page visits) as part of its hosting service. All code for this application is
            openly available at",
            tags$a(href = "https://github.com/rbcavanaugh/table1", target = "_blank", icon("github"),
                   " github.com/rbcavanaugh/table1"), "You are welcome to inspect it, adapt it, or use it as a starting point for your
            own projects."),
          hr(),
          h4(icon("book"), " Acknowledgements & Citations"),
                    p("The core table generation in this app is powered by the ", tags$code("table1"),
            " R package by Benjamin Rich:"),
          div(class = "hint-box", style = "font-family: monospace; font-size: 0.85em;",
            "Rich, B. (2023).", tags$em("table1: Tables of Descriptive Statistics in HTML."),
            "R package.", tags$a("https://CRAN.R-project.org/package=table1",
              href = "https://CRAN.R-project.org/package=table1", target = "_blank")
          ),

          h5("R Packages", style = "margin-top:16px;"),
          p("This app also depends on the following R packages:"),
          tags$ul(
            tags$li(tags$b("shiny"), " \u2014 Chang W, Cheng J, Allaire J, et al. ",
              tags$em("shiny: Web Application Framework for R."),
              tags$a(" https://CRAN.R-project.org/package=shiny",
                href = "https://CRAN.R-project.org/package=shiny", target = "_blank")),
            tags$li(tags$b("flextable"), " \u2014 Gohel D, Skintzos P. ",
              tags$em("flextable: Functions for Tabular Reporting."),
              tags$a(" https://CRAN.R-project.org/package=flextable",
                href = "https://CRAN.R-project.org/package=flextable", target = "_blank")),
            tags$li(tags$b("officer"), " \u2014 Gohel D. ",
              tags$em("officer: Manipulation of Microsoft Word and PowerPoint Documents."),
              tags$a(" https://CRAN.R-project.org/package=officer",
                href = "https://CRAN.R-project.org/package=officer", target = "_blank")),
            tags$li(tags$b("DT"), " \u2014 Xie Y, Cheng J, Tan X. ",
              tags$em("DT: A Wrapper of the JavaScript Library \u2018DataTables\u2019."),
              tags$a(" https://CRAN.R-project.org/package=DT",
                href = "https://CRAN.R-project.org/package=DT", target = "_blank")),
            tags$li(tags$b("readxl"), " \u2014 Wickham H, Bryan J. ",
              tags$em("readxl: Read Excel Files."),
              tags$a(" https://CRAN.R-project.org/package=readxl",
                href = "https://CRAN.R-project.org/package=readxl", target = "_blank")),
            tags$li(tags$b("base64enc"), " \u2014 Urbanek S. ",
              tags$em("base64enc: Tools for base64 Encoding."),
              tags$a(" https://CRAN.R-project.org/package=base64enc",
                href = "https://CRAN.R-project.org/package=base64enc", target = "_blank"))
          ),

          h5("WebR & Shinylive", style = "margin-top:16px;"),
          p("This app is deployed using Shinylive and WebR, which together compile R to
            WebAssembly so the application runs entirely in the browser without a server."),
          tags$ul(
            tags$li(tags$b("shinylive"), " \u2014 Schloerke B, Cheng J. ",
              tags$em("shinylive: Run \u2018shiny\u2019 Applications in the Browser."),
              tags$a(" https://CRAN.R-project.org/package=shinylive",
                href = "https://CRAN.R-project.org/package=shinylive", target = "_blank")),
            tags$li(tags$b("webR"), " \u2014 Stagg G. ",
              tags$em("webR: R in the Browser."),
              tags$a(" https://docs.r-wasm.org/webr/latest/",
                href = "https://docs.r-wasm.org/webr/latest/", target = "_blank"))
          ),

          hr(),
          p(class = "text-muted", style = "font-size:0.85em;",
            "This app was built using ",
            tags$a("Claude Code", href = "https://claude.ai/code", target = "_blank"),
            " by Anthropic."
          ),

          br()
        )
      )
    )
  )
)
