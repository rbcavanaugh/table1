# Entry point for the Table 1 Generator Shiny app.
# R/helpers.R is auto-sourced by Shiny before this file runs.
# UI is defined in ui.R; server logic is in server.R.

source("ui.R")
source("server.R")

shinyApp(ui = ui, server = server)
