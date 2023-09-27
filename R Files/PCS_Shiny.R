library(shiny)
library(shinythemes)

ui <- fluidPage(
  theme = shinythemes::shinytheme("cosmo"),
  navbarPage("Ventura Foods PCS Data Automation App",
             tabPanel("Step 1 (Import files)",
                      wellPanel(
                        tags$h3("Import Original Files"),
                        fileInput("mfg_location", "MFG Locations (.csv)"),
                        fileInput("pcs_rnd", "PCS R&D Primary (.csv)"),
                        fileInput("velocity_opp", "Velocity Opp (.xlsx)"),
                        fileInput("prm", "PRM (.xlsx)"),
                        fileInput("rnd_unique", "PCS R&D Unique (.csv)"),
                        fileInput("comp", "Comp (.xlsx)")  # Added this line
                      ),
                      wellPanel(
                        tags$h3("Import Data Fixed Files"),
                        fileInput("mpi", "MPI (.xlsx)"),
                        fileInput("coordinator", "Coordinator (.xlsx)"),
                        fileInput("process_type", "Process Type (.xlsx)"),
                        fileInput("macro", "Macro Platform (.xlsx)"),
                        fileInput("clean_customer", "Clean Customer (.xlsx)"),
                        fileInput("master_data", "Master Data (.csv)")
                      )
             )
  )
)

server <- function(input, output, session) {
  # Server logic for handling file uploads will go here
}

shinyApp(ui, server)
