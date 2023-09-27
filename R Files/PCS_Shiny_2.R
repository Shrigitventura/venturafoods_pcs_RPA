# Increase max upload size to 50 MB
options(shiny.maxRequestSize = 50 * 1024 * 1024)

library(shiny)
library(shinythemes)
library(dplyr)
library(janitor)
library(readxl)

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
                        fileInput("comp", "Comp (.xlsx)")
                      ),
                      wellPanel(
                        tags$h3("Import Data Fixed Files"),
                        fileInput("mpi", "MPI (.xlsx)"),
                        fileInput("coordinator", "Coordinator (.xlsx)"),
                        fileInput("process_type", "Process Type (.xlsx)"),
                        fileInput("macro", "Macro Platform (.xlsx)"),
                        fileInput("clean_customer", "Clean Customer (.xlsx)"),
                        fileInput("master_data", "Master Data (.csv)")
                      ),
                      actionButton("run_data_change", "Run Data Change")
             )
  )
)

server <- function(input, output, session) {
  observeEvent(input$run_data_change, {
    inFile_mfg <- input$mfg_location
    inFile_comp <- input$comp
    # Additional file variables can be defined similarly
    
    if (is.null(inFile_mfg) || is.null(inFile_comp)) {
      return(NULL)
    }
    
    mfg_location_tab_raw <- read.csv(inFile_mfg$datapath)
    comp <- read_excel(inFile_comp$datapath)
    # Additional file read operations can go here
    
    comp %>% 
      janitor::clean_names() %>% 
      dplyr::rename(pcs_fs_id_number = project_number,
                    pcs_stage_1 = status) -> comp_velocity
    
    comp_velocity[!duplicated(comp_velocity[,c("pcs_fs_id_number")]),] -> comp_velocity
    
    mfg_location_tab_raw %>% 
      janitor::clean_names() %>% 
      dplyr::select(1, 2, 15) %>% 
      dplyr::rename(pcs_fs_id_number = project_number,
                    pcs_stage_2 = project_status,
                    pcs_task = task_name) -> proj_velocity
    
    proj_velocity[!duplicated(proj_velocity[,c("pcs_fs_id_number")]),] -> proj_velocity
    
    velocity_opp_overview %>% 
      janitor::clean_names() %>% 
      dplyr::filter(project_status == "Submitted to PCS") %>% 
      dplyr::select(pcs_fs_id_number, forecast_submission_id, operations_confirmed_capacity, capacity_confirm_date) %>% 
      dplyr::mutate(pcs_fs_id_number = as.double(pcs_fs_id_number),
                    capacity_confirm_date = as.Date(capacity_confirm_date)) %>% 
      dplyr::left_join(comp_velocity) %>% 
      dplyr::left_join(proj_velocity) %>% 
      dplyr::mutate(pcs_stage = ifelse(is.na(pcs_stage_1), pcs_stage_2, pcs_stage_1)) %>% 
      dplyr::select(-pcs_stage_1, -pcs_stage_2) %>% 
      dplyr::relocate(pcs_fs_id_number, forecast_submission_id, pcs_stage, pcs_task, operations_confirmed_capacity, capacity_confirm_date) -> velocity_1
    
    
    mfg_location_tab_raw %>% 
      dplyr::select(1, 2) %>% 
      janitor::clean_names() %>% 
      dplyr::rename(pcs_stage = project_status,
                    pcs_fs_id_number = project_number) -> proj_pcs_status
    
    proj_pcs_status[!duplicated(proj_pcs_status[,c("pcs_fs_id_number")]),] -> proj_pcs_status
    
    velocity_1 %>% 
      dplyr::filter(!is.na(pcs_task)) %>% 
      dplyr::filter(!is.na(pcs_stage)) %>% 
      dplyr::left_join(proj_pcs_status) -> velocity_2
    
    
    comp_velocity %>% 
      dplyr::rename(pcs_stage = pcs_stage_1) -> comp_velocity_3
    
    
    velocity_1 %>% 
      dplyr::filter(is.na(pcs_task)) %>% 
      dplyr::filter(is.na(pcs_stage)) %>% 
      dplyr::left_join(comp_velocity_3) -> velocity_3
    
    velocity_3 %>% 
      dplyr::select(1) %>% 
      dplyr::rename(project_number = pcs_fs_id_number) %>% 
      dplyr::mutate(status = "null") %>% 
      dplyr::mutate(date_as_of = Sys.Date()) -> prm_stack
    
    comp %>% 
      janitor::clean_names() %>% 
      dplyr::mutate(date_as_of = as.Date(date_as_of)) -> comp_update
    
    rbind(comp_update, prm_stack) -> comp_update_2
    
    output$view_comp <- renderTable({
      comp
    })

  })
}

shinyApp(ui, server)
