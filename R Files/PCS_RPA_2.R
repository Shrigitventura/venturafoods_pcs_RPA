library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)
library(rio)

###########################################################################################################################################################
############################################################# Data Input ##################################################################################
###########################################################################################################################################################

################ Read original files ####################

mfg_location_tab_raw <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/Stan's folder/All PCS Projects - With MFG Locations (59).csv")



################ Read Data Fixed files ####################
mpi <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MPI.xlsx")
coordinator <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Project Coordinator.xlsx")




##################### (Data - MFG Locs tab) #####################

mfg_location_tab_raw %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(project_number, project_status, project_name, requestor, project_submitted_date, project_type, business_channel, pcs_product_category,
                pcs_product_sub_category, pcs_customer_name, project_minimum_met, task_name, total_actual_duration_days, stage_name, current_responsible_user_name,
                pcs_manufacturing_location, pcs_manufacturing_location_annual_volume_lbs, total_weighted_vm_lb, pcs_manufacturing_location_annual_vm,
                platform_and_packaging_type, pcs_incremental_volume, market_test_start_ship_date, market_test_end_ship_date, market_test_volume,
                new_or_replacing_existing_volume, is_a_market_test, first_production_date, lto_estimated_end_date, desired_launch_date,
                actual_completion_date,  mfg_packet_release_date, sku_label_being_replaced, limited_time_offer, estimated_probability_of_success,
                formula_request_type, sku_base_being_replaced, export_region, suggested_line_number, processing_attribute_type, suggested_deck,
                r_d_formula_number, requested_pack_size, new_sku_base, bulk_oil_type, new_sku_label, manufacturing_packet_needed) -> mfg_location_tab



mfg_location_tab %>% 
  dplyr::mutate(dplyr::across(everything(), ~ifelse(is.na(.), "-", .))) -> mfg_location_tab


# Column: MPI
mpi %>% 
  data.frame() %>% 
  janitor::clean_names() -> mpi_data

mpi_data[!duplicated(mpi_data[,c("project_number")]),] -> mpi_data

mfg_location_tab %>% 
  dplyr::left_join(mpi_data) -> mfg_location_tab



##### date work ####
mfg_location_tab %>% 
  dplyr::select(-contains("date")) -> mfg_location_tab
##### ######### ####

mfg_location_tab_raw %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(contains("date")) -> mfg_location_tab_raw_dates


cbind(mfg_location_tab, mfg_location_tab_raw_dates) -> mfg_location_tab






# Column: Desired Month
mfg_location_tab %>% 
  dplyr::mutate(desired_month = lubridate::month(desired_launch_date),
                desired_month = lubridate::month(desired_month, label = TRUE)) -> mfg_location_tab

# Column: Desired Year
mfg_location_tab %>% 
  dplyr::mutate(desired_year = lubridate::year(desired_launch_date)) -> mfg_location_tab

# Column: Month Year
mfg_location_tab %>% 
  dplyr::mutate(month_year = paste0(desired_month, "-", desired_year)) -> mfg_location_tab


# Column: Project Coordinator
coordinator %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = project_id) -> coordinator_data

coordinator_data[!duplicated(coordinator_data[,c("project_number")]),] -> coordinator_data

mfg_location_tab %>% 
  dplyr::left_join(coordinator_data) -> mfg_location_tab




