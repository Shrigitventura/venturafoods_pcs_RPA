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

options(scipen=999)

###########################################################################################################################################################
############################################################# Data Input ##################################################################################
###########################################################################################################################################################

################ Read original files ####################
pre_final_product <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PCS Weekly files from vscode/2023/10.02.2023/mfg_location_tab_10.02.2023.xlsx")

mfg_location_tab_raw <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/PCS Weekly Raw Data/2023/10.10.23/All PCS Projects - With MFG Locations (72).csv")
pcs_rnd_primary_pack_graphics <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/PCS Weekly Raw Data/2023/10.10.23/PCS R&D Primary & Pack Graphics (53).csv")
velocity_opp_overview <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/Velocity Reports/Opportunity Overview Reports/2023/7 - July/7.17.2023/Velocity Opp Overview.xlsx")
rnd_unique_ingredient_info <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/PCS Weekly Raw Data/2023/10.10.23/PCS R&D Unique Ingredients (54).csv")
master_data <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/PCS Weekly Raw Data/2023/10.10.23/SRCH266209.csv")

################ Read Data Fixed files ####################
coordinator <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/PCS Weekly Raw Data/2023/10.10.23/ProjectList_10_10_2023 11_22_04 AM.xlsx")
process_type <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Process Type.xlsx")
macro <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Macro Platform.xlsx")

##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################

comp <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM Status Update/PRM Status 10.02.23.xlsx") # input Previous week's data
prm <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM/PRM.xlsx") # Make sure to update .xlsx file before you run this tool


##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################

pre_final_product %>% 
  dplyr::mutate(project_submitted_date = lubridate::mdy(project_submitted_date)) %>% 
  dplyr::filter(project_submitted_date >= as.Date("2023-10-02") & project_submitted_date <= as.Date("2023-10-08")) %>%  #### Input the past date to figure out new projects ####
  dplyr::select(project_number) -> new_projects


mpi <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MPI/MPI 10.02.2023.xlsx") 
clean_customer <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Customer Clean/Customer Clean 10.02.2023.xlsx") 

##################### (Velocity Tab) #####################
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
  dplyr::mutate(status = "null") -> prm_stack

comp %>% 
  janitor::clean_names() -> comp_update

rbind(comp_update, prm_stack) -> comp_update_2

comp_update_2



##################### (MPI New Projects) #####################
mpi %>% 
  janitor::clean_names() -> mpi_adding_new

new_projects %>% 
  dplyr::mutate(mpi = "Null",
                project_number = as.double(project_number)) -> new_projects_mpi

rbind(mpi_adding_new, new_projects_mpi) -> new_projects_mpi_added


##################### (Clean Customer) #####################
pre_final_product %>% 
  dplyr::select(project_number, pcs_customer_name, clean_customer) -> pre_final_clean_customer

pre_final_clean_customer[!duplicated(pre_final_clean_customer[,c("project_number")]),] -> pre_final_clean_customer


new_projects %>% 
  dplyr::left_join(pre_final_clean_customer) %>% 
  dplyr::select(pcs_customer_name, clean_customer) -> new_projects_customer_clean_1

clean_customer %>% 
  janitor::clean_names() %>% 
  dplyr::select(pcs_customer_name) -> customer_clean_ref

customer_clean_ref[!duplicated(customer_clean_ref[,c("pcs_customer_name")]),] -> customer_clean_ref

new_projects_customer_clean_1 %>% 
  dplyr::left_join(customer_clean_ref) %>% 
  dplyr::filter(is.na(clean_customer)) %>% 
  dplyr::mutate(clean_customer = "Null",
                cvm = "Null") -> new_projects_customer_clean_2


clean_customer %>% 
  janitor::clean_names() -> clean_customer_2


rbind(clean_customer_2, new_projects_customer_clean_2) -> clean_customer_new_added

##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################

writexl::write_xlsx(comp_update_2, "S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM Status Update/PRM Status 10.10.23.xlsx")
writexl::write_xlsx(new_projects_mpi_added, "S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MPI/MPI 10.10.2023.xlsx")
writexl::write_xlsx(clean_customer_new_added, "S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Customer Clean/Customer Clean 10.10.2023.xlsx")

##############################################################################################################################################################
####################################### Now you go back to the .xlsx file & finish your manual work from PCS System ##########################################
##############################################################################################################################################################

comp <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM Status Update/PRM Status 10.10.23.xlsx")
mpi <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MPI/MPI 10.10.2023.xlsx")
clean_customer <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Customer Clean/Customer Clean 10.10.2023.xlsx")

######################################################################################################################################################

# Clean customer arrange
clean_customer %>%
  dplyr::arrange(pcs_customer_name) -> clean_customer

# Comp 
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

velocity_1 %>% 
  dplyr::filter(is.na(pcs_task)) -> velocity_4


rbind(velocity_2, velocity_4) -> velocity



##################### (Process Type 2 Tab) #####################
process_type %>% 
  janitor::clean_names() %>% 
  dplyr::rename(processing_attribute_type = process_identifier,
                processing_type_desc = process_attribute) %>% 
  dplyr::select(1, 2) -> process_type_desc


mfg_location_tab_raw %>% 
  janitor::clean_names() %>% 
  dplyr::select(project_number, processing_attribute_type) %>% 
  dplyr::left_join(process_type_desc) %>% 
  dplyr::mutate(processing_attribute_type = replace(processing_attribute_type, is.na(processing_attribute_type) | processing_attribute_type == 0, "-")) %>% 
  dplyr::mutate(processing_type_desc = replace(processing_type_desc, is.na(processing_type_desc) | processing_type_desc == 0, "-")) -> process_type_2_tab

process_type_2_tab[!duplicated(process_type_2_tab[,c("project_number")]),] -> process_type_2_tab




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
coordinator[-1:-5, -1] -> coordinator
colnames(coordinator) <- coordinator[1, ]
coordinator %>% 
  dplyr::slice(-1) %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(project_id, project_coordinator) -> coordinator


coordinator %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = project_id) %>% 
  dplyr::mutate(project_number = as.numeric(project_number)) -> coordinator_data

coordinator_data[!duplicated(coordinator_data[,c("project_number")]),] -> coordinator_data



mfg_location_tab %>% 
  dplyr::left_join(coordinator_data) -> mfg_location_tab


# Column: Processing Type
process_type_2_tab %>% 
  dplyr::select(1, 3) -> process_type_desc_data

mfg_location_tab %>% 
  dplyr::left_join(process_type_desc_data) %>% 
  dplyr::rename(processing_type = processing_type_desc) -> mfg_location_tab


# Column: # days since project was created
mfg_location_tab %>% 
  dplyr::mutate(days_since_project_was_created = Sys.Date() - project_submitted_date,
                days_since_project_was_created = as.integer(days_since_project_was_created)) -> mfg_location_tab


# Column: Net Incremental
mfg_location_tab %>% 
  dplyr::mutate(net_incremental = ifelse(pcs_incremental_volume > 0, pcs_incremental_volume, pcs_manufacturing_location_annual_volume_lbs)) -> mfg_location_tab


# Column: R&D Primary
pcs_rnd_primary_pack_graphics %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(project_number, responsible_user_name) -> responsible_user_name_data

responsible_user_name_data[!duplicated(responsible_user_name_data[,c("project_number")]),] -> responsible_user_name_data

mfg_location_tab %>% 
  dplyr::left_join(responsible_user_name_data) %>% 
  dplyr::mutate(responsible_user_name = ifelse(is.na(responsible_user_name), "-", responsible_user_name)) %>% 
  dplyr::rename(rnd_primary = responsible_user_name) -> mfg_location_tab


# Column: Velocity Opp
velocity %>% 
  dplyr::select(1, 2) %>% 
  dplyr::rename(project_number = pcs_fs_id_number,
                velocity_opp = forecast_submission_id) -> velocity_data

mfg_location_tab %>% 
  dplyr::left_join(velocity_data) %>% 
  dplyr::mutate(velocity_opp = ifelse(is.na(velocity_opp), "-", velocity_opp)) -> mfg_location_tab


# Column: Macro Platform
macro %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 4) %>% 
  dplyr::rename(platform_and_packaging_type = platform_1) -> macro_data


mfg_location_tab %>% 
  dplyr::left_join(macro_data) %>% 
  dplyr::mutate(macro_platform = ifelse(is.na(macro_platform), "-", macro_platform)) -> mfg_location_tab


# Column: Past Due
mfg_location_tab %>% 
  dplyr::mutate(past_due = desired_launch_date - Sys.Date(),
                past_due = as.integer(past_due)) -> mfg_location_tab



# Column: PRM Date
prm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(1, 2) %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.numeric(project_number))-> prm_date_data

prm_date_data[!duplicated(prm_date_data[,c("project_number")]),] -> prm_date_data

mfg_location_tab %>% 
  dplyr::left_join(prm_date_data) -> mfg_location_tab

# Column: PRM Topic
prm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(1, 3) %>% 
  dplyr::rename(project_number = pcs_number,
                prm_topic = topic) %>% 
  dplyr::mutate(project_number = as.numeric(project_number)) -> prm_topic_data

prm_topic_data[!duplicated(prm_topic_data[,c("project_number")]),] -> prm_topic_data

mfg_location_tab %>% 
  dplyr::left_join(prm_topic_data) %>% 
  dplyr::mutate(prm_topic = ifelse(is.na(prm_topic), "-", prm_topic)) -> mfg_location_tab

# Column: PRM Expedite Status
prm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(1, 6) %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.numeric(project_number)) -> prm_expedited_status_data

prm_expedited_status_data[!duplicated(prm_expedited_status_data[,c("project_number")]),] -> prm_expedited_status_data

mfg_location_tab %>% 
  dplyr::left_join(prm_expedited_status_data) %>% 
  dplyr::mutate(pcs_expedited_status = ifelse(is.na(pcs_expedited_status), "-", pcs_expedited_status)) -> mfg_location_tab

# Column: Reason for Expedite
prm %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::select(1, 9) %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.numeric(project_number)) -> prm_reason_for_expedite_data

prm_reason_for_expedite_data[!duplicated(prm_reason_for_expedite_data[,c("project_number")]),] -> prm_reason_for_expedite_data

mfg_location_tab %>% 
  dplyr::left_join(prm_reason_for_expedite_data) %>% 
  dplyr::mutate(reason_for_expedite = ifelse(is.na(reason_for_expedite), "-", reason_for_expedite)) -> mfg_location_tab


# Column: PRM Submit Month
mfg_location_tab %>% 
  dplyr::mutate(prm_submit_month = lubridate::month(prm_date, label = TRUE)) -> mfg_location_tab

# Column: PRM Submit Year
mfg_location_tab %>% 
  dplyr::mutate(prm_submit_year = lubridate::year(prm_date)) -> mfg_location_tab


# Column: Clean Customer & CVM
clean_customer %>% 
  janitor::clean_names() %>% 
  data.frame() -> clean_customer_data

clean_customer_data[!duplicated(clean_customer_data[,c("pcs_customer_name")]),] -> clean_customer_data

mfg_location_tab %>% 
  dplyr::left_join(clean_customer_data) -> mfg_location_tab

mfg_location_tab %>% 
  dplyr::mutate(cvm = ifelse(is.na(cvm), "N", cvm)) -> mfg_location_tab


# Column: MDM Total Oil %
# Define the column names
master_data %>% 
  dplyr::slice(-(1:5)) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(a = product_search_excel_option_finished_good) %>% 
  tidyr::separate(a, c("new_sku_base", "label", "sku_description", "bulk_oil", "refrigeration_code", 
                       "percent_of_oil_in_product", "cases_per_pallet", "product_platform"), sep = ",") %>%
  dplyr::mutate(percent_of_oil_in_product = as.double(percent_of_oil_in_product),
                new_sku_base = as.double(new_sku_base)) %>% 
  dplyr::mutate(mdm_total_oil_percent = percent_of_oil_in_product * 0.01) %>% 
  dplyr::select(new_sku_base, mdm_total_oil_percent) -> master_data_oil_percent


mfg_location_tab %>% 
  dplyr::left_join(master_data_oil_percent) %>% 
  dplyr::mutate(mdm_total_oil_percent = ifelse(is.na(mdm_total_oil_percent), 0, mdm_total_oil_percent)) -> mfg_location_tab



# Column: Total Incremental Oil Volume 
mfg_location_tab %>% 
  dplyr::mutate(net_incremental = as.double(net_incremental)) %>% 
  dplyr::mutate(total_incremental_oil_volume = mdm_total_oil_percent * net_incremental,
                total_incremental_oil_volume = round(total_incremental_oil_volume, 0)) -> mfg_location_tab


# Date format work
mfg_location_tab %>% 
  dplyr::mutate(dplyr::across(contains("date"), ~ as.Date(.x, format="%Y-%m-%d"))) %>%
  dplyr::mutate(dplyr::across(contains("date"), ~ format(.x, "%m/%d/%Y"))) -> mfg_location_tab

# Data Clean Work
mfg_location_tab %>% 
  dplyr::mutate(total_actual_duration_days = as.numeric(total_actual_duration_days),
                total_actual_duration_days = round(total_actual_duration_days, 0)) -> mfg_location_tab

# Duplication removal
mfg_location_tab
mfg_location_tab[!duplicated(mfg_location_tab[,c("project_number", "pcs_manufacturing_location")]),] -> mfg_location_tab




################################################################################################################################################################
# Date work for Velocity Tab
velocity %>% 
  dplyr::mutate(dplyr::across(contains("date"), ~ as.Date(.x, format="%Y-%m-%d"))) %>%
  dplyr::mutate(dplyr::across(contains("date"), ~ format(.x, "%m/%d/%Y"))) -> velocity







##################### (R&D Unique Ingredients) #####################
rnd_unique_ingredient_info %>% 
  janitor::clean_names() %>% 
  data.frame() -> rnd_unique_ingredient_info






##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################
##############################################################################################################################################################



### Exporting ###

list("Data" = mfg_location_tab,
     "Velocity" = velocity,
     "R&D Unique Ingredient Info" = rnd_unique_ingredient_info) -> list_of_dfs


writexl::write_xlsx(mfg_location_tab, "S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PCS Weekly files from vscode/2023/10.10.2023/mfg_location_tab_10.10.2023.xlsx")
writexl::write_xlsx(list_of_dfs, "S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PCS Weekly files from vscode/2023/10.10.2023/pcs_data_10.10.2023.xlsx")



