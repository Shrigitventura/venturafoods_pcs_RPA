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

# Question #1: "MajorProj MonthYr" is this changing data?


# Sdrive - Global Shared - Large - S&OP - PCS - Reporting - Stans's folder
# plan: create combined excel file, and put them into microstrategy

################ Read original files ####################

projectlist <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/Weekly PCS Project Rpt/2023/7 July/ProjectList_7_3_2023 11_44_28 AM.xlsx")
all_pcs_project_with_mfg_location <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/Stan's folder/All PCS Projects - With MFG Locations (59).csv")
all_pcs_project_without_location <- read_csv("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/Stan's folder/All PCS Projects - Without Locations (43).csv")

##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################


################ Read Data Fixed files ####################
active_pack_comp <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Active Pack Comp.xlsx")
majorproj_monthyr <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MajorProj MonthYr.xlsx")
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/DSX.xlsx")
r_d_primary <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/R & D Primary.xlsx")
minimums <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Minimums.xlsx")
comp <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM Status.xlsx")
prm <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/PRM.xlsx")
velocity <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Velocity.xlsx")
macro <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Macro Platform.xlsx")
stocking <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Stocking.xlsx")
master_data <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Master Data.xlsx")
snop_comment_latest <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/S&OP Comment Latest.xlsx")
packaging_specialist <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Packaging Specialist.xlsx")
test <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Test Flag Code.xlsx")
pmo <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Reg PMO Team.xlsx")
stage_3_capacity_check <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/S&OP CapS3 Comment Latest.xlsx")
stage_1_capacity_check <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/S&OP CapS1 Comment Latest.xlsx")
comp_type <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Component Type.xlsx")

################# ETL (Extract, Transform, Load) ###################

##################### (Project Coordinator tab) #####################
projectlist[-1:-5, -1] -> projectlist
colnames(projectlist) <- projectlist[1, ]
projectlist[-1, ] -> projectlist

projectlist %>% 
  janitor::clean_names() %>% 
  dplyr::select(project_id, project_coordinator) %>% 
  readr::type_convert() -> project_coordinator


##################### (Data - MFG Locs tab) #####################
# main board (col A to BD)
all_pcs_project_with_mfg_location %>% 
  data.frame() %>% 
  janitor::clean_names() -> data_mfg_locs

# dsx_loc (col BN)
data_mfg_locs %>% 
  dplyr::mutate(dsx_loc = paste0(project_number, pcs_manufacturing_location)) -> data_mfg_locs

# Component (Col BE)
active_pack_comp %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 5) %>% 
  dplyr::rename(dsx_loc = projloc) -> active_pack_comp_1

active_pack_comp_1[!duplicated(active_pack_comp_1[,c("dsx_loc")]),] -> active_pack_comp_1

data_mfg_locs %>% 
  dplyr::left_join(active_pack_comp_1) %>% 
  dplyr::mutate(component_type_code = replace(component_type_code, is.na(component_type_code), 0)) -> data_mfg_locs


# Old Component Code (Col BF)
active_pack_comp %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 6) %>% 
  dplyr::rename(dsx_loc = projloc) -> active_pack_comp_2

active_pack_comp_2[!duplicated(active_pack_comp_2[,c("dsx_loc")]),] -> active_pack_comp_2

data_mfg_locs %>% 
  dplyr::left_join(active_pack_comp_2) %>% 
  dplyr::mutate(old_component_code = replace(old_component_code, is.na(old_component_code), 0)) -> data_mfg_locs


# New Component Code (Col BG)
active_pack_comp %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 7) %>% 
  dplyr::rename(dsx_loc = projloc) -> active_pack_comp_3

active_pack_comp_3[!duplicated(active_pack_comp_3[,c("dsx_loc")]),] -> active_pack_comp_3

data_mfg_locs %>% 
  dplyr::left_join(active_pack_comp_3) %>% 
  dplyr::mutate(new_component_code = replace(new_component_code, is.na(new_component_code), 0)) -> data_mfg_locs

# MPI (Col BH) 
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::rename(mpi = major_initative) -> mpi

data_mfg_locs %>% 
  dplyr::left_join(mpi) -> data_mfg_locs

# Desired launch month (Col BI)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,3) %>% 
  dplyr::rename(desired_launch_month = desired_launch_month_name) -> desired_launch_month

data_mfg_locs %>% 
  dplyr::left_join(desired_launch_month) -> data_mfg_locs


# Desired Launch Year (Col BJ)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,5) -> desired_launch_year

data_mfg_locs %>% 
  dplyr::left_join(desired_launch_year) -> data_mfg_locs


# Project Coordinator (Col BK)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,6) -> project_coordinator

data_mfg_locs %>% 
  dplyr::left_join(project_coordinator) %>% 
  dplyr::mutate(project_coordinator = replace(project_coordinator, is.na(project_coordinator), 0)) -> data_mfg_locs


# Processing Type (Col BL)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,8) -> processing_type

data_mfg_locs %>% 
  dplyr::left_join(processing_type) -> data_mfg_locs


# Submitted Month (Col BP)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,12) %>% 
  dplyr::rename(submitted_month = submitted_month_name) -> submitted_month

data_mfg_locs %>% 
  dplyr::left_join(submitted_month) %>% 
  dplyr::mutate(submitted_month = replace(submitted_month, is.na(submitted_month), 0)) -> data_mfg_locs


# Submitted Year (Col BQ)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,14) -> submitted_year

data_mfg_locs %>% 
  dplyr::left_join(submitted_year) %>% 
  dplyr::mutate(submitted_year = replace(submitted_year, is.na(submitted_year), 0)) -> data_mfg_locs

# Channel 1 Approver (Col BR)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,17) -> channel_1_approver

data_mfg_locs %>% 
  dplyr::left_join(channel_1_approver) %>% 
  dplyr::mutate(channel_1_approver = replace(channel_1_approver, is.na(channel_1_approver), 0)) -> data_mfg_locs


# Channel 2 Approver (Col BS)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,18) -> channel_2_approver

data_mfg_locs %>% 
  dplyr::left_join(channel_2_approver) %>% 
  dplyr::mutate(channel_2_approver = replace(channel_2_approver, is.na(channel_2_approver), 0)) -> data_mfg_locs


# Sales Project Notes (Col BT)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,19) -> sales_project_notes

data_mfg_locs %>% 
  dplyr::left_join(sales_project_notes) %>% 
  dplyr::mutate(sales_project_notes = replace(sales_project_notes, is.na(sales_project_notes) | sales_project_notes == 0, "-")) -> data_mfg_locs


# Sales Project Note Timestamp (Col BU)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,20) -> sales_project_note_timestamp

data_mfg_locs %>% 
  dplyr::left_join(sales_project_note_timestamp) %>% 
  dplyr::mutate(sales_project_note_timestamp = as.Date(sales_project_note_timestamp)) -> data_mfg_locs

# Sales Comment Author (Col BV)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,21) -> sales_comment_author

data_mfg_locs %>% 
  dplyr::left_join(sales_comment_author) %>% 
  dplyr::mutate(sales_comment_author = replace(sales_comment_author, is.na(sales_comment_author) | sales_comment_author == 0, "-")) -> data_mfg_locs

# Latest Comment Type (Col BW)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,22) -> latest_comment_type

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_type) %>% 
  dplyr::mutate(latest_comment_type = replace(latest_comment_type, is.na(latest_comment_type) | latest_comment_type == 0, "-")) -> data_mfg_locs


# Latest Comment (Col BX)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,23) -> latest_comment

data_mfg_locs %>% 
  dplyr::left_join(latest_comment) %>% 
  dplyr::mutate(latest_comment = replace(latest_comment, is.na(latest_comment) | latest_comment == 0, "-")) -> data_mfg_locs

# Latest Comment Date (Col BY)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,24) -> latest_comment_date

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_date) %>% 
  dplyr::mutate(latest_comment_date = as.Date(latest_comment_date)) -> data_mfg_locs

# Latest Comment Author (Col BZ)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,25) -> latest_comment_author

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_author) %>% 
  dplyr::mutate(latest_comment_author = replace(latest_comment_author, is.na(latest_comment_author) | latest_comment_author == 0, "-")) -> data_mfg_locs


# days since Project was created (Col CA)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,29) %>% 
  dplyr::rename(days_since_projecet_was_created = x_days_since_project_was_created) -> days_since_projecet_was_created

data_mfg_locs %>% 
  dplyr::left_join(days_since_projecet_was_created) %>% 
  dplyr::mutate(days_since_projecet_was_created = replace(days_since_projecet_was_created, is.na(days_since_projecet_was_created) | days_since_projecet_was_created == 0, "-")) -> data_mfg_locs


# Actual Comp Month (Col CB)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,32) -> actual_comp_month

data_mfg_locs %>% 
  dplyr::left_join(actual_comp_month) %>% 
  dplyr::mutate(actual_comp_month = replace(actual_comp_month, is.na(actual_comp_month) | actual_comp_month == 0, "-")) -> data_mfg_locs


# Actual Comp Year (Col CC)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,33) -> actual_comp_year

data_mfg_locs %>% 
  dplyr::left_join(actual_comp_year) %>% 
  dplyr::mutate(actual_comp_year = replace(actual_comp_year, is.na(actual_comp_year) | actual_comp_year == 0, "-")) -> data_mfg_locs


# Customer Clean (Col CE)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,34) -> customer_clean

data_mfg_locs %>% 
  dplyr::left_join(customer_clean) %>% 
  dplyr::mutate(customer_clean = replace(customer_clean, is.na(customer_clean) | customer_clean == 0, "-")) -> data_mfg_locs


# Requestor Latest (Col CF)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,35) -> requestor_latest

data_mfg_locs %>% 
  dplyr::left_join(requestor_latest) %>% 
  dplyr::mutate(requestor_latest = replace(requestor_latest, is.na(requestor_latest) | requestor_latest == 0, "-")) -> data_mfg_locs


# Export Region2 (Col CG)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,37) %>% 
  dplyr::rename(export_region_2 = export_region) -> export_region_2

data_mfg_locs %>% 
  dplyr::left_join(export_region_2) %>% 
  dplyr::mutate(export_region_2 = replace(export_region_2, is.na(export_region_2) | export_region_2 == 0, "-")) -> data_mfg_locs


# Project Type Shortened (Col CM)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,7) -> project_type_shortened
  
data_mfg_locs %>% 
  dplyr::left_join(project_type_shortened) %>% 
  dplyr::mutate(project_type_shortened = replace(project_type_shortened, is.na(project_type_shortened) | project_type_shortened == 0, "-")) -> data_mfg_locs


# PRM Date (Col CO)
majorproj_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,38) -> prm_date

data_mfg_locs %>% 
  dplyr::left_join(prm_date) %>% 
  dplyr::mutate(prm_date = as.Date(prm_date)) -> data_mfg_locs


# DSX (Col BO)
dsx %>% 
  janitor::clean_names() %>% 
  dplyr::rename(dsx_loc = row_labels,
                dsx = sum_of_current_calendar_2yrs_forecast_by_ship_to_mfg_for_pcs_team) %>% 
  data.frame() -> dsx

data_mfg_locs %>% 
  dplyr::left_join(dsx) %>% 
  dplyr::mutate(dsx = replace(dsx, is.na(dsx) | dsx == 0, "-")) -> data_mfg_locs


# R&D Primary
r_d_primary %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 6) %>% 
  dplyr::rename(project_number = project_number_1,
                r_d_primary = responsible_user_name) -> r_d_primary

data_mfg_locs %>% 
  dplyr::left_join(r_d_primary) %>% 
  dplyr::mutate(r_d_primary = replace(r_d_primary, is.na(r_d_primary) | r_d_primary == 0, "-")) -> data_mfg_locs


# Min Key (Col CI)
data_mfg_locs %>% 
  dplyr::mutate(min_key = paste0(project_type, pcs_product_category, pcs_product_sub_category)) -> data_mfg_locs

# Min Annual Volume (Col CJ)
minimums %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 5) %>% 
  data.frame() %>% 
  dplyr::rename(min_key = key,
                min_annual_volume = mn_volume) -> minimums_1

data_mfg_locs %>% 
  dplyr::left_join(minimums_1) -> data_mfg_locs

# Min Annual VM$ (Col CK)
minimums %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 6) %>% 
  data.frame() %>% 
  dplyr::rename(min_key = key,
                min_annual_vm = mn_vmdol) -> minimums_2

data_mfg_locs %>% 
  dplyr::left_join(minimums_2) %>% 
  dplyr::mutate(min_annual_vm = sprintf("$%.2f", min_annual_vm)) -> data_mfg_locs


# Min Annual VM$/LB (Col CL)
minimums %>% 
  janitor::clean_names() %>% 
  dplyr::select(1, 7) %>% 
  data.frame() %>% 
  dplyr::rename(min_key = key,
                min_annual_vm_LB = mn_vmlbs) -> minimums_3

data_mfg_locs %>% 
  dplyr::left_join(minimums_3) %>% 
  dplyr::mutate(min_annual_vm_LB = sprintf("$%.2f", min_annual_vm_LB)) -> data_mfg_locs


# Net Incremental (Col CD)
data_mfg_locs %>% 
  dplyr::mutate(net_incremental = ifelse(pcs_incremental_volume > 0, pcs_incremental_volume, pcs_manufacturing_location_annual_volume_lbs)) -> data_mfg_locs

                
# PRM Topic (Col CP)
prm %>% 
  dplyr::select(1, 3) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_number,
                prm_topic = topic) %>% 
  dplyr::mutate(project_number = as.double(project_number)) -> prm_topic

prm_topic[!duplicated(prm_topic[,c("project_number")]),] -> prm_topic

data_mfg_locs %>% 
  dplyr::left_join(prm_topic) %>% 
  dplyr::mutate(prm_topic = replace(prm_topic, is.na(prm_topic) | prm_topic == 0, "-")) -> data_mfg_locs  

# PCS Expedited Status (Col CQ)
prm %>% 
  dplyr::select(1, 6) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.double(project_number)) -> pcs_expedited_status

pcs_expedited_status[!duplicated(pcs_expedited_status[,c("project_number")]),] -> pcs_expedited_status

data_mfg_locs %>% 
  dplyr::left_join(pcs_expedited_status) %>% 
  dplyr::mutate(pcs_expedited_status = replace(pcs_expedited_status, is.na(pcs_expedited_status) | pcs_expedited_status == 0, "-")) -> data_mfg_locs  

# PRM Submit Month (Col CR)
prm %>% 
  dplyr::select(1, 7) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_number,
                prm_submit_month = prm_submit_month_year) %>% 
  dplyr::mutate(project_number = as.double(project_number)) -> prm_submit_month

prm_submit_month[!duplicated(prm_submit_month[,c("project_number")]),] -> prm_submit_month

data_mfg_locs %>% 
  dplyr::left_join(prm_submit_month) %>% 
  dplyr::mutate(prm_submit_month = replace(prm_submit_month, is.na(prm_submit_month) | prm_submit_month == 0, "-")) -> data_mfg_locs 

# PRM Submit Year (Col CS)
prm %>% 
  dplyr::select(1, 8) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.double(project_number)) -> prm_submit_year

prm_submit_year[!duplicated(prm_submit_year[,c("project_number")]),] -> prm_submit_year

data_mfg_locs %>% 
  dplyr::left_join(prm_submit_year) %>% 
  dplyr::mutate(prm_submit_year = replace(prm_submit_year, is.na(prm_submit_year) | prm_submit_year == 0, "-")) -> data_mfg_locs 

# Reason for Expedite (Col CT)
prm %>% 
  dplyr::select(1, 9) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_number) %>% 
  dplyr::mutate(project_number = as.double(project_number)) -> reason_for_expedite

reason_for_expedite[!duplicated(reason_for_expedite[,c("project_number")]),] -> reason_for_expedite

data_mfg_locs %>% 
  dplyr::left_join(reason_for_expedite) %>% 
  dplyr::mutate(reason_for_expedite = replace(reason_for_expedite, is.na(reason_for_expedite) | reason_for_expedite == 0, "-")) -> data_mfg_locs


# Velocity Opp (Col CN)
velocity %>% 
  dplyr::select(1, 2) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_fs_id_number, 
                velocity_opp = forecast_submission_id) -> velocity_opp

data_mfg_locs %>% 
  dplyr::left_join(velocity_opp) %>% 
  dplyr::mutate(velocity_opp = replace(velocity_opp, is.na(velocity_opp) | velocity_opp == 0, "-")) -> data_mfg_locs


# Velocity Capacity Check (Col DI)
velocity %>% 
  dplyr::select(1, 5) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_fs_id_number, 
                velocity_capacity_check = operations_confirmed_capacity) -> velocity_capacity_check

data_mfg_locs %>% 
  dplyr::left_join(velocity_capacity_check) %>% 
  dplyr::mutate(velocity_capacity_check = replace(velocity_capacity_check, is.na(velocity_capacity_check) | velocity_capacity_check == 0, "-")) -> data_mfg_locs


# Velocity Capacity Check Date (DJ)
velocity %>% 
  dplyr::select(1, 6) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = pcs_fs_id_number, 
                velocity_capacity_check_date = capacity_confirm_date) %>% 
  dplyr::mutate(velocity_capacity_check_date = as.Date(velocity_capacity_check_date)) -> velocity_capacity_check_date

data_mfg_locs %>% 
  dplyr::left_join(velocity_capacity_check_date) %>% 
  dplyr::mutate(velocity_capacity_check_date = ifelse(is.na(velocity_capacity_check_date), "_", velocity_capacity_check_date)) -> data_mfg_locs

# Macro Platform (Col CU)
macro %>% 
  dplyr::select(1, 4) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(platform_and_packaging_type = platform_1) -> macro_platform

data_mfg_locs %>% 
  dplyr::left_join(macro_platform) %>% 
  dplyr::mutate(macro_platform = ifelse(is.na(macro_platform), "_", macro_platform)) -> data_mfg_locs


# Less than 1 pallet/month Min (Col CV)
stocking %>% 
  dplyr::select(1, 67) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = project_number_1) -> stocking_less_than_1_pallet

stocking_less_than_1_pallet[!duplicated(stocking_less_than_1_pallet[,c("project_number")]),] -> stocking_less_than_1_pallet

data_mfg_locs %>% 
  dplyr::left_join(stocking_less_than_1_pallet) %>% 
  dplyr::mutate(less_than_1_pallet_month_min = replace(less_than_1_pallet_month_min, is.na(less_than_1_pallet_month_min) | less_than_1_pallet_month_min == 0, "-")) -> data_mfg_locs


# MDM Base Bulk Oil (Col  CW)
master_data %>% 
  dplyr::select(1, 4) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(new_sku_base = base_product_code,
                mdm_base_bulk_oil = bulk_oil_type) -> master_data_bulk_oil

master_data_bulk_oil[!duplicated(master_data_bulk_oil[,c("new_sku_base")]),] -> master_data_bulk_oil

data_mfg_locs %>% 
  dplyr::left_join(master_data_bulk_oil) %>% 
  dplyr::mutate(mdm_base_bulk_oil = replace(mdm_base_bulk_oil, is.na(mdm_base_bulk_oil) | mdm_base_bulk_oil == 0, "-")) -> data_mfg_locs

# MDM Total Oil % (Col CX)
master_data %>% 
  dplyr::select(1, 9) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(new_sku_base = base_product_code,
                mdm_total_oil_percent = percent_of_oil) -> master_data_total_percent

master_data_total_percent[!duplicated(master_data_total_percent[,c("new_sku_base")]),] -> master_data_total_percent

data_mfg_locs %>% 
  dplyr::left_join(master_data_total_percent) -> data_mfg_locs


# Total Incremental Oil Volume (Col CY)
data_mfg_locs %>% 
  dplyr::mutate(total_incremental_oil_volume = mdm_total_oil_percent * net_incremental) %>% 
  dplyr::mutate(total_incremental_oil_volume = round(total_incremental_oil_volume, 0)) %>% 
  dplyr::mutate(total_incremental_oil_volume = replace(total_incremental_oil_volume, is.na(total_incremental_oil_volume) | total_incremental_oil_volume == 0, "-")) -> data_mfg_locs


# PCS NF or MDM Bulk Oil (Col CZ)
data_mfg_locs %>% 
  dplyr::mutate(pcs_nf_or_mdm_bulk_oil = ifelse(mdm_base_bulk_oil == 0, new_sku_base, mdm_base_bulk_oil)) %>% 
  dplyr::mutate(mdm_total_oil_percent = sprintf("%1.2f%%", 100 * mdm_total_oil_percent)) %>% 
  dplyr::mutate(mdm_total_oil_percent = ifelse(mdm_total_oil_percent == "NA%", "-", mdm_total_oil_percent)) -> data_mfg_locs

# S&OP Comment (Col DA)
snop_comment_latest %>% 
  dplyr::select(1, 3) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(snop_comment = pcs_comment) -> snop_comment

data_mfg_locs %>% 
  dplyr::left_join(snop_comment) %>% 
  dplyr::mutate(snop_comment = replace(snop_comment, is.na(snop_comment) | snop_comment == 0, "-")) -> data_mfg_locs

# S&OP Comment Date (Col DB)
snop_comment_latest %>% 
  dplyr::select(1, 4) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(snop_comment_date = pcs_comment_last_updated_date) %>% 
  dplyr::mutate(snop_comment_date = as.Date(snop_comment_date)) -> snop_comment_date

data_mfg_locs %>% 
  dplyr::left_join(snop_comment_date) %>% 
  dplyr::mutate(snop_comment_date = ifelse(is.na(snop_comment_date), 0, snop_comment_date)) -> data_mfg_locs


# S&OP Comment Author (Col DC)
snop_comment_latest %>% 
  dplyr::select(1, 5) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(snop_comment_author = pcs_comment_last_updated_by) -> snop_comment_author

data_mfg_locs %>% 
  dplyr::left_join(snop_comment_author) %>% 
  dplyr::mutate(snop_comment_author = ifelse(is.na(snop_comment_author), "-", snop_comment_author)) -> data_mfg_locs


# Packaging Graphics Coordinator (Col DD)
packaging_specialist %>% 
  dplyr::select(1, 6) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(project_number = project_number_1,
                packaging_graphics_coordinator = responsible_user_name) -> packaging_graphics_coordinator


data_mfg_locs %>% 
  dplyr::left_join(packaging_graphics_coordinator) %>% 
  dplyr::mutate(packaging_graphics_coordinator = ifelse(is.na(packaging_graphics_coordinator), "-", packaging_graphics_coordinator)) -> data_mfg_locs



# Multiple MFG Plants per Project (Col DE)
data_mfg_locs %>% 
  dplyr::group_by(project_number) %>% 
  dplyr::summarize(project_number_count_if = n()) %>% 
  dplyr::mutate(count = table(project_number)) %>% 
  dplyr::select(1, 2) -> project_number_count_if


data_mfg_locs %>% 
  dplyr::left_join(project_number_count_if) %>% 
  dplyr::mutate(multiple_mfg_plants_per_project = ifelse(project_number_count_if > 1, "Yes", "No")) %>% 
  dplyr::select(-project_number_count_if) -> data_mfg_locs


# Test Run Required2 (Col DF)
test %>% 
  janitor::clean_names() %>% 
  dplyr::rename(test_run_required2 = test,
                manufacturing_packet_needed = x2) -> test_flag_code

data_mfg_locs %>% 
  dplyr::left_join(test_flag_code) %>% 
  dplyr::mutate(test_run_required2 = replace(test_run_required2, is.na(test_run_required2), 0)) -> data_mfg_locs


# Plant PMO Region (Col DG)
pmo %>% 
  dplyr::select(1, 3) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(pcs_manufacturing_location = plant,
                plant_pmo_region = pmo_region) -> pmo_region


data_mfg_locs %>% 
  dplyr::left_join(pmo_region) %>% 
  dplyr::mutate(plant_pmo_region = replace(plant_pmo_region, is.na(plant_pmo_region) | plant_pmo_region == 0, "-")) -> data_mfg_locs


# Plant PMO Lead (Col DH)
pmo %>% 
  dplyr::select(1, 2) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(pcs_manufacturing_location = plant,
                plant_pmo_lead = pmo_lead) -> pmo_lead

data_mfg_locs %>% 
  dplyr::left_join(pmo_lead) %>% 
  dplyr::mutate(plant_pmo_lead = replace(plant_pmo_lead, is.na(plant_pmo_lead) | plant_pmo_lead == 0, "-")) -> data_mfg_locs


# Stage 3 Capacity Check (Col DK)
stage_3_capacity_check %>% 
  dplyr::select(1, 3) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(stage_3_capacity_check = pcs_comment) -> stage_3_capacity_check_1

data_mfg_locs %>% 
  dplyr::left_join(stage_3_capacity_check_1) %>% 
  dplyr::mutate(stage_3_capacity_check = replace(stage_3_capacity_check, is.na(stage_3_capacity_check) | stage_3_capacity_check == 0, "-")) -> data_mfg_locs


# Stage 3 Capacity Check Date (Col DL)
stage_3_capacity_check %>% 
  dplyr::select(1, 4) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(stage_3_capacity_check_date = pcs_comment_last_updated_date) %>% 
  dplyr::mutate(stage_3_capacity_check_date = as.Date(stage_3_capacity_check_date)) -> stage_3_capacity_check_date

data_mfg_locs %>% 
  dplyr::left_join(stage_3_capacity_check_date) %>% 
  dplyr::mutate(stage_3_capacity_check_date = ifelse(is.na(stage_3_capacity_check_date), 0, stage_3_capacity_check_date)) -> data_mfg_locs



# Stage 1 Capacity Check (Col DM)
stage_1_capacity_check %>% 
  dplyr::select(1, 3) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(stage_1_capacity_check = pcs_comment) -> stage_1_capacity_check_1

data_mfg_locs %>% 
  dplyr::left_join(stage_1_capacity_check_1) %>% 
  dplyr::mutate(stage_1_capacity_check = replace(stage_1_capacity_check, is.na(stage_1_capacity_check) | stage_1_capacity_check == 0, "-")) -> data_mfg_locs

# Stage 1 Capacity Check Date (Col DN)
stage_1_capacity_check %>% 
  dplyr::select(1, 4) %>% 
  janitor::clean_names() %>% 
  dplyr::rename(stage_1_capacity_check_date = pcs_comment_last_updated_date) %>% 
  dplyr::mutate(stage_1_capacity_check_date = as.Date(stage_1_capacity_check_date)) -> stage_1_capacity_check_date

data_mfg_locs %>% 
  dplyr::left_join(stage_1_capacity_check_date) %>% 
  dplyr::mutate(stage_1_capacity_check_date = ifelse(is.na(stage_1_capacity_check_date), 0, stage_1_capacity_check_date)) -> data_mfg_locs


# Component Type2 (Col DO)
comp_type %>% 
  janitor::clean_names() %>% 
  dplyr::rename(component_type_2 = component_type_name) -> comp_type_1


data_mfg_locs %>% 
  dplyr::left_join(comp_type_1) %>% 
  dplyr::mutate(component_type_2 = ifelse(is.na(component_type_2), 0, component_type_2)) -> data_mfg_locs


# Tracker Project Type (Col DP)
data_mfg_locs %>% 
  dplyr::mutate(tracker_project_type = "PCS") -> data_mfg_locs


# Plant (Col DQ)
data_mfg_locs %>% 
  dplyr::mutate(plant = pcs_manufacturing_location) -> data_mfg_locs

# Project Onwer (Col DS)
data_mfg_locs %>% 
  dplyr::mutate(project_owner = requestor) -> data_mfg_locs


# Line (Col DT)
data_mfg_locs %>% 
  dplyr::mutate(line = suggested_line_number) %>% 
  dplyr::mutate(suggested_line_number = ifelse(is.na(suggested_line_number), 0, suggested_line_number)) -> data_mfg_locs

# Duration (# of weeks) (Col EB)
data_mfg_locs %>% 
  dplyr::mutate(duration_number_of_weeks = ifelse(lto_estimated_end_date == "1900-01-01", 0, (lto_estimated_end_date - desired_launch_date)/7),
                duration_number_of_weeks = round(duration_number_of_weeks, 0)) -> data_mfg_locs




##################### (Data2 - No Locs tab) #####################   

# main board (col A to Q)

all_pcs_project_without_location %>% 
  janitor::clean_names() %>% 
  data.frame() -> data2_no_locs


# MPI (Col R)
data2_no_locs %>% 
  dplyr::left_join(mpi) -> data2_no_locs

# Desired Launch Year (Col S)
data2_no_locs %>% 
  dplyr::left_join(desired_launch_year) -> data2_no_locs

# Project Coordinator (Col T)
data2_no_locs %>% 
  dplyr::left_join(project_coordinator) -> data2_no_locs

# Processing Type (Col U)
data2_no_locs %>% 
  dplyr::left_join(processing_type) -> data2_no_locs

# Customer Clean (Col W)
data2_no_locs %>% 
  dplyr::left_join(customer_clean) %>% 
  dplyr::mutate(customer_clean = replace(customer_clean, is.na(customer_clean) | customer_clean == 0, "-")) -> data2_no_locs




##################### (Data3 - Stocking Locs tab) #####################



