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

##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################
##########################################################################################################################################################################


################ Read Data Fixed files ####################
active_pack_comp <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/Active Pack Comp.xlsx")
majorprof_monthyr <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/PCS/Reporting/RStudio/MajorProj MonthYr.xlsx")

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
  dplyr::rename(component = component_type_code) %>% 
  dplyr::mutate(component = replace(component, is.na(component), 0)) -> data_mfg_locs


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
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::rename(mpi = major_initative) -> mpi

data_mfg_locs %>% 
  dplyr::left_join(mpi) -> data_mfg_locs

# Desired launch month (Col BI)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,3) %>% 
  dplyr::rename(desired_launch_month = desired_launch_month_name) -> desired_launch_month

data_mfg_locs %>% 
  dplyr::left_join(desired_launch_month) -> data_mfg_locs


# Desired Launch Year (Col BJ)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,5) -> desired_launch_year

data_mfg_locs %>% 
  dplyr::left_join(desired_launch_year) -> data_mfg_locs


# Project Coordinator (Col BK)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,6) -> project_coordinator

data_mfg_locs %>% 
  dplyr::left_join(project_coordinator) %>% 
  dplyr::mutate(project_coordinator = replace(project_coordinator, is.na(project_coordinator), 0)) -> data_mfg_locs


# Processing Type (Col BL)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,8) -> processing_type

data_mfg_locs %>% 
  dplyr::left_join(processing_type) -> data_mfg_locs


# Submitted Month (Col BP)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,12) %>% 
  dplyr::rename(submitted_month = submitted_month_name) -> submitted_month

data_mfg_locs %>% 
  dplyr::left_join(submitted_month) %>% 
  dplyr::mutate(submitted_month = replace(submitted_month, is.na(submitted_month), 0)) -> data_mfg_locs


# Submitted Year (Col BQ)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,14) -> submitted_year

data_mfg_locs %>% 
  dplyr::left_join(submitted_year) %>% 
  dplyr::mutate(submitted_year = replace(submitted_year, is.na(submitted_year), 0)) -> data_mfg_locs

# Channel 1 Approver (Col BR)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,17) -> channel_1_approver

data_mfg_locs %>% 
  dplyr::left_join(channel_1_approver) %>% 
  dplyr::mutate(channel_1_approver = replace(channel_1_approver, is.na(channel_1_approver), 0)) -> data_mfg_locs


# Channel 2 Approver (Col BS)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,18) -> channel_2_approver

data_mfg_locs %>% 
  dplyr::left_join(channel_2_approver) %>% 
  dplyr::mutate(channel_2_approver = replace(channel_2_approver, is.na(channel_2_approver), 0)) -> data_mfg_locs


# Sales Project Notes (Col BT)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,19) -> sales_project_notes

data_mfg_locs %>% 
  dplyr::left_join(sales_project_notes) %>% 
  dplyr::mutate(sales_project_notes = replace(sales_project_notes, is.na(sales_project_notes) | sales_project_notes == 0, "-")) -> data_mfg_locs


# Sales Project Note Timestamp (Col BU)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,20) -> sales_project_note_timestamp

data_mfg_locs %>% 
  dplyr::left_join(sales_project_note_timestamp) %>% 
  dplyr::mutate(sales_project_note_timestamp = as.Date(sales_project_note_timestamp)) -> data_mfg_locs

# Sales Comment Author (Col BV)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,21) -> sales_comment_author

data_mfg_locs %>% 
  dplyr::left_join(sales_comment_author) %>% 
  dplyr::mutate(sales_comment_author = replace(sales_comment_author, is.na(sales_comment_author) | sales_comment_author == 0, "-")) -> data_mfg_locs

# Latest Comment Type (Col BW)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,22) -> latest_comment_type

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_type) %>% 
  dplyr::mutate(latest_comment_type = replace(latest_comment_type, is.na(latest_comment_type) | latest_comment_type == 0, "-")) -> data_mfg_locs


# Latest Comment (Col BX)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,23) -> latest_comment

data_mfg_locs %>% 
  dplyr::left_join(latest_comment) %>% 
  dplyr::mutate(latest_comment = replace(latest_comment, is.na(latest_comment) | latest_comment == 0, "-")) -> data_mfg_locs

# Latest Comment Date (Col BY)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,24) -> latest_comment_date

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_date) %>% 
  dplyr::mutate(latest_comment_date = as.Date(latest_comment_date)) -> data_mfg_locs

# Latest Comment Author (Col BZ)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,25) -> latest_comment_author

data_mfg_locs %>% 
  dplyr::left_join(latest_comment_author) %>% 
  dplyr::mutate(latest_comment_author = replace(latest_comment_author, is.na(latest_comment_author) | latest_comment_author == 0, "-")) -> data_mfg_locs


# days since Project was created (Col CA)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,29) %>% 
  dplyr::rename(days_since_projecet_was_created = x_days_since_project_was_created) -> days_since_projecet_was_created

data_mfg_locs %>% 
  dplyr::left_join(days_since_projecet_was_created) %>% 
  dplyr::mutate(days_since_projecet_was_created = replace(days_since_projecet_was_created, is.na(days_since_projecet_was_created) | days_since_projecet_was_created == 0, "-")) -> data_mfg_locs


# Actual Comp Month (Col CB)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,32) -> actual_comp_month

data_mfg_locs %>% 
  dplyr::left_join(actual_comp_month) %>% 
  dplyr::mutate(actual_comp_month = replace(actual_comp_month, is.na(actual_comp_month) | actual_comp_month == 0, "-")) -> data_mfg_locs


# Actual Comp Year (Col CC)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,33) -> actual_comp_year

data_mfg_locs %>% 
  dplyr::left_join(actual_comp_year) %>% 
  dplyr::mutate(actual_comp_year = replace(actual_comp_year, is.na(actual_comp_year) | actual_comp_year == 0, "-")) -> data_mfg_locs


# Customer Clean (Col CE)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,34) -> customer_clean

data_mfg_locs %>% 
  dplyr::left_join(customer_clean) %>% 
  dplyr::mutate(customer_clean = replace(customer_clean, is.na(customer_clean) | customer_clean == 0, "-")) -> data_mfg_locs


# Requestor Latest (Col CF)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,35) -> requestor_latest

data_mfg_locs %>% 
  dplyr::left_join(requestor_latest) %>% 
  dplyr::mutate(requestor_latest = replace(requestor_latest, is.na(requestor_latest) | requestor_latest == 0, "-")) -> data_mfg_locs


# Export Region2 (Col CG)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,37) %>% 
  dplyr::rename(export_region_2 = export_region) -> export_region_2

data_mfg_locs %>% 
  dplyr::left_join(export_region_2) %>% 
  dplyr::mutate(export_region_2 = replace(export_region_2, is.na(export_region_2) | export_region_2 == 0, "-")) -> data_mfg_locs


# Project Type Shortened (Col CM)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,7) -> project_type_shortened
  
data_mfg_locs %>% 
  dplyr::left_join(project_type_shortened) %>% 
  dplyr::mutate(project_type_shortened = replace(project_type_shortened, is.na(project_type_shortened) | project_type_shortened == 0, "-")) -> data_mfg_locs


# PRM Date (Col CO)
majorprof_monthyr %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::select(1,38) -> prm_date

data_mfg_locs %>% 
  dplyr::left_join(prm_date) %>% 
  dplyr::mutate(prm_date = as.Date(prm_date)) -> data_mfg_locs




