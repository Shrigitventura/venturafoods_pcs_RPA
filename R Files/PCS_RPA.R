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




