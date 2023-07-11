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
all_pcs_projects_with_mfg_locations <- 




################# ETL (Extract, Transform, Load) ###################

projectlist[-1:-5, -1] -> projectlist
colnames(projectlist) <- projectlist[1, ]
projectlist[-1, ] -> projectlist

projectlist %>% 
  janitor::clean_names() %>% 
  dplyr::select(project_id, project_coordinator) %>% 
  readr::type_convert() -> projectlist


# 6:50
