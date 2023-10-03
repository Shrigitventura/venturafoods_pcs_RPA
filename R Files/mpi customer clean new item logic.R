pre_final_product$project_submitted_date



pre_final_product %>% 
  dplyr::mutate(project_submitted_date = mdy(project_submitted_date)) %>% 
  dplyr::filter(project_submitted_date >= as.Date("2023-09-25") & project_submitted_date <= as.Date("2023-10-01")) -> a
