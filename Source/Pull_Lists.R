required_packages <- c("tidyverse", "readxl", "openxlsx", "xlsx", "rlang")

check_and_install_packages <- function(packages) {
  for (pkg in packages) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      message("Installing missing package: ", pkg)
      install.packages(pkg)
    }
    library(pkg, character.only = TRUE)
  }
}

check_and_install_packages(required_packages)

# Finds Local Directory (Important for running using the R executable)

args <- commandArgs(trailingOnly = FALSE)
script_path <- sub("--file=", "", args[grep("--file=", args)])
script_dir <- normalizePath(dirname(script_path))
setwd(script_dir)

# WD one folder higher

setwd('..')
#setwd("C:\\Users\\jsp80\\Documents\\FIN 450\\Recital-Automation")

# Libraries

library(tidyverse)
library(readxl)
library(openxlsx)
library(xlsx)
library(rlang)

# Global Config

fileIn <- "Scheduling_Sheet.xlsx"

# Preprocessing. Filling excluded ordering values.
#Indexes used in case the naming convention in the template file changes. 
# Order and indexing of columns must remain consistent.
df1 <- read_excel(fileIn, sheet = "Scheduling") %>% 
  dplyr::filter((!is.na(.[[1]]))&(!.[[1]]=="Max 500 Entries")) %>% 
  dplyr::mutate(time_slot = .[[1]],
                order = .[[2]],
                PK = .[[7]]) %>% 
  dplyr::select(time_slot, order, PK) %>% 
  dplyr::group_by(time_slot) %>% 
  dplyr::mutate(mid = median(as.numeric(order), na.rm = TRUE),
                order = case_when(is.na(order) ~ as.numeric(mid),
                                  TRUE ~ as.numeric(order))) %>% 
  dplyr::arrange(order, .by_group = TRUE) %>% 
  dplyr::mutate(temp = 1, 
                order = cumsum(temp),
                PK = as.numeric(PK)) %>% 
  dplyr::select(-mid, - temp)

df2 <- read_excel(fileIn, sheet = "Master_Input") %>% 
  dplyr::mutate(PK = as.numeric(PK))

master_df <- left_join(df1, df2, by = "PK") %>% ungroup()

# Start point for next session
multisheet_excel <- function(data_in, sep_column, data_conversion = select_none){
  
  wb <- openxlsx::createWorkbook()
  sheetnames <- unique(data_in[[sep_column]])
  
  for(sheetname in sheetnames){
    sheetname_excel <- sheetname %>% str_replace_all(pattern = "\\:", replacement = "_")
    sheetdata <- data_in %>% 
      dplyr::filter(!!sym(sep_column) == sheetname) %>% 
      data_conversion()
    addWorksheet(wb, sheetname_excel)
    writeData(wb, sheet = sheetname_excel, sheetdata)
  }
  return(wb)
}

select_none <- function(df_in){
  return(df_in)
}

# Selection and formatting functions for each sheet style. Accepts df.
select_info_program <- function(df_in){
  df_in %>% dplyr::select(order, 
                          `Auto Entry` = `Program_Entry`,
                          `Band Name:`,
                          `Performers (Full Names, Comma Separated):`, 
                          `Song1 Title:`,
                          `Song1 Composer/Arr.:`,
                          `Song2 Title:`,
                          `Song2 Composer/Arr.:`,
                          `Song3 Title:`,
                          `Song3 Composer/Arr.:` 
                          )
}

select_info_awards <- function(df_in){
  df_in %>% dplyr::select(order, 
                          `Band Name:`,
                          `Performers (Full Names, Comma Separated):`,
                          `Positive Comment / Award:`
  )
}

select_info_tech_sheet <- function(df_in){
  df_in <- df_in %>% dplyr::select(order, 
                          `Band Name:`,
                          `Performers (Full Names, Comma Separated):`,
                          `Song1 Title:`,
                          `Song1 Staging/Instruments:`,
                          `Song2 Title:`,
                          `Song2 Staging/Instruments:`,
                          `Song3 Title:`,
                          `Song3 Staging/Instruments:`
                          
  )
  
  # Pivoting wider, accounting for pairwise columns.
  df_out <- bind_rows(df_in %>% dplyr::select(order, `Band Name:`,`Performers (Full Names, Comma Separated):`, Song = `Song1 Title:`, Staging = `Song1 Staging/Instruments:`) %>% dplyr::mutate(song_num = 1),
                      df_in %>% dplyr::select(order, `Band Name:`,`Performers (Full Names, Comma Separated):`, Song = `Song2 Title:`, Staging = `Song2 Staging/Instruments:`) %>% dplyr::mutate(song_num = 2),
                      df_in %>% dplyr::select(order, `Band Name:`,`Performers (Full Names, Comma Separated):`, Song = `Song3 Title:`, Staging = `Song3 Staging/Instruments:`) %>% dplyr::mutate(song_num = 3)) %>% 
    arrange(order, song_num) %>% 
    dplyr::mutate(Song = trimws(Song)) %>% 
    dplyr::filter(!(is.na(Song)|Song == "")) %>% 
    dplyr::select(-song_num) %>% 
    dplyr::rename(Performer = order)
}

# Calling processing and saving of each filtered sheet.
multisheet_excel(master_df, "time_slot", data_conversion = select_info_program) %>% openxlsx::saveWorkbook(file = "program_list.xlsx", overwrite = TRUE)
multisheet_excel(master_df, "time_slot", data_conversion = select_info_awards) %>% openxlsx::saveWorkbook(file = "award_list.xlsx", overwrite = TRUE)
multisheet_excel(master_df, "time_slot", data_conversion = select_info_tech_sheet) %>% openxlsx::saveWorkbook(file = "tech_sheet.xlsx", overwrite = TRUE)
