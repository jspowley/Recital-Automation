# Check dependencies and install if necessary.
# Credit to ChatGPT

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
# setwd("C:\\Users\\jsp80\\Documents\\FIN 450\\Recital-Automation")

# Libraries

library(tidyverse)
library(readxl)
library(openxlsx)
library(xlsx)
library(rlang)

# Global Config

fileIn <- "Input.xlsx"

# Configuring parsing tags

Tags_Solo <- c("First Name: ",
               "Last Name: ",
               "Positive Comment / Award: ",
               "Song1 Title: ",
               "Song1 Length (minutes): ",
               "Song1 Composer/Arr.: ",
               "Song1 Staging/Instruments: ",
               "Song2 Title: ",
               "Song2 Length (minutes): ",
               "Song2 Composer/Arr.: ",
               "Song2 Staging/Instruments: ",
               "Song3 Title: ",
               "Song3 Length (minutes): ",
               "Song3 Composer/Arr.: ",
               "Song3 Staging/Instruments: ")

Tags_Band <- c("Band Name: ",
               "Performers (Full Names, Comma Separated): ",
               "Positive Comment / Award: ",
               "Song1 Title: ",
               "Song1 Length (minutes): ",
               "Song1 Composer/Arr.: ",
               "Song1 Staging/Instruments: ",
               "Song2 Title: ",
               "Song2 Length (minutes): ",
               "Song2 Composer/Arr.: ",
               "Song2 Staging/Instruments: ")

add_escapes <- function(string_in){
  # Function for adding escape characterss for literal RegEx interpretation. Covers ()/.
  str_replace_all(string_in, pattern = "\\(", replacement = "\\\\(") %>% 
    str_replace_all(pattern = "\\)", replacement = "\\\\)") %>% 
    str_replace_all(pattern = "\\/", replacement = "\\\\/") %>% 
    return()
}

# Adding an identifier for the last piece of information in an entry.
# Flattening the list into RegEx pattern format.
End_Tag <- "&&&ENDINGTAG&&&"

Tags_Solo <- append(Tags_Solo, End_Tag)
Tag_List_Solo <- str_flatten(Tags_Solo %>% add_escapes(), collapse = "|")

Tags_Band <- append(Tags_Band, End_Tag)
Tag_List_Band <- str_flatten(Tags_Band %>% add_escapes(), collapse = "|")

# Parsing

extract_key_to_col <- function(data, start_id, end_ids, reference_col){
  # Extracts data by start and end RegEx key.
  # The Secret Sauce is (?<=Look Behind Pattern)(.*?)(?=Look Ahead Pattern), where (.*?) is lazy (searches for closest start and end patterns)
  start_id_escaped <- start_id %>% add_escapes()
  
  data <- data %>% dplyr::mutate({{start_id}} := 
                                   !!sym(reference_col) %>%
                                   str_extract(pattern = paste0("(?<=",start_id_escaped,")(.*?)(?=",end_ids,")")) %>% 
                                   trimws() %>% 
                                   #remove end commas only
                                   str_sub(1,-2) %>% 
                                   trimws()) %>% 
    return()
}

separate_multi_table <- function(filename, target_col, tags_in, tags_regex){
  # filename is the local path to the excel file
  # target_col is the column containing the unparsed multiline table
  # Can now accept cases where early columns are ommited, or when a single row is used.
  output <- readxl::read_xlsx(filename) %>% 
    dplyr::select(Teacher_Name = name, 
                  filled_at,
                  email, 
                  column = !!sym(target_col)) %>% 
    dplyr::mutate(column = column %>% 
                   str_replace_all(
                   pattern = paste0("[0-9]+: (",tags_regex,")"),
                   function(m) paste0("&&&Split&&&", m)
                   )
                ) %>% 
    separate_rows(column, sep = paste0("&&&Split&&&")) %>% 
    dplyr::filter(!column=="") %>% 
    dplyr::mutate(column = paste0(column,",",End_Tag))
  
  for(i in 1:(length(tags_in)-1)){
    output <- output %>% extract_key_to_col(tags_in[i], tags_regex, "column")
  }
  
  output <- output %>% dplyr::select(-column)
  
  return(output)
}

# Parsing Solo Performances

solo <- separate_multi_table(fileIn, "Solo Performances and Teacher Duets", Tags_Solo, Tag_List_Solo) %>% 
  dplyr::mutate(Program_Entry = paste0(!!sym(Tags_Solo[1])," ",!!sym(Tags_Solo[2]),": ", !!sym(Tags_Solo[4]), " - ", !!sym(Tags_Solo[6]), " / ",!!sym(Tags_Solo[8]), " - ", !!sym(Tags_Solo[10]), " / ",!!sym(Tags_Solo[12]), " - ", !!sym(Tags_Solo[14])) %>% 
                  str_replace_all(pattern = " / NA - NA", replacement = ""),
                #Ensures band and solo performers names' are under the same column convention.
                `Performers (Full Names, Comma Separated): `=paste0(`First Name: `," ",`Last Name: `),
                Type = "Solo") %>% 
  dplyr::select(-`First Name: `,-`Last Name: `)

# Parsing Band Performances

band <- separate_multi_table(fileIn, "Bands / Multi-Student Duets (2 Song Max):", Tags_Band, Tag_List_Band) %>% 
  dplyr::mutate(Program_Entry = paste0(!!sym(Tags_Band[1]),": ",!!sym(Tags_Band[2]),"; ", !!sym(Tags_Band[4]), " - ", !!sym(Tags_Band[6]), " / ",!!sym(Tags_Band[8]), " - ", !!sym(Tags_Band[10])) %>% 
                  str_replace_all(pattern = " / NA - NA", replacement = "") %>% 
                  str_replace_all(pattern = " NA - NA", replacement = ""),
                Type = "Band")

# Merging, Primary Keying, Ordering

data_out <- bind_rows(solo,band) %>% 
  dplyr::mutate(temp = 1, PK = cumsum(temp)) %>% 
  dplyr::select(-temp) %>% 
  dplyr::select(PK, Teacher_Name, filled_at, email, Type, `Performers (Full Names, Comma Separated): `,`Band Name: `, everything())

# Saving to template xlsx

wb <- openxlsx::loadWorkbook("Source/Template.xlsx")
writeData(wb, sheet = "Master_Input", data_out)
openxlsx::saveWorkbook(wb, "Scheduling_Sheet.xlsx", overwrite = FALSE)