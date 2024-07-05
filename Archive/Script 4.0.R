# Config

# Setup for script to run using the R Executable

args <- commandArgs(trailingOnly = FALSE)
script_path <- sub("--file=", "", args[grep("--file=", args)])
script_dir <- normalizePath(dirname(script_path))
setwd(script_dir)

# Sets the working directory one layer higher
getwd()
setwd('..')
getwd()
sink("Source\\output_log.txt", append = TRUE)

library(tidyverse)
library(readxl)
library(openxlsx)
library(xlsx)
library(rlang)

# Form header config. Place identifiers for extractable information here, exactly as they appear in the form content. Include all spaces.
# Please list tags in the order in which they appear in the form. Student first and last name should be the first 2 entries.
tryCatch({
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

#Adds an end tag so the last piece of information isn't truncated.
End_Tag <- "&&&ENDINGTAG&&&"
Tags_Solo <- append(Tags_Solo, End_Tag)

# We want to be able to quickly add escapes for RegEx purposes without labeling our data in a difficult to read format. 
# The following function lets us easily transition easily switch between the escaped and non-escaped formats.
add_escapes <- function(string_in){
  str_replace_all(string_in, pattern = "\\(", replacement = "\\\\(") %>% 
    str_replace_all(pattern = "\\)", replacement = "\\\\)") %>% 
    str_replace_all(pattern = "\\/", replacement = "\\\\/") %>% 
    return()
}

# RegEx list for identifying the endpoint for a specific field.
Tag_List <- str_flatten(Tags_Solo %>% add_escapes(), collapse = "|")

extract_key_to_col <- function(data, start_id, end_ids, reference_col){
  # Extracts data by start and end RegEx key.
  # The Secret Sauce is (?<=Look Behind Pattern)(.*?)(?=Look Ahead Pattern), where (.*?) is lazy (searches for closest start and end patterns)
  start_id_escaped <- start_id %>% add_escapes()
  
  data %>% dplyr::mutate({{start_id}} := 
                           !!sym(reference_col) %>% 
                           str_extract(pattern = paste0("(?<=",start_id_escaped,")(.*?)(?=",end_ids,")")) %>% 
                           str_replace_all(",","") %>% 
                           trimws()) %>% return()
}

# File read, narrowed, entries separated (multiple rows), cleaned for extracting keyed values.
file <- readxl::read_xlsx("Input.xlsx") %>% 
  dplyr::select(Teacher_Name = name, 
                filled_at,
                email, 
                Performance_List = `Solo Performances and Teacher Duets`) %>% 
  separate_rows(Performance_List, sep = paste0("[0-9]+: ",Tags_Solo[1])) %>% 
  dplyr::filter(!Performance_List=="") %>% 
  dplyr::mutate(Performance_List = paste0(Tags_Solo[1],Performance_List,End_Tag))

# Extracting keyed values embedded in a string
for(i in 1:(length(Tags_Solo)-1)){
  file <- file %>% extract_key_to_col(Tags_Solo[i], Tag_List, "Performance_List")
}

# Final processing before export
file <- file %>% dplyr::select(-Performance_List) %>% 
  # Index of tags must be configured to reflect Tag_List contents
  # Name
  # Song 1
  # Song 2
  # Song 3
  dplyr::mutate(Program_Entry = paste0(!!sym(Tags_Solo[1])," ",!!sym(Tags_Solo[2]),": ", 
                                       !!sym(Tags_Solo[4]), " - ", !!sym(Tags_Solo[6]), " / ",
                                       !!sym(Tags_Solo[8]), " - ", !!sym(Tags_Solo[10]), " / ",
                                       !!sym(Tags_Solo[12]), " - ", !!sym(Tags_Solo[14])) %>% 
                  str_replace_all(pattern = " / NA - NA", replacement = ""),
                temp = 1, PK = cumsum(temp)) %>% 
  dplyr::select(PK, everything(), -temp)

# Adding template Excel File. This merges the converted data into a prefabricated model, for immediate use.
# wb <- loadWorkbook(file = "Template.xlsx")
# removeSheet(wb, sheetName = "Master_Input")
# input_pane <- createSheet(wb, sheetName = "Master_Input")
# addDataFrame(file, input_pane)
# saveWorkbook(wb, file = "Output.xlsx"


# Section for bands + ensembles
# ideally, we'd define this as its own function, to not repeat code. However, since this is a budget constricted project, 
# and we anticipate little overlaying functionality will be added later, copying code is justified.
band_tags <- c("Band Name: ",
               "Performers (Full Names, Comma Separated): ",
               "Song1 Title: ",
               "Song1 Length: ",
               "Song1 Composer/Arr.: ",
               "Song1 Staging/Instruments: ",
               "Song2 Title: ",
               "Song2 Length: ",
               "Song2 Composer/Arr.: ",
               "Song2 Staging/Instruments: ")

band_tags <- append(band_tags, End_Tag)
Band_Tag_List <- str_flatten(band_tags %>% add_escapes(), collapse = "|")
  
band_file <- readxl::read_xlsx("Input.xlsx") %>% 
  dplyr::select(Teacher_Name = name, 
                filled_at,
                email, 
                Performance_List = `Bands / Multi-Student Duets (2 Song Max):`) %>% 
  separate_rows(Performance_List, sep = paste0("[0-9]+: ",band_tags[1])) %>% 
  dplyr::filter(!Performance_List=="") %>% 
  dplyr::mutate(Performance_List = paste0(band_tags[1],Performance_List,End_Tag))

for(i in 1:(length(band_tags)-1)){
  band_file <- band_file %>% extract_key_to_col(band_tags[i], Band_Tag_List, "Performance_List")
}

band_file <- band_file %>% dplyr::select(-Performance_List) %>% 
  # Index of tags must be configured to reflect Tag_List contents
  # Name
  # Song 1
  # Song 2
  # Song 3
  dplyr::mutate(Program_Entry = paste0(!!sym(band_tags[1])," ",!!sym(band_tags[2]),": ", 
                                       !!sym(band_tags[4]), " - ", !!sym(band_tags[6]), " / ",
                                       !!sym(band_tags[8]), " - ", !!sym(band_tags[10]), " / ",
                                       !!sym(band_tags[12]), " - ", !!sym(band_tags[14])) %>% 
                  str_replace_all(pattern = " / NA - NA", replacement = ""),
                temp = 1, PK = cumsum(temp)) %>% 
  dplyr::select(PK, everything(), -temp)

# Version 2
wb <- openxlsx::loadWorkbook("Source/Template.xlsx")
writeData(wb, sheet = "Master_Input", file)
openxlsx::saveWorkbook(wb, "Scheduling_Sheet.xlsx", overwrite = FALSE)
}, finally = Sys.sleep(3)
)
