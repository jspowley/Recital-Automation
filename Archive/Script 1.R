library(tidyverse)
library(readxl)

Tag_List <- "First Name: |Last Name: |Song1 Title: |Song1 Length \\(minutes\\): |Song1 Composer\\/Arr.: |Song2 Title: |Song2 Length \\(minutes\\): |Song2 Composer\\/Arr.: |Song3 Title: |Song3 Length \\(minutes\\): |Song3 Composer\\/Arr.: |&&&ENDINGTAG&&&"

# The Secret Sauce is (?<=Look Behind Pattern)(.*?)(?=Look Ahead Pattern)
data_extract <- function(var_in, start_id, end_ids){
  output <- str_extract(var_in, 
                     pattern = paste0("(?<=",start_id,")(.*?)(?=",end_ids,")")) %>% 
           str_replace(",","")
return(output)
}

file <- readxl::read_xlsx("Input.xlsx") %>% 
  dplyr::select(id, 
                Teacher_Name = name, 
                filled_at, 
                email, 
                Performance_List = `Solo Performances and Teacher Duets (Style 1)`) %>% 
  #dplyr::mutate(entry = str_split(Performance_List, pattern = "[0-9]+: ")) %>% 
  separate_rows(Performance_List, sep = "[0-9]+: ") %>%
  dplyr::filter(!Performance_List=="") %>% 
  dplyr::mutate(Performance_List = paste0(Performance_List,"&&&ENDINGTAG&&&")) %>% 
  dplyr::mutate(First_Name = data_extract(Performance_List, "First Name: ", Tag_List),
                Last_Name = data_extract(Performance_List, "Last Name: ", Tag_List),
                Song1 = data_extract(Performance_List, "Song1 Title: ", Tag_List),
                Song1_Length = data_extract(Performance_List, "Song1 Length \\(minutes\\): ", Tag_List),
                Song1_Composer_Arr = data_extract(Performance_List, "Song1 Composer\\/Arr.: ", Tag_List),
                Song2 = data_extract(Performance_List, "Song2 Title:", Tag_List),
                Song2_Length = data_extract(Performance_List, "Song2 Length \\(minutes\\): ", Tag_List),
                Song2_Composer_Arr = data_extract(Performance_List, "Song2 Composer\\/Arr.: ", Tag_List),
                Song3 = data_extract(Performance_List, "Song3 Title:", Tag_List),
                Song3_Length = data_extract(Performance_List, "Song3 Length \\(minutes\\): ", Tag_List),
                Song3_Composer_Arr = data_extract(Performance_List, "Song3 Composer\\/Arr.: ", Tag_List))

                