##### installing tabulizer ####
install.packages("rJava")
library(rJava)
install.packages("devtools")
devtools::install_github("ropensci/tabulizer", args="--no-multiarch")

# Profiles:
#     New Well I
#     Nitrate_Nitrite
#     Inorganic Chemical + Metals I
#     Lead - Water Investigation
#     Default

# Suggested Directory Format:
  # Inorganic Chemistry Reports
  #    2009
  #       Report_1.pdf
  #       Report_2.pdf
  #       Report_3.pdf
  #    2010
  #       Report_1.pdf
  #       Report_2.pdf
  #       Report_3.pdf
  #    2011
  #       Report_1.pdf
  #       Report_2.pdf
  #       Report_3.pdf
  #    2012
  #       Report_1.pdf
  #       Report_2.pdf
  #       Report_3.pdf


#### load packages ####
library(pdfsearch)
library(tidyverse)
library(tabulizer)
library(openxlsx)

working_directory <- "directory to folders of reports by year"
year <- "folder year/"


#### Retrieve Lead - Water Investigation Reports ####
lwi_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Profile: Lead - Water Investigation")
lwi_reports_file_names <- lwi_reports_query$pdf_name %>%
  as.list()


#### Retrieve Reports with Comments ####
commented_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "comment", ignore_case = TRUE)
commented_reports_query <- commented_reports_query %>%
  filter(!(pdf_name %in% lwi_reports_file_names))
commented_reports_file_names <- commented_reports_query$pdf_name %>%
  as.list()


#### Retrieve Nitrate_Nitrite Reports ####
nn_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Profile: Nitrate_Nitrite")
nn_reports_query <- nn_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
nn_reports_file_names_v11 <- nn_reports_query$pdf_name %>%
  as.list()

nn_date_coords <- list(c(247, 196, 291, 332))
# coords <- paste0(first file in list) %>% locate_areas()
# nn_date_coords <- paste0("Inorganic Chemistry Reports/", year, nn_reports_file_names[1]) %>%
#   locate_areas() #246.8144, 196.3047, 290.6925, 332.3269

nn_contact_coords <- list(c(133, 387, 238, 569))
# nn_contact_coords <- paste0("Inorganic Chemistry Reports/", year, nn_reports_file_names[1]) %>%
#   locate_areas() #132.7313, 387.1745, 238.0388, 569.2687

# nn_results_coords <- list(c(363, 27, 407, 589))
nn_results_coords <- list(c(400, 26, 455, 586))
# nn_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                             year, nn_reports_file_names[1]) %>%
#   locate_areas() #363.09141, 27.37396, 406.96953, 589.01385

nn_details_coords_1 <- list(c(288, 15, 330, 167))
# nn_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_details_coords_2 <- list(c(291, 191, 324, 342))
# nn_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_details_coords_3 <- list(c(293, 388, 332, 557))
# nn_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_samples_list <- nn_reports_file_names %>%
  map(function(x) {
    print(paste0("Processing: ", x))

    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)

    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)
    
    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nn_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit No." = "na", "GPS Number:" = "na",
                                 check.names = FALSE)
    }
    
    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Test Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Test Result")

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2, results_df)
    print(sample_row)
    return(sample_row)
    
  })

print(nn_samples_list)
nn_tests_df <- nn_samples_list %>%
  bind_rows()
print(nn_tests_df)

nn_excelname <- "Nitrate_Nitrite_2018.xlsx"
write.xlsx(nn_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", nn_excelname),
           overwrite = FALSE)


#### Retrieve Nitrate_Nitrite Reports ####
nn_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Nitrate_Nitrite \\(Profile\\)", ignore_case = TRUE)
nn_reports_query <- nn_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
nn_reports_file_names_v9 <- nn_reports_query$pdf_name %>%
  as.list()

nn_date_coords <- list(c(245, 248, 278, 377))
# nn_date_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                          year, nn_reports_file_names[1]) %>%
#   locate_areas() #246.8144, 196.3047, 290.6925, 332.3269

nn_contact_coords <- list(c(144, 375, 238, 533))
# nn_contact_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                             year, nn_reports_file_names[1]) %>%
#   locate_areas() #132.7313, 387.1745, 238.0388, 569.2687

nn_results_coords <- list(c(376, 47, 445, 579))
# nn_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                             year, nn_reports_file_names_v9[1]) %>%
#   locate_areas() #363.09141, 27.37396, 406.96953, 589.01385

nn_details_coords_1 <- list(c(284, 53, 315, 186))
# nn_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_details_coords_2 <- list(c(284, 208, 313, 355))
# nn_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_details_coords_3 <- list(c(283, 400, 316, 536))
# nn_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nn_reports_file_names[1]) %>%
#   locate_areas()

nn_comment_coords <- list(c(318, 143, 359, 563))
# nn_comment_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                             year, nn_reports_file_names_v9[1]) %>%
#   locate_areas(pages = 1)

nn_samples_list <- nn_reports_file_names_v9 %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")

    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")

    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nn_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit #:" = "na", "GPS #:" = "na",
                                 check.names = FALSE)
    }

    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = nn_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    print(results_df)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Result")
    results_df$Nitrate <- results_df$Nitrate %>%
      as.character()
    results_df$Nitrite <- results_df$Nitrite %>%
      as.character()
    
    comment <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nn_comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)
    
  })

print(nn_samples_list)
nn_tests_df <- nn_samples_list %>%
  bind_rows()
print(nn_tests_df)

nn_excelname <- "Nitrate_Nitrite.xlsx"
write.xlsx(nn_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", nn_excelname),
           overwrite = FALSE)


#### Retrieve New Well I Reports ####
nw_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Profile: New Well I")
nw_reports_query <- nw_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
nw_reports_file_names <- nw_reports_query$pdf_name %>%
  as.list()

nw_date_coords <- list(c(245, 195, 295, 342))
# nw_date_coords <- paste0("Inorganic Chemistry Reports/", year, nw_reports_file_names[1]) %>%
#   locate_areas(pages = 1)

nw_contact_coords <- list(c(135, 388, 236, 548))
# nw_contact_coords <- paste0("Inorganic Chemistry Reports/", year, nw_reports_file_names[1]) %>%
#   locate_areas(pages = 1)

nw_details_coords_1 <- list(c(288, 15, 330, 167))
# nw_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names[1]) %>%
#   locate_areas()

nw_details_coords_2 <- list(c(291, 191, 324, 342))
# nw_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names[1]) %>%
#   locate_areas()

nw_details_coords_3 <- list(c(293, 388, 332, 557))
# nw_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names[1]) %>%
#   locate_areas()

# nw_results_coords <- list(c(365, 24, 679, 581))
nw_results_coords <- list(c(405, 31, 714, 584))
# nw_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                             year, nw_reports_file_names[1]) %>%
#   locate_areas(pages = 1)

nw_comment_coords <- list(c(354, 75, 378, 408))
# nw_comment_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                             year, nw_reports_file_names[1]) %>%
#   locate_areas(pages = 1)

nw_samples_list <- nw_reports_file_names %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")

    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit No." = "na", "GPS Number:" = "na",
                                 check.names = FALSE)
    }
   
    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Test Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Test Result")
    
    comment <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)
    
  })

print(nw_samples_list)
nw_tests_df <- nw_samples_list %>%
  bind_rows()
print(nw_tests_df)

nw_excelname <- "New_Well_I_2018.xlsx"
write.xlsx(nw_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", nw_excelname),
           overwrite = FALSE)


#### Retrieve New Well I Reports ####
nw_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "New Well \\(Profile\\)")
nw_reports_query <- nw_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
nw_reports_file_names_v9 <- nw_reports_query$pdf_name %>%
  as.list()

nw_date_coords <- list(c(249, 247, 278, 379))
# nw_date_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                          year, nw_reports_file_names_v9[1]) %>%
#   locate_areas(pages = 1)

nw_contact_coords <- list(c(145, 371, 233, 560))
# nw_contact_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                             year, nw_reports_file_names_v9[1]) %>%
#   locate_areas(pages = 1)

nw_details_coords_1 <- list(c(285, 54, 312, 190))
# nw_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names_v9[1]) %>%
#   locate_areas()

nw_details_coords_2 <- list(c(284, 209, 316, 356))
# nw_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names_v9[1]) %>%
#   locate_areas()

nw_details_coords_3 <- list(c(284, 403, 309, 561))
# nw_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, nw_reports_file_names_v9[1]) %>%
#   locate_areas()

nw_results_coords <- list(c(371, 49, 689, 566))
# nw_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                             year, nw_reports_file_names_v9[1]) %>%
#   locate_areas(pages = 1)

nw_comment_coords <- list(c(318, 142, 361, 565))
# nw_comment_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                             year, nw_reports_file_names_v9[22]) %>%
#   locate_areas(pages = 1)

nw_samples_list <- nw_reports_file_names_v9 %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)

    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")

    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")

    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit #:" = "na", "GPS #:" = "na",
                                 check.names = FALSE)
    }

    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Result")

    comment <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = nw_comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)
    
  })

print(nw_samples_list)
nw_tests_df <- nw_samples_list %>%
  bind_rows()
print(nw_tests_df)

nw_excelname <- "New_Well_I.xlsx"
write.xlsx(nw_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", nw_excelname),
           overwrite = FALSE)


#### Retrieve Inorganic Chemical + Metals I Reports ####
icm_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Profile: Inorganic Chemical \\+ Metals I")
icm_reports_query <- icm_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
icm_reports_file_names <- icm_reports_query$pdf_name %>%
  as.list()

icm_date_coords <- list(c(247, 194, 293, 330))
# icm_date_coords <- paste0("Inorganic Chemistry Reports/", year, icm_reports_file_names[1]) %>%
#   locate_areas()

icm_contact_coords <- list(c(133, 383, 238, 569))
# icm_contact_coords <- paste0("Inorganic Chemistry Reports/", year, icm_reports_file_names[1]) %>%
#   locate_areas()

icm_details_coords_1 <- list(c(288, 15, 330, 167))
# icm_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names[1]) %>%
#   locate_areas()

icm_details_coords_2 <- list(c(291, 191, 324, 342))
# icm_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names[1]) %>%
#   locate_areas()

icm_details_coords_3 <- list(c(293, 388, 332, 557))
# icm_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names[1]) %>%
#   locate_areas()

# icm_results_coords <- list(c(365, 23, 657, 582))
icm_results_coords <- list(c(406, 26, 688, 584))
# icm_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                              year, icm_reports_file_names[1]) %>%
#   locate_areas()

icm_samples_list <- icm_reports_file_names %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)
    
    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit No." = "na", "GPS Number:" = "na",
                                 check.names = FALSE)
    }
    
    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Test Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Test Result")

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df)
    print(sample_row)
    return(sample_row)
    
  })

print(icm_samples_list)
icm_tests_df <- icm_samples_list %>%
  bind_rows()
print(icm_tests_df)

icm_excelname <- "Inorganic_Chemical_Metals_I_2018.xlsx"
write.xlsx(icm_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", icm_excelname),
           overwrite = FALSE)


#### Retrieve Inorganic Chemical + Metals I Reports ####
icm_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Inorganic Chemical \\+ Metals \\(Profile\\)")
icm_reports_query <- icm_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
icm_reports_file_names_v9 <- icm_reports_query$pdf_name %>%
  as.list()

icm_date_coords <- list(c(247, 246, 280, 374))
# icm_date_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                           year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_contact_coords <- list(c(145, 374, 233, 543))
# icm_contact_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", 
#                              year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_details_coords_1 <- list(c(285, 52, 312, 179))
# icm_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_details_coords_2 <- list(c(283, 209, 314, 372))
# icm_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_details_coords_3 <- list(c(284, 404, 309, 547))
# icm_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_results_coords <- list(c(373, 51, 680, 562))
# icm_results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                              year, icm_reports_file_names_v9[1]) %>%
#   locate_areas()

icm_comment_coords <- list(c(318, 144, 357, 560))
# icm_comment_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                             year, icm_reports_file_names_v9[19]) %>%
#   locate_areas(pages = 1)

icm_samples_list <- icm_reports_file_names_v9  %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)
    
    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_1, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_1)) {
      details_df_1 <- details_df_1 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_1 <- data.frame("Sample Type:" = "na", "Sample Source:" = "na",
                                 check.names = FALSE)
    }
    
    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_2, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_2)) {
      details_df_2 <- details_df_2 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_2 <- data.frame("Sampling Point:" = "na", "Temp. at Receipt:" = "na",
                                 check.names = FALSE)
    }
    
    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit #:" = "na", "GPS #:" = "na",
                                 check.names = FALSE)
    }
    
    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = icm_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Result")

    comment <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = icm_comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)
    
  })

print(icm_samples_list)
icm_tests_df <- icm_samples_list %>%
  bind_rows()
print(icm_tests_df)

icm_excelname <- "Inorganic_Chemical_Metals_I.xlsx"
write.xlsx(icm_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", icm_excelname),
           overwrite = FALSE)


#### Retrieve Default Reports ####
d_reports_query <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  keyword_directory(keyword = "Profile: Default")
d_reports_query <- d_reports_query %>%
  filter(!(pdf_name %in% commented_reports_file_names))
d_reports_file_names <- d_reports_query$pdf_name %>%
  as.list()

d_date_coords <- list(c(245, 194, 293, 343))
# d_date_coords <- paste0("Inorganic Chemistry Reports/", year, d_reports_file_names[1]) %>%
#   locate_areas()

d_contact_coords <- list(c(131, 385, 234, 571))
# d_contact_coords <- paste0("Inorganic Chemistry Reports/", year, d_reports_file_names[1]) %>%
#   locate_areas()

d_details_coords_1 <- list(c(288, 15, 330, 167))
# d_details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, d_reports_file_names[1]) %>%
#   locate_areas()

d_details_coords_2 <- list(c(291, 191, 324, 342))
# d_details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, d_reports_file_names[1]) %>%
#   locate_areas()

d_details_coords_3 <- list(c(293, 388, 332, 557))
# d_details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/",
#                               year, d_reports_file_names[1]) %>%
#   locate_areas()

d_results_coords <- list(c(365, 21, 449, 589))
# d_results_coords <- paste0("Inorganic Chemistry Reports/", year, d_reports_file_names[1]) %>%
#   locate_areas()

d_samples_list <- d_reports_file_names %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = d_date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = d_contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = d_details_coords_1, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = d_details_coords_2, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    
    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = d_details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit No." = "na", "GPS Number:" = "na",
                                 check.names = FALSE)
    }

    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(area = d_results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Test Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Test Result")

    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df)
    print(sample_row)
    return(sample_row)
    
  })

print(d_samples_list)
d_tests_df <- d_samples_list %>%
  bind_rows()
print(d_tests_df)

d_excelname <- "Default_2022.xlsx"
write.xlsx(d_tests_df, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", d_excelname),
           overwrite = FALSE)


##### Look for skipped reports ####
IC_report_file_names <- paste0(working_directory, "/Inorganic Chemistry Reports/", year) %>%
  list.files()
IC_report_file_names <- IC_report_file_names[IC_report_file_names != "skipped"]

# commented_reports_file_names,
IC_processed_reports <- list(lwi_reports_file_names, nn_reports_file_names_v9,
                             nn_reports_file_names_v11, nw_reports_file_names,
                             nw_reports_file_names_v9, nw_reports_file_names_v9_2,
                             icm_reports_file_names, 
                             icm_reports_file_names_v9, icm_reports_file_names_v9_2,
                             d_reports_file_names) %>% 
  unlist()

IC_skipped_reports <- IC_report_file_names[!(IC_report_file_names %in% IC_processed_reports)]

# Review skipped reports and commmented reports
print(commented_reports_file_names)

# commented_reports_file_names
commented_samples_list <- IC_skipped_reports %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    print("Highlight Dates")
    date_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)

    print("Highlight Name and Address (include 'Name of System')")
    contact_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)

    print("Highlight Details")
    details_coords_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)

    print("Highlight Details")
    details_coords_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)

    print("Highlight Details")
    details_coords_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)

    print("Highlight Results")
    results_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)
    
    print("Highlight Comment")
    comment_coords <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      locate_areas(pages = 1)
    
    dates_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    contact_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)

    details_df_1 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_1, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_1)) {
      details_df_1 <- details_df_1 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_1 <- data.frame("Sample Type:" = "na", "Sample Source:" = "na",
                                 check.names = FALSE)
    }

    details_df_2 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_2, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_2)) {
      details_df_2 <- details_df_2 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_2 <- data.frame("Sampling Point:" = "na", "Temp. at Receipt:" = "na",
                                 check.names = FALSE)
    }

    details_df_3 <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_3, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_3)) {
      details_df_3 <- details_df_3 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_3 <- data.frame("Well Permit No." = "na", "GPS Number:" = "na",
                                 check.names = FALSE)
    }

    results_df <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    results_df <- results_df[[1]] %>%
      select("Analyte", "Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Result")

    comment <- paste0(working_directory, "/Inorganic Chemistry Reports/", year, x) %>%
      extract_tables(pages = 1, area = comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)

    print(comment_df)
    sample_row <- bind_cols(dates_df, contact_df, details_df_1, details_df_2,
                            details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)

  })

commented_samples_list %>%
  map(view)

commented_excelname <- "Flagged.xlsx"
write.xlsx(commented_samples_list, 
           paste0(working_directory, "/Inorganic Chemistry Reports Data/", commented_excelname),
           overwrite = FALSE)













