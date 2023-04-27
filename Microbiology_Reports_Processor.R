library(pdfsearch)
library(tidyverse)
library(tabulizer)
library(openxlsx)

# Suggested Directory Format:
  # Microbiology Reports
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


working_directory <- "directory to folders of reports by year"
year <- "folder year/"

MB_report_file_names <- paste0(working_directory, "/Microbiology Reports/", year) %>%
  list.files()

MB_report_file_names <- MB_report_file_names[MB_report_file_names != "skipped"]

contact_coords <- list(c(183, 346, 269, 570))
# contact_coords <- paste0(working_directory, "/Microbiology Reports/", year,
#                             MB_report_file_names[1]) %>%
#   locate_areas()

date_coords <- list(c(282, 268, 317, 387))
# date_coords <- paste0(working_directory, "/Microbiology Reports/", year,
#                          MB_report_file_names[1]) %>%
#   locate_areas()

details_coords_1 <- list(c(317, 48, 355, 258))
# details_coords_1 <- paste0(working_directory, "/Microbiology Reports/",
#                               year, MB_report_file_names[87]) %>%
#   locate_areas()

details_coords_2 <- list(c(318, 269, 355, 449))
# details_coords_2 <- paste0(working_directory, "/Microbiology Reports/",
#                               year, MB_report_file_names[55]) %>%
#   locate_areas()

details_coords_3 <- list(c(335, 455, 364, 575))
# details_coords_3 <- paste0(working_directory, "/Microbiology Reports/",
#                               year, MB_report_file_names[55]) %>%
#   locate_areas()

comment_coords <- list(c(361, 149, 409, 535))
# comment_coords <- paste0(working_directory, "/Microbiology Reports/",
#                          year, MB_report_file_names[184]) %>%
#   locate_areas()

profile_coords <- list(c(428, 105, 450, 266))
# profile_coords <- paste0(working_directory, "/Microbiology Reports/",
#                               year, MB_report_file_names[1]) %>%
#   locate_areas()

results_coords <- list(c(446, 50, 519, 569))
# results_coords <- paste0(working_directory, "/Microbiology Reports/", year,
#                             MB_report_file_names[1]) %>%
#   locate_areas()

mb_samples_list <- MB_report_file_names %>%
  map(function(x) {
    print(paste0("Processing: ", x))
    
    dates_df <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(area = date_coords, guess = FALSE) %>%
      data.frame() %>%
      pivot_wider(names_from = "X1", values_from = "X2")
    dates_df <- data.frame("PDF" = x) %>%
      bind_cols(dates_df)

    contact_df <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(area = contact_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)
    contact_df <- contact_df[[1]]
    contact_df <- data.frame("Name" = contact_df[1,],
                             "Street Address" = contact_df[2,],
                             "City State Zipcode" = contact_df[3,],
                             check.names = FALSE)
    
    details_df_1 <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_1, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_1)) {
      details_df_1 <- details_df_1 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_1 <- data.frame("ES Microbiology ID:" = "na",
                                 "GPS Number:" = "na",
                                 check.names = FALSE)
    }
    
    details_df_2 <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_2, guess = FALSE) %>%
      data.frame()
    if ("X1" %in% colnames(details_df_2)) {
      details_df_2 <- details_df_2 %>%
        pivot_wider(names_from = "X1", values_from = "X2")
    } else {
      details_df_2 <- data.frame("Sample Source:" = "na",
                                 "Sampling Point:" = "na",
                                 check.names = FALSE)
    }
    
    details_df_3 <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(pages = 1, area = details_coords_3, output = "character",
                     guess = FALSE)
    details_df_3 <- data.frame("Well Permit Number:" = unlist(details_df_3), check.names = FALSE)
    
    comment <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(pages = 1, area = comment_coords, output = "character",
                     guess = FALSE)
    comment_df <- data.frame("Comment" = unlist(comment), check.names = FALSE)
    
    profile <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(pages = 1, area = profile_coords, output = "character",
                     guess = FALSE)
    profile_df <- data.frame("Profile" = unlist(profile), check.names = FALSE)

    results_df <- paste0(working_directory, "/Microbiology Reports/", year, x) %>%
      extract_tables(area = results_coords, output = "data.frame",
                     guess = FALSE, check.names = FALSE)

    results_df <- results_df[[1]] %>%
      select("Analyte", "Test Result") %>%
      pivot_wider(names_from = "Analyte", values_from = "Test Result")
    colnames(results_df) <- c("Total Coliform", "E. coli")

    sample_row <- bind_cols(dates_df, contact_df, profile_df, details_df_2, 
                            details_df_1, details_df_3, results_df, comment_df)
    print(sample_row)
    return(sample_row)
    
  })

print(mb_samples_list)
MB_tests_df <- mb_samples_list %>%
  bind_rows()
print(MB_tests_df)


excelname <- "Microbiology.xlsx"
write.xlsx(MB_tests_df, 
           paste0(working_directory, "/Microbiology Reports Data/", excelname),
           overwrite = FALSE)

