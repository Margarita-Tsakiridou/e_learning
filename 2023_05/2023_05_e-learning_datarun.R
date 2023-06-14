library(tidyverse)

staff <- ".\\2023_05\\02_Staff Counts 31 May 2023.xlsx"
cop <- ".\\2023_05\\codeofpractice Introduction to the Code of Practice for Statistics.xlsx"
qsig <- ".\\2023_05\\QualityStats Quality Statistics in Government.xlsx"
date <- "2023-05-31"


staff_df <- readxl::read_excel(staff, sheet = 2) %>%
  janitor::clean_names() %>%
  select(primary_email_address, group, directorate, division, grade, person_number) %>%
  dplyr::rename(email = primary_email_address, area = group) %>%
  transform(area = as.factor(area), directorate = as.factor(directorate), division = as.factor(division))



#' Formatting the Learning Hub outputs
#'
#' @param course the path to where the output from learning hub lives
#'
#' @return a dataframe with email, attempts, date, completion
#' @export
#'
#' @examples qsig_df <- completion(".\\e-learning\\qsig.xlsx")
#'
completion <- function(course){
  readxl::read_excel(course) %>%
    janitor::clean_names() %>%
    select(email_address, attempt, started_on) %>%
    arrange(email_address, desc(attempt)) %>%
    distinct(email_address, .keep_all = TRUE) %>%
    dplyr::rename(email = email_address) %>%
    filter(attempt>=1) %>%
    separate(started_on, into = c("Date", "Time"), sep = ",") %>%
    transform(Date =as.Date(Date, "%d %B %Y")) %>%
    filter(Date <= as.Date(date)) %>% select(-Time) %>%
    mutate(course = 1)
}

cop_df <- completion(cop) %>% rename(cop=course, cop_attempts = attempt, cop_date = Date)
qsig_df <- completion(qsig) %>% rename(qsig = course, qsig_attempts = attempt, qsig_date = Date)

rm(cop, date,qsig,staff, completion)

#' Merges qsig, cop and staff counts
#'
#' @param staff_df
#' @param cop_df
#' @param qsig_df
#'
#' @return dataframe with area, directorate,division,id, grade, cop and qsig completion for 6k employees
#' @export
#'
#' @examples
#'
merging <- function(staff_df, cop_df, qsig_df){
  fs <- merge(cop_df, qsig_df, all.x = TRUE, all.y = TRUE, by = "email")

  ss <- merge(staff_df, fs, all.x = TRUE, by = "email") %>%
    select(area, directorate,division,email, person_number, grade, cop, qsig) %>%
    arrange(area,directorate,division) %>%
    replace_na(list(cop = 0, qsig = 0))
}

fd <- merging(staff_df, cop_df, qsig_df) %>% select(-email)

rm(cop_df, qsig_df, staff_df, merging)

##############################################################################
#metrics
cop_prc <- round(sum(fd$cop/nrow(fd)*100),2)
qsig_prc <- round(sum(fd$qsig/nrow(fd)*100), 2)
all_ONS <- data.frame("month" = "May_2023","cop" = cop_prc, "qsig" = qsig_prc)
rm(cop_prc,qsig_prc)
###############################################################################


#' List with various breakdowns
#'
#' @param fd
#'
#' @return list with breakdowns by ONS areas and grades
#' @export
#'
#' @examples
breakdowns <- function(fd){

group <- fd %>%
  group_by(area) %>%
  summarize(cop_prc = round(mean(cop)*100,2),
            qsig_prc = round(mean(qsig)*100,2))


directorate <- fd %>%
  group_by(area, directorate) %>%
  summarize(cop_prc = round(mean(cop)*100,2),
           qsig_prc = round(mean(qsig)*100,2))

divisions <- fd %>%
  group_by(area, directorate, division) %>%
  summarize(cop_prc = round(mean(cop)*100,2),
            qsig_prc = round(mean(qsig)*100,2))

grade <- fd %>%
  group_by(grade) %>%
  summarize(cop_prc = round(mean(cop)*100,2),
            qsig_prc = round(mean(qsig)*100,2))

breakdowns <- list(group = group,
                   directorate = directorate,
                   divisions = divisions,
                   grade = grade)

}

breakdown_list <- breakdowns(fd)
rm(breakdowns)


###############################################################################


export <- list("May_metrics" = all_ONS,
      "by_group" = breakdown_list[["group"]],
      "by_directorate" = breakdown_list[["directorate"]],
      "by_divisions" = breakdown_list[["divisions"]],
      "by_grade" = breakdown_list[["grade"]],
      "all_data" = fd)

writexl::write_xlsx(export, ".\\2023_05\\May23_export.xlsx")
