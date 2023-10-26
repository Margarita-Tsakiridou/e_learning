#Introduction
#What this code does

## packages
library(tidyverse)


##set up

if (!dir.exists("./1.input")){
  dir.create("./1.input")
}


##load config

config<-config::get()



## Load data

staff_df <- readxl::read_excel(config[["staff_counts"]], sheet = 2) %>%
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



