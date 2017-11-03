# Created by Andrew Russell, 2017.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

## This script check to see if previous Helpline suggested changes were made in the CRM

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions

wd = dirname(getwd())

## Fetch previous data ##
Prev_AK <- read_excel(file.choose(), 1, col_types = "text")

## Fetch current data ##
Cur_AK <- read_excel(file.choose(), 1, col_types = "text")

## Anti-join to check if the rows from the previous file are in the current file
No_match <- anti_join(Prev_AK, Cur_AK)


