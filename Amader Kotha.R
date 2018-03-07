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








library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(openxlsx) # writes Excel files
library(dplyr) # data manipulation

wd = dirname(getwd())

# Read in raw Helpline data, then assign unique ID to Phone #s
H <- read.csv("C:/Users/Andrew/Desktop/NU CRM.csv", 1)
D = as.data.table(H[, c("Factory.FFC.Number", "Reasons", "Call.Date")])

# Create new column that combines Caller Group and Reason
D = mutate(D, `IssueName` = paste(`Factory.FFC.Number`, `Reasons`, `Call.Date`, sep = ""))

comb <- with(D, IssueName)
D <- within(D, Issue_ID <- match(comb, unique(comb)))

H$Issue_ID = D$Issue_ID
H = H[, c("Sl.", "Issue_ID")]

write.csv(H, "Amader Kotha Data with Unique ID for Issues.csv", na="")