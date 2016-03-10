# Created by Andrew Russell, 2016.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions

#### Dashboard Workbook ####
Combined <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Dashboard Workbook.xlsm", "Combined", skip = 2)

# Remove unnecessary columns for Combined
Combined = Combined[, 1:8]

#### Current website statuses ####
Web_old <- read_excel("C:/Users/Andrew/Desktop/Website statuses (copy).xlsx", 1)

# Join old Factory Monthly to Combined
Website = left_join(Combined, Web_old, by = "Factory Name")

# Save the new Factory Monthly list, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Website, "Website statuses.csv", na="")
