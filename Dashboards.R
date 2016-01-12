# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions
# library(reshape2) # a few pivot-table functions


#### Fetch the Remediation Workbook spreadsheets and put the results in dataframes ####
# CAPs <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 2, skip = 1)
CAPs_pivot <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 1, skip = 3)
# Remove Grand Total row and change column names
CAPs_pivot = CAPs_pivot[1:nrow(CAPs_pivot)-1,]
setnames(CAPs_pivot, c(2:16,18:32,34:48), 
         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total",
           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total",
           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total"))
setnames(CAPs_pivot, c(2:20,22:41,43:62), 
         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total", "Electrical - No Level - Completed", "Electrical - No Level - In progress - on track", "Electrical - No Level - Not started", "Electrical - No Level - Total",
           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total", "Fire - No Level - Completed", "Fire - No Level - In progress - on track", "Fire - No Level - In progress - not on track", "Fire - No Level - Not started", "Fire - No Level - Total",
           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total", "Structural - No Level - Completed", "Structural - No Level - In progress - on track", "Structural - No Level - In progress - not on track", "Structural - No Level - Not started", "Structural - No Level - Total"))
# Fetch statuses of each RVV and CCVV, remove Grand Total row, and change column names
CAPs_RVVs <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 2, skip = 1)
RVV1 = CAPs_RVVs[,1:5]
setnames(RVV1, names(RVV1), c("Account ID", "Completed - RVV1", "In progress - on track - RVV1", "In progress - not on track - RVV1", "Not started - RVV1"))
RVV1 = RVV1[1:nrow(RVV1)-1,]
RVV2 = CAPs_RVVs[,7:11]
setnames(RVV2, names(RVV2), c("Account ID", "Completed - RVV2", "In progress - on track - RVV2", "In progress - not on track - RVV2", "Not started - RVV2"))
RVV2 = RVV2[complete.cases(RVV2$`Account ID`),]
RVV2 = RVV2[1:nrow(RVV2)-1,]
RVV3 = CAPs_RVVs[,13:17]
setnames(RVV3, names(RVV3), c("Account ID", "Completed - RVV3", "In progress - on track - RVV3", "In progress - not on track - RVV3", "Not started - RVV3"))
RVV3 = RVV3[complete.cases(RVV3$`Account ID`),]
RVV3 = RVV3[1:nrow(RVV3)-1,]
RVV3$`Account ID` <- as.numeric(RVV3$`Account ID`)
CCVV = CAPs_RVVs[,19:23]
setnames(CCVV, names(CCVV), c("Account ID", "Completed - CCVV", "In progress - on track - CCVV", "In progress - not on track - CCVV", "Not started - CCVV"))
CCVV = CCVV[complete.cases(CCVV$`Account ID`),]
CCVV = CCVV[1:nrow(CCVV)-1,]
CCVV$`Account ID` <- as.numeric(CCVV$`Account ID`)

# Join the CAP tables
CAPs = left_join(CAPs_pivot, RVV1, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV2, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV3, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, CCVV, by = c("Row Labels" = "Account ID"))

# Find the count of NCs
Electrical_count = CAPs %>% 
  group_by(FFC_ID, Current_Status_formatted, Sheet) %>% 
  summarise(Count = n()) %>%
  dcast(FFC_ID ~ Sheet + Current_Status_formatted + Count, fill = 0)

# Find the count of NCs by discipline
NCs_count = CAPs %>% 
  group_by(`Account ID`, `Current Status formatted`, Sheet) %>% 
  dcast(`Account ID` ~ Sheet + `Current Status formatted`, fill = 0)

# Find the count of NCs by discipline and priority level
NCs_count = CAPs %>% 
  group_by(`Account ID`, `Current Status formatted`, Sheet, Level) %>% 
  # filter(`Current Status formatted` == "*omplete" | `Current Status formatted` == "*progress*" | `Current Status formatted` == "*tarted", Level == "High" | Level == "Medium" | Level == "Low") %>%
  dcast(`Account ID` ~ Sheet + Level + `Current Status formatted`, fill = 0)

#### Fetch the Master table from the Excel spreadsheet and put the results in a dataframe ####
Master <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/MASTER FACTORY STATUS_Dec 31, 2015_AR.xlsx", "Master Factory List")
Master = Master[complete.cases(Master$`Account ID`),]

# Join the tables
Combined = left_join(Master, CAPs, by = c("Account ID" = "Row Labels"))

#### Fetch the Plan Review Tracker table from the Excel spreadsheet and put the results in a dataframe ####
PR <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Master Tracker.xls", 1, skip = 1)
PR <- PR[complete.cases(PR$`Account ID`),]
setnames(PR, c(5,8,11,14,17,20,23,26), c("DEA Status","Design Status","Central Fire Status","Hydrant Status","Sprinkler Status","Fire Door Status","Lightning Status","Single Line Diagram Status"))

# Join the tables
Combined = left_join(Combined, PR, by = "Account ID")

#### Fetch the Train the Trainer table from the Excel spreadsheet and put the results in a dataframe ####
Training <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Training Implementation_Dec 31 15_AR.xlsx", 1)
# If factory is in phase 3 and had phase 1 or 2, remove phase 1 or 2 rows
Training = arrange(Training, desc(`Training Phase`))
Training = distinct(Training, `Account ID`)
Training$`Refresher Training` <- "No"
# Training$`Training Phase`[is.na(Training$`Training Phase`)] <- "NA"
# Training$`Refresher Training`[Training$`Training Phase` == 3] <- "Yes"
Training$`Refresher Training`[!is.na(Training$`Training Phase`) & Training$`Training Phase` == 3] <- "Yes"
Training$STATUS <- gsub("Refresher Training", "Retraining", Training$STATUS)

table(Training$STATUS)

# Join the tables
Combined = left_join(Combined, Training, by = "Account ID")

##### Fetch the Security Guard Training table from the Excel spreadsheet and put the results in a dataframe ####
SG_Training <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Security Guard Training Implementation_Dec 31 15_AR.xlsx", 1)

# Join the tables
Combined = left_join(Combined, SG_Training, by = "Account ID")

#### Helpline ####
H_calls <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Amader Kotha - Dec1.14 - Dec31.15 (Alliance).xlsx", 2, skip = 1)
H_calls = H_calls[complete.cases(H_calls$`Row Labels`),]
H_factories <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/FFC/FFC Updates/Factory Profile Updates_for helpline.xlsx", 1)
H_factories = H_factories[, 1]
H_factories$Implemented <- "Yes"

# Join the tables
Helpline = left_join(H_factories, H_calls, by = c("Account ID" = "Row Labels"))
Combined = left_join(Combined, Helpline, by = "Account ID")

write.csv(Combined, "Combined.csv", na="")
