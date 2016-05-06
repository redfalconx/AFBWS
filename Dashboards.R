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
CAPs_pivot <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 1, skip = 3)
# Remove Grand Total row and change column names
CAPs_pivot = CAPs_pivot[1:nrow(CAPs_pivot)-1, ]
# Set names if not including "no level" Accord CAPs
#setnames(CAPs_pivot, c(2:16,18:32,34:48), 
#         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total",
#           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total",
#           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total"))

# Set names if including "no level" Accord CAPs !!! Double Check to make sure headers are correct !!!
setnames(CAPs_pivot, c(2:62), 
         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total", "Electrical - No Level - Completed", "Electrical - No Level - In progress - on track", "Electrical - No Level - Not started", "Electrical - No Level - Total", "Electrical Total",
           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total", "Fire - No Level - Completed", "Fire - No Level - In progress - on track", "Fire - No Level - Not started", "Fire - No Level - Total", "Fire Total",
           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total", "Structural - No Level - Completed", "Structural - No Level - In progress - on track", "Structural - No Level - Not started", "Structural - No Level - Total", "Structural Total", "Grand Total"))
# Fetch statuses of each RVV and CCVV, remove Grand Total row, and change column names
CAPs_RVVs <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 2, skip = 1)
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



#### Master Factory List ####
Master <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status_2016-April-26_AR.xlsx", "Master Factory List")
Master = Master[complete.cases(Master$`Account ID`),]

#Remove unnecessary columns !!! Double Check to make sure columns are correct !!!
Master[, c("Working Comments", "Factory Closed", "Factory Closure Reason", "Date Added to FFC (Activated as Pending)", "Deactivated brands (Date)", "Building Expanded? \r\n(if yes, list date)", "Date Approved (for factories added after April 2015)", "Thermal Scan Report Sending Date", "Linked Factories Building", "Linked Factories Compound",
           "Case Team Number", "Yet to assign team number", "QAF", "Worker Compensation Required?", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel)", "DEA / Core Test Status (Not Started, In Progress, Completed)", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel) [Second Round Review]", "Access Denied?", "1st RVV Done?", 
           "Contact information", "Contact Email", "Phone Extension", "Contacts Management", "Email of Management", "Contacts Technical staff", "Address1", "Address2", "City", "Postal Code",
           "Number of separate buildings belonging to production facility", "Number of stories of each building", "Floors of the building which the factory occupies", "Helpline Launched", "link check 15.10.4")] <- list(NULL)
Master = Master[, 1:ncol(Master)-1]

# Clean up Review Panel data 
Master$`Review Panel` <- ifelse(!is.na(Master$`Recommended to Review Panel?`), Master$`CAP Approved by Alliance`, NA)

# Clean up Number of Active Members data
Master$`Number of Active Members.2` <- ifelse(grepl("*", Master$`Number of Active Members`, fixed = TRUE) == TRUE, 
                                              "*", "")
Master$`Number of Active Members` <- gsub(" *", "", Master$`Number of Active Members`, fixed = TRUE)
Master$`Number of Active Members` = as.numeric(Master$`Number of Active Members`)
Master$`Number of Active Members` <- ifelse(grepl("*", Master$`Number of Active Members.2`, fixed = TRUE) == TRUE,
                                            Master$`Number of Active Members` + 1, Master$`Number of Active Members`)

# Join the tables
#Master = rbind(Master, Suspended, use.names = FALSE)
Combined = left_join(Master, CAPs, by = c("Account ID" = "Row Labels"))


#### Plan Review Tracker ####
PR <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Master Tracker.xls", "Master Tracker", skip = 1)
PR <- PR[complete.cases(PR$`Account ID`),]
setnames(PR, c(5,8,11,14,17,20,23,26), c("DEA Status","Design Status","Central Fire Status","Hydrant Status","Sprinkler Status","Fire Door Status","Lightning Status","Single Line Diagram Status"))
PR = PR[, c(3,5,8,11,14,17,20,23,26)]
PR[is.na(PR)] <- "Not Required or N/A"

# Join the tables
Combined = left_join(Combined, PR, by = "Account ID")


#### Basic Fire Safety Training ####
Training <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Training Implementation_Apr 26 16_AR.xlsx", 1)

# Remove unnecessary columns
Training[, c(4:32, 38:41, 48:57)] <- list(NULL)

# If factory is in phase 3 or 4 and had phase 1 or 2, remove phase 1 or 2 rows
Training = arrange(Training, desc(`Training Phase`))
Training = distinct(Training, `Account ID`)
Training$`Refresher Training` <- "No"
# Training$`Training Phase`[is.na(Training$`Training Phase`)] <- "NA"
# Training$`Refresher Training`[Training$`Training Phase` == 3] <- "Yes"
Training$`Refresher Training`[!is.na(Training$`Training Phase`) & Training$`Training Phase` == 3] <- "Yes"
Training$`Refresher Training`[!is.na(Training$`Training Phase`) & Training$`Training Phase` == 4] <- "Yes"
#Training$STATUS <- gsub("Refresher Training", "Retraining", Training$STATUS)

table(Training$STATUS)

# Join the tables
Combined = left_join(Combined, Training, by = "Account ID")

##### Security Guard Training ####
SG_Training <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Security Guard Training Implementation_Apr 26 16_AR.xlsx", 1)

# Remove unnecessary columns
SG_Training[, c(3:29, 33:36, 40:53)] <- list(NULL)

# Join the tables
Combined = left_join(Combined, SG_Training, by = "Account ID")


#### Helpline ####
## Helpline calls ##
H_calls <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", 2, skip = 1)
H_calls = H_calls[complete.cases(H_calls$`Row Labels`),]

## Helpline urgent safety calls by reason ##
H_reasons <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", 3, skip = 2)
H_reasons$`Row Labels` = as.numeric(H_reasons$`Row Labels`)
H_reasons = H_reasons[complete.cases(H_reasons$`Row Labels`),]

## Helpline factories ##
H_factories <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Factory Profile Updates_for helpline.xlsx", 1)
H_factories = H_factories[, 1]
H_factories = distinct(H_factories, `Account ID`)
H_factories$Implemented <- "Yes"

# Join the tables
Helpline = left_join(H_factories, H_calls, by = c("Account ID" = "Row Labels"))
Helpline = left_join(Helpline, H_reasons, by = c("Account ID" = "Row Labels"))
Combined = left_join(Combined, Helpline, by = "Account ID")


#### Safety Committees ####
SC <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/List of Factory for Pilot Safety Committee.xlsx", "Final List for Safety Com.Pilot")
SC = SC[complete.cases(SC$`FFC Account ID`),]
SC = SC[, 3]
SC$`Safety Committee` <- "Yes"

# Join the tables
Combined = left_join(Combined, SC, by = c("Account ID" = "FFC Account ID"))


# Save the combined data, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Combined, "Combined.csv", na="")


#### Dashboard Workbook ####
Combined <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Dashboard Workbook.xlsm", "Combined", skip = 2)

Factory_Monthly <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Dashboard Workbook.xlsm", "Factory Monthly", skip = 1)

# Add column number to column names
nm = c()
x = 1
for (i in colnames(Factory_Monthly)) {
  i = paste(i, x, sep = "_")
  nm = c(nm, i)
  x = x + 1
}
colnames(Factory_Monthly) <- nm

# Remove unnecessary columns and rows for Combined and Factory_Monthly
Combined = Combined[, 1:8]
Combined = Combined[complete.cases(Combined$`Account ID`),]
Factory_Monthly = Factory_Monthly[complete.cases(Factory_Monthly$`Account ID_1`),]

# Join old Factory Monthly to Combined
Mon = left_join(Combined, Factory_Monthly, by = c("Account ID" = "Account ID_1"))

# Save the new Factory Monthly list, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Mon, "Factory Monthly.csv", na="")


#### Tracking individual NCs over time ####
# Load raw CAP data #
CAPs_Data <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "Data", skip = 1)

# Load Master Factory List #
Master <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status_2016-April-26_AR.xlsx", "Master Factory List")

# Clean up Master, remove unnecessary columns
Master = Master[complete.cases(Master$`Account ID`),]
Master = Master[, c("Account ID", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "CCVV Date")]

# Join Master to CAPs_Data
CAPs_Data = left_join(CAPs_Data, Master, by = "Account ID")

# Save the file
write.csv(CAPs_Data, "CAP Data with RVV dates.csv", na="")