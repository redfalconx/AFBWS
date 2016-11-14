# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions


#### Tracking individual NCs over time ####
# Load raw CAP data #
CAPs_Data <- as.data.table(read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "Data", skip = 1))
CAPs_Data[, (41:ncol(CAPs_Data))] <- list(NULL)

# Remove E from Account ID # Remove in Excel (does not load Account IDs with E here)
# CAPs_Data$`Account ID` <- gsub("/E", "", CAPs_Data$`Account ID`)
# CAPs_Data$`Account ID` <- as.numeric(CAPs_Data$`Account ID`)

# Load Master Factory List #
Master <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", "Master Factory List")

# Clean up Master, remove unnecessary columns
Master = Master[complete.cases(Master$`Account ID`),]
Master = Master[, c("Account ID", "Active Brands", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "CCVV 1 Date")]
setcolorder(Master, c("Account ID", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "CCVV 1 Date", "Active Brands"))
Master$`Account ID` <- as.numeric(Master$`Account ID`)
Master = Master[complete.cases(Master$`Account ID`),]

# Clean up / consolidate variations of NCs into same NC
CAPs_Data$Question[CAPs_Data$Question == "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review? \r\n"] <- "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review?"
CAPs_Data$Question[CAPs_Data$Question == "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review?\r\n"] <- "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review?"
CAPs_Data$Question[CAPs_Data$Question == "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review? "] <- "Are as-built electrical drawings indicating information such as panel and circuit locations throughout the building(s) available for review?"

CAPs_Data$Question[CAPs_Data$Question == "Combustible materials are not stored within the substation room.\r\n"] <- "Combustible materials are not stored within the substation room."
CAPs_Data$Question[CAPs_Data$Question == "Combustible materials are not stored within the substation room. \r\n"] <- "Combustible materials are not stored within the substation room."

CAPs_Data$Question[CAPs_Data$Question == "Electrical connections at equipment, fixtures, etc are properly secured."] <- "Electrical connections at equipment, fixtures, etc. are properly secured."
CAPs_Data$Question[CAPs_Data$Question == "Electrical connections at equipment, fixtures, etc are properly secured.\r\n"] <- "Electrical connections at equipment, fixtures, etc. are properly secured."
CAPs_Data$Question[CAPs_Data$Question == "Electrical connections at equipment, fixtures, etc. are properly secured.\r\n"] <- "Electrical connections at equipment, fixtures, etc. are properly secured."

CAPs_Data$Question[CAPs_Data$Question == "Indications of overheating, overloading, or signs of burning were not observed. \r\n"] <- "Indications of overheating, overloading, or signs of burning were not observed."
CAPs_Data$Question[CAPs_Data$Question == "Indications of overheating, overloading, or signs of burning were not observed.\r\n"] <- "Indications of overheating, overloading, or signs of burning were not observed."

CAPs_Data$Question[CAPs_Data$Question == "Is a lightning protection system installed on the buildind\r\n"] <- "Is a lightning protection system installed on the building?"
CAPs_Data$Question[CAPs_Data$Question == "Is a lightning protection system installed on the building? \r\n"] <- "Is a lightning protection system installed on the building?"
CAPs_Data$Question[CAPs_Data$Question == "Is a lightning protection system installed on the building?\r\n"] <- "Is a lightning protection system installed on the building?"

CAPs_Data$Question[CAPs_Data$Question == "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)? "] <- "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?"
CAPs_Data$Question[CAPs_Data$Question == "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)? \r\n"] <- "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?"
CAPs_Data$Question[CAPs_Data$Question == "Is electrical wiring/cables sized according to CAPacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)? \r\n"] <- "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?"
CAPs_Data$Question[CAPs_Data$Question == "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?\r\n"] <- "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?"
CAPs_Data$Question[CAPs_Data$Question == "Is electrical wiring/cables sized according to d8acity of circuit breakers (No higher rated circuit breakers with lower rated wiring)? \r\n"] <- "Is electrical wiring/cables sized according to capacity of circuit breakers (No higher rated circuit breakers with lower rated wiring)?"

CAPs_Data$Question[CAPs_Data$Question == "No circuits are drawn for loads without the incorporation of a overcurrent protection device (circuit breaker).\r\n"] <- "No circuits are drawn for loads without the incorporation of a overcurrent protection device (circuit breaker)."

CAPs_Data$Question[CAPs_Data$Question == "Aisles are provided with the minimum unobstructed clear width of 0.9 m (36 in) based on occupant loads.\r\n"] <- "Aisles are provided with the minimum unobstructed clear width of 0.9 m (36 in) based on occupant loads."

CAPs_Data$Question[CAPs_Data$Question == "All doors in a means of egress are of the side-hinged swinging type.\r\n"] <- "All doors in a means of egress are of the side-hinged swinging type."

CAPs_Data$Question[CAPs_Data$Question == "Are exit enclosures provided with fire-resistive rated construction barriers?\r\n"] <- "Are exit enclosures provided with fire-resistive rated construction barriers?"

CAPs_Data$Question[CAPs_Data$Question == "Doors are not locked in the direction of egress under any conditions. All hasps, locks, slide bolts, and other locking devices have been removed where required.\r\n"] <- "Doors are not locked in the direction of egress under any conditions. All hasps, locks, slide bolts, and other locking devices have been removed where required."

CAPs_Data$Question[CAPs_Data$Question == "Illuminated exit signs are provided with battery backup or emergency power and are continuously illuminated.\r\n"] <- "Illuminated exit signs are provided with battery backup or emergency power and are continuously illuminated."

CAPs_Data$Question[CAPs_Data$Question == "Means of egress have a minimum ceiling height of 2.3 m (7 ft 6 in.) no projections from ceiling not less than 2.03 m (6 ft 8 in.)"] <- "Means of egress have a minimum ceiling height of 2.3 m (7 ft. 6 in.) with projections from the ceiling not less than 2.03 m (6 ft. 8 in.)."
CAPs_Data$Question[CAPs_Data$Question == "Means of egress have a minimum ceiling height of 2.3 m (7 ft 6 in.) with projections from the ceiling not less than 2.03 m (6 ft 8 in.)."] <- "Means of egress have a minimum ceiling height of 2.3 m (7 ft. 6 in.) with projections from the ceiling not less than 2.03 m (6 ft. 8 in.)."
CAPs_Data$Question[CAPs_Data$Question == "Means of egress should have a minimum ceiling height of 2.3 m (7 ft 6 in.) with projections from the ceiling not less than 2.03 m (6 ft 8 in.)."] <- "Means of egress have a minimum ceiling height of 2.3 m (7 ft. 6 in.) with projections from the ceiling not less than 2.03 m (6 ft. 8 in.)."

CAPs_Data$Question[CAPs_Data$Question == "Stairwells are not utilized as storage spaces"] <- "Stairwells are not utilized as storage spaces."
CAPs_Data$Question[CAPs_Data$Question == "Stairwells are not utilized as storage spaces. \r\n"] <- "Stairwells are not utilized as storage spaces."
CAPs_Data$Question[CAPs_Data$Question == "Stairwells are not utilized as storage spaces.\r\n"] <- "Stairwells are not utilized as storage spaces."

CAPs_Data$Question[CAPs_Data$Question == "The path of egress along the means of egress is not reduced at any point along the path of travel and is sufficient for the occupant load.\r\n"] <- "The path of egress along the means of egress is not reduced at any point along the path of travel and is sufficient for the occupant load."

CAPs_Data$Question[CAPs_Data$Question == "Travel distance to reach an exit does not exceed the maximum distance allowed by Occupancy Type.\r\n"] <- "Travel distance to reach an exit does not exceed the maximum distance allowed by Occupancy Type."

CAPs_Data$Question[CAPs_Data$Question == "Are the available For for the columns adequate based on Preliminary calculation?"] <- "Are the available FoS for the columns adequate based on Preliminary calculation?"
CAPs_Data$Question[CAPs_Data$Question == "Are the available FoS for the columns adequate based on Preliminary calculation?+B17:B23"] <- "Are the available FoS for the columns adequate based on Preliminary calculation?"

CAPs_Data$Question[CAPs_Data$Question == "Are the performance of key structural elements such as columns, slender\r\ncolumns, flat plates and transfer structures satisfactory?"] <- "Are the performance of key structural elements such as columns, slender columns, flat plates and transfer structures satisfactory?"

CAPs_Data$Question[CAPs_Data$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded? "] <- "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded?"
CAPs_Data$Question[CAPs_Data$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded? \r\n"] <- "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded?"
CAPs_Data$Question[CAPs_Data$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded?\r\n"] <- "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded?"

CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system? "] <- "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system?"
CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system? \r\n"] <- "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system?"
CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system?\r\n"] <- "Is the structural system free of distress, separations, or cracking that indicates lack of performance or overstress of the lateral load-carrying system?"

CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls? "] <- "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls?"
CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls? \r\n"] <- "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls?"
CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls?\r\n"] <- "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls?"
CAPs_Data$Question[CAPs_Data$Question == "Is the structural system free of distress, settlement, shifting, or cracking in\r\ncolumns or walls?"] <- "Is the structural system free of distress, settlement, shifting, or cracking in columns or walls?"

# Change priority level to High for Highest priority NCs
CAPs_Data$Level[CAPs_Data$Question == "Illuminated exit signs are provided with battery backup or emergency power and are continuously illuminated."] <- "High"
CAPs_Data$Level[CAPs_Data$Question == "Means of egress have a minimum ceiling height of 2.3 m (7 ft. 6 in.) with projections from the ceiling not less than 2.03 m (6 ft. 8 in.)."] <- "High"
CAPs_Data$Level[CAPs_Data$Question == "Are the performance of key structural elements such as columns, slender columns, flat plates and transfer structures satisfactory?"] <- "High"
CAPs_Data$Level[CAPs_Data$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded?"] <- "High"

CAPs_Data$Level = ifelse(grepl("high", CAPs_Data$Level, ignore.case = TRUE) == TRUE, "High", CAPs_Data$Level)


# Join Master to CAPs_Data
CAPs_Data = left_join(CAPs_Data, Master, by = "Account ID")

# Save the file
write.csv(CAPs_Data, "CAP Data with RVV dates.csv", na="")

### Open the csv file, paste only the Account ID, Question, and Level columns and the join from Master into Remediation Workbook ##
## Refresh all data in Remediation Workbook ##

#### Fetch the Remediation Workbook spreadsheets and put the results in dataframes ####
# CAPs <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 2, skip = 1)
CAPs_pivot <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "Current Status Pivot", skip = 3)
# Remove Grand Total row and change column names
CAPs_pivot = CAPs_pivot[1:nrow(CAPs_pivot)-1, ]

# Set names if not including "no level" Accord CAPs
#setnames(CAPs_pivot, c(2:16,18:32,34:48), 
#         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total",
#           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total",
#           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total"))

# Set names if including "no level" Accord CAPs !!! Double Check to make sure headers are correct !!!
setnames(CAPs_pivot, c(2:63), 
         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High - Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total", "Electrical - No Level - Completed", "Electrical - No Level - In progress - on track", "Electrical - No Level - Not started", "Electrical - No Level - Total", "Electrical Total",
           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High - Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total", "Fire - No Level - Completed", "Fire - No Level - In progress - not on track", "Fire - No Level - In progress - on track", "Fire - No Level - Not started", "Fire - No Level - Total", "Fire Total",
           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High - Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total", "Structural - No Level - Completed", "Structural - No Level - In progress - on track", "Structural - No Level - Not started", "Structural - No Level - Total", "Structural Total", "Grand Total"))

# Fetch statuses of each RVV and CCVV, remove Grand Total row, and change column names
CAPs_RVVs <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "RVV Pivots", skip = 2)
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
RVV4 = CAPs_RVVs[,19:23]
setnames(RVV4, names(RVV4), c("Account ID", "Completed - RVV4", "In progress - on track - RVV4", "In progress - not on track - RVV4", "Not started - RVV4"))
RVV4 = RVV4[complete.cases(RVV4$`Account ID`),]
RVV4 = RVV4[1:nrow(RVV4)-1,]
RVV4$`Account ID` <- as.numeric(RVV4$`Account ID`)

# Join the CAP tables
CAPs = left_join(CAPs_pivot, RVV1, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV2, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV3, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV4, by = c("Row Labels" = "Account ID"))

# Fetch Highest Priority (Urgent Life Safety) NCs
HPNCs_all <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "All Factories", skip = 3)
HPNCs_com <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", "Completed", skip = 8)

# Remove unnecessary columns
HPNCs_all = HPNCs_all[complete.cases(HPNCs_all$`Row Labels`),]
HPNCs_all = HPNCs_all[, c("Row Labels", "Grand Total")]
HPNCs_com = HPNCs_com[complete.cases(HPNCs_com$`Row Labels`),]
HPNCs_com = HPNCs_com[, c("Row Labels", "Grand Total")]

# Combined datasets, then calculate % completed
HPNCs = left_join(HPNCs_all, HPNCs_com, by = "Row Labels")
setnames(HPNCs, "Grand Total.y", "# of Highest Priority NCs completed")
setnames(HPNCs, "Grand Total.x", "# of Highest Priority NCs")
HPNCs = HPNCs %>% mutate("% of Highest Priority NCs completed" = `# of Highest Priority NCs completed`/`# of Highest Priority NCs`)
HPNCs$`% of Highest Priority NCs completed` = ifelse(is.na(HPNCs$`% of Highest Priority NCs completed`) == TRUE, 0, HPNCs$`% of Highest Priority NCs completed`)

# Join to CAPs
CAPs = left_join(CAPs, HPNCs, by = "Row Labels")


#### Master Factory List ####
Master <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", "Master Factory List")
Master$`Account ID` = as.numeric(Master$`Account ID`)
Master = Master[complete.cases(Master$`Account ID`),]
Master$`Recommended to Review Panel?` <- as.character(Master$`Recommended to Review Panel?`)
Master = Master[, 1:91]
Suspended <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", "Suspended Factories")
Suspended = Suspended[complete.cases(Suspended$`Account ID`),]
Suspended = Suspended[, c("Account Name", "Account ID", "Active Brands", "Shared with Accord", "Inspected by Alliance / Accord / All&Acc", "Recommended to Review Panel?", "CAP Approved by Alliance", "Escalation Status", "Remediation Factory Status")]
Suspended$`Escalation Status` <- "Suspended Approval Notification"
Suspended$`Remediation Factory Status` <- "Critical"
#Transferred <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", "Moved to Accord")
#Transferred = Transferred[, c("Account Name", "Account ID", "Active Brands", "Recommended to Review Panel?", "CAP Approved by Alliance", "Remediation Factory Status")]
#Transferred$`Remediation Factory Status` <- "Transferred to Accord"

Master = full_join(Master, Suspended)

# Remove unnecessary columns !!! Double Check to make sure columns are correct !!!
Master[, c("Working Comments", "Factory Closed", "Factory Closure Reason", "Date Added to FFC (Activated as Pending)", "Deactivated brands (Date)", "Building Expanded? \r\n(if yes, list date)", "Date Approved (for factories added after April 2015)", "Thermal Scan Report Sending Date", "Linked Factories Building", "Linked Factories Compound",
           "Case Team Number", "Yet to assign team number", "QAF", "Worker Compensation Required?", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel)", "DEA / Core Test Status (Not Started, In Progress, Completed)", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel) [Second Round Review]", "Access Denied?", "1st RVV Done?", 
           "Contact information", "Contact Email", "Phone Extension", "Contacts Management", "Email of Management", "Contacts Technical staff", "Address1", "Address2", "City", "Postal Code",
           "Number of separate buildings belonging to production facility", "Number of stories of each building", "Floors of the building which the factory occupies", "Helpline Launched", "link check 15.10.4")] <- list(NULL)
#Master = Master[, 1:ncol(Master)-1]

# Clean up Review Panel data 
Master$`Review Panel` <- ifelse(!is.na(Master$`Recommended to Review Panel?`), Master$`CAP Approved by Alliance`, NA)

# Clean up Number of Active Members data
Master$`Number of Active Members.2` <- ifelse(grepl("*", Master$`Number of Active Members`, fixed = TRUE) == TRUE, 
                                              "*", "")
Master$`Number of Active Members` <- gsub(" *", "", Master$`Number of Active Members`, fixed = TRUE)
Master$`Number of Active Members` = as.numeric(Master$`Number of Active Members`)
Master$`Number of Active Members` <- ifelse(grepl("*", Master$`Number of Active Members.2`, fixed = TRUE) == TRUE,
                                            Master$`Number of Active Members` + 1, Master$`Number of Active Members`)

# Clean up Shared with Accord data
# Master$`Shared with Accord` <- ifelse(Master$`Shared with Accord` == "Yes", "Yes", "No")
Master$`Shared with Accord`[is.na(Master$`Shared with Accord`)] <- "No"

# Join the tables
#Master = rbind(Master, Suspended, use.names = FALSE)
Combined = left_join(Master, CAPs, by = c("Account ID" = "Row Labels"))


#### Plan Review Tracker ####
PR <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Master Tracker - Drawing Design.xls", "Master Tracker", skip = 1)
PR$`Account ID` <- as.numeric(PR$`Account ID`)
PR <- PR[complete.cases(PR$`Account ID`),]
# PR$`Account ID` <- gsub("/E", "", PR$`Account ID`)
setnames(PR, c(5,8,11,14,17,20,23,26), c("DEA Status","Design Status","Central Fire Status","Hydrant Status","Sprinkler Status","Fire Door Status","Lightning Status","Single Line Diagram Status"))
PR = PR[, c(3,5,8,11,14,17,20,23,26)]
PR[is.na(PR)] <- "Not Required or N/A"

# Join the tables
Combined = left_join(Combined, PR, by = "Account ID")


#### DEA Tracker ####
DEA <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/DEA Tracker.xls", "DEA")
DEA$`Account ID` <- as.numeric(DEA$`Account ID`)
DEA <- DEA[complete.cases(DEA$`Account ID`),]
DEA = DEA[ , c("Account ID", "Retrofitting Status")]

# Join the tables
Combined = left_join(Combined, DEA, by = "Account ID")


#### Basic Fire Safety Training ####
Training <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Basic Fire Safety & Helpline Training Implementation.xlsx", 1)

# Remove unnecessary columns
Training[, c(4:24, 26:31, 36:41, 47:57)] <- list(NULL)

# If factory is in phase 3 or 4 and had phase 1 or 2, remove phase 1 or 2 rows
Training$`Training Phase`[Training$`Training Phase` == "3a"] <- "3"
Training$`Training Phase`[Training$`Training Phase` == "4a"] <- "4"
Training$`Training Phase` <- as.numeric(Training$`Training Phase`)
Training = arrange(Training, desc(`Training Phase`))
Training = distinct(Training, `Account ID`, .keep_all = TRUE)
Training$`Refresher Training` <- "No"
Training$`Refresher Training`[!is.na(Training$`Training Phase`) & Training$`Training Phase` == 3] <- "Yes"
Training$`Refresher Training`[!is.na(Training$`Training Phase`) & Training$`Training Phase` == 4] <- "Yes"

table(Training$STATUS)

# Join the tables
Combined = left_join(Combined, Training, by = "Account ID")

#Combined$`Refresher Training` = ifelse(Combined$`Refresher Training` != "Yes", "No")


##### Security Guard Training ####
SG_Training <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Security Guard Training Implementation.xlsx", 1)

# Remove unnecessary columns
SG_Training[, c(3:28, 32:35, 43:52)] <- list(NULL)

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
H_factories <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Factory Profile Update_Helpline.xlsx", 1)
H_factories = H_factories[, c("Account ID", "Workers Trained")]
H_factories = H_factories[complete.cases(H_factories$`Account ID`), ]
H_factories$Implemented <- "Yes"

# Join the tables
Helpline = left_join(H_factories, H_calls, by = c("Account ID" = "Row Labels"))
Helpline = left_join(Helpline, H_reasons, by = c("Account ID" = "Row Labels"))
Combined = left_join(Combined, Helpline, by = "Account ID")

# Add "No" to Helpline Implemented column
Combined$Implemented[is.na(Combined$Implemented)] <- "No"

#### Safety Committees ####
SC <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/SC Implementation.xlsx", 1)

SC = SC[complete.cases(SC$`Account ID`),]
setnames(SC, names(SC), gsub("\\r\\n", " ", names(SC)))
# SC = SC[SC$`SC Formation  (Yes/No) %` > 0,]
# SC = SC[!is.na(SC$`SC Formation  (Yes/No) %`),]
SC = SC[, c("Account ID", "PC / CBA or TU / WWA (Yes/No) %" , "SC Formation  (Yes/No) %", "SC Formation Date", "SC Formation Process", "Total SC members", "TtT Received from Alliance (Yes/No) %", "Number of Participants", "Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators", "SC Activity Implementation Completion Date (Total - 100 Days)", "Action Plan Submitted (Yes/No) %", "Introduction of SC (Yes/No) (%)", "Conducted Training for rest of SC members by  Factory Facilitators (Yes/No) %", "Safety Policy &  Emergency Response Procedure  (Yes/No) %", "Risk Assessment by Safety Committee & Follow Up (Yes/No) % Total", "Formal SC Meeting Arrangement (Yes/No) % Total", "Conduct Fire Or Evacuation Drill (Yes/No) % Total", "Final Percentage (Achievement) %", "Status")]

# setnames(SC, "TtT Received from Alliance (Yes/No) %", "TtT Received from Alliance")
setnames(SC, "SC Activity Implementation Completion Date (Total - 100 Days)", "SC Activity Implementation Completion Date")
SC$Status <- gsub("^([a-z])", "\\U\\1", tolower(SC$Status), perl=TRUE)

# Change all to characters
SC[] <- lapply(SC, as.character)
# Change these back to numeric
SC$`Total SC members` <- as.numeric(SC$`Total SC members`)
SC$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators` <- as.numeric(SC$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators`)
SC$`Account ID` <- as.numeric(SC$`Account ID`)

# Change time values to completed, in progress, or NA
SC[SC == "1900-01-09"] <- "10"
SC[SC == "10"] <- "Completed"
SC[SC == "1900-01-01"] <- "In progress"
SC[SC == "1900-01-02"] <- "In progress"
SC[SC == "1900-01-03"] <- "In progress"
SC[SC == "1900-01-04"] <- "In progress"
SC[SC == "1900-01-05"] <- "In progress"
SC[SC == "1900-01-06"] <- "In progress"
SC[SC == "1900-01-07"] <- "In progress"
SC[SC == "1900-01-08"] <- "In progress"
SC[SC == "1899-12-31"] <- "In progress"
SC[SC == "3.33"] <- "In progress"
SC[SC == "6.66"] <- "In progress"
SC[SC == "1899-12-30"] <- NA


# SC$`TtT Received from Alliance (Yes/No) %` = ifelse(is.na(SC$`TtT Received from Alliance (Yes/No) %`) == TRUE, "No", "Yes")

# Join the tables
Combined = left_join(Combined, SC, by = "Account ID")

# Remove duplicate rows
Combined = Combined[!duplicated(Combined[,"Account ID"]),]


#### Expansions ####
# Load factories that have had expansion audits from FFC Expansion Audits.R script
source('~/R/AFBWS/FFC Expansion Audits.R', echo=TRUE)
Expansions$`Account ID` <- as.numeric(as.character(Expansions$`Account ID`))

# Remove duplicate rows
Expansions = Expansions[!duplicated(Expansions[,"Account ID"]),]

# Join the tables
Combined = left_join(Combined, Expansions, by = "Account ID")

# Add column for Yes
Combined$Expansion = ifelse(is.na(Combined$`Audit Scope`) == TRUE, "No", "Yes")

# Remove duplicate rows
Combined = Combined[!duplicated(Combined[,"Account ID"]),]


# Save the combined data, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Combined, "Combined.csv", na="")


#### Lockable Gates ####
LG <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Statement of Lockable Exit.xls", "Statement Lockable exits_CM1", skip = 2)
LG = LG[complete.cases(LG$`Factory ID`),]
LG = LG[, c(4, 9)]

# Join the tables
Combined = left_join(Combined, LG, by = c("Account ID" = "Factory ID"))

# Add "Not visited yet" to factories not in lockable gates tracker
Combined$`No. of lockable Exits`[is.na(Combined$`No. of lockable Exits`)] <- "Not visited yet"

# Remove Duplicate rows
Combined = Combined[!duplicated(Combined[,"Account ID"]),]


#### Reorder the columns of Combined to match the Dashboard Workbook more closely ####
someCol <- c("Account ID", "Account Name.x", "Active Brands", "Number of Active Members", "Remediation Factory Status", "Shared with Accord", "Inspected by Alliance / Accord / All&Acc", "Expansion", "Number of employees to be trained.", "Number of workers employed by factory (all buildings)", "Province", "RENTED ", "Mixed Occupancy", "Factory housing in multi-factory building", "Case Group", 
             "Escalation Date", "Escalation Status", "Review Panel", 
             "Electrical High - Completed", "Electrical High - In progress - on track", "Electrical High - In progress - not on track", "Electrical High - Not Started", "Electrical High Total",
             "Electrical Medium - Completed", "Electrical Medium - In progress - on track", "Electrical Medium - In progress - not on track", "Electrical Medium - Not Started", "Electrical Medium Total",
             "Electrical Low - Completed", "Electrical Low - In progress - on track", "Electrical Low - In progress - not on track", "Electrical Low - Not Started", "Electrical Low Total",
             "Electrical - No Level - Completed", "Electrical - No Level - In progress - on track", "Electrical - No Level - Not started", "Electrical - No Level - Total",
             "Fire High - Completed", "Fire High - In progress - on track", "Fire High - In progress - not on track", "Fire High - Not Started", "Fire High Total",
             "Fire Medium - Completed", "Fire Medium - In progress - on track", "Fire Medium - In progress - not on track", "Fire Medium - Not Started", "Fire Medium Total",
             "Fire Low - Completed", "Fire Low - In progress - on track", "Fire Low - In progress - not on track", "Fire Low - Not Started", "Fire Low Total",
             "Fire - No Level - Completed", "Fire - No Level - In progress - on track", "Fire - No Level - In progress - not on track", "Fire - No Level - Not started", "Fire - No Level - Total",
             "Structural High - Completed", "Structural High - In progress - on track", "Structural High - In progress - not on track", "Structural High - Not Started", "Structural High Total",
             "Structural Medium - Completed", "Structural Medium - In progress - on track", "Structural Medium - In progress - not on track", "Structural Medium - Not Started", "Structural Medium Total",
             "Structural Low - Completed", "Structural Low - In progress - on track", "Structural Low - In progress - not on track", "Structural Low - Not Started", "Structural Low Total",
             "Structural - No Level - Completed", "Structural - No Level - In progress - on track", "Structural - No Level - Not started", "Structural - No Level - Total",
             "# of Highest Priority NCs", "# of Highest Priority NCs completed", "% of Highest Priority NCs completed", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Completed - RVV1", "In progress - on track - RVV1", "In progress - not on track - RVV1", "Not started - RVV1",
             "Confirmed Date of 2nd RVV", "Completed - RVV2", "In progress - on track - RVV2", "In progress - not on track - RVV2", "Not started - RVV2",
             "Confirmed Date of 3rd RVV", "Completed - RVV3", "In progress - on track - RVV3", "In progress - not on track - RVV3", "Not started - RVV3",
             "Confirmed Date of 4th RVV", "Completed - RVV4", "In progress - on track - RVV4", "In progress - not on track - RVV4", "Not started - RVV4",
             "CCVV 1 Date", "CCVV 1 % of Completion", "CCVV 2 Date", "CCVV 2 % of Completion", "CCVV 1 Result", 
             "Retrofitting Status", "DEA Status", "Design Status", "Central Fire Status", "Hydrant Status", "Sprinkler Status", "Fire Door Status", "Lightning Status", "Single Line Diagram Status",
             "Refresher Training", "Total number of employees trained so far.", "Total number of employees trained (Duplicates Removed)", "Percentage of Workers Trained", "STATUS", "Spot Check Results \r\n(Pass or Fail)", "Support Visit Required?",
             " Total number of security staff trained so far", "Percentage of security staff trained", "STATUS ", "Spot Check Results (Pass or Fail)", "Support Visit",
             "PC / CBA or TU / WWA (Yes/No) %", "SC Formation  (Yes/No) %", "SC Formation Date", "SC Formation Process", "TtT Received from Alliance (Yes/No) %", "Number of Participants", "Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators", "SC Activity Implementation Completion Date", "Status",
             "Implemented", "Workers Trained", "General Inquiries", "No Category", "Non-urgent: Non-safety", "Non-urgent: Safety", "Urgent: Non-safety", "Urgent: Safety",
             "Fire - Active (factory)", "Fire - Danger (factory)", "Locked factory exit or blocked egress route", "Other", "Sparking / short circuit", "Structural - Cracks in beams, columns or walls", "Structural - walls or windows shaking", "Unattended / bare electric wires", "Unauthorized subcontracting", "Undisclosed")

setcolorder(Combined, c(someCol, colnames(Combined)[!(colnames(Combined) %in% someCol)]))



#### Dashboard Workbook ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions

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
Combined = Combined[, 1:10]
Combined = Combined[complete.cases(Combined$`Account ID`),]
Factory_Monthly = Factory_Monthly[complete.cases(Factory_Monthly$`Account ID_1`),]

# Join old Factory Monthly to Combined
Mon = left_join(Combined, Factory_Monthly, by = c("Account ID" = "Account ID_1"))

# Save the new Factory Monthly list, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Mon, "Factory Monthly.csv", na="")



