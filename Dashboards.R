# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions

wd = dirname(getwd())
wd = sub("/OneDrive", "", wd)

#### Tracking individual NCs over time ####
# Load raw CAP data #
CAPs_Data <- as.data.table(read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsm", sep = ""), "Data", skip = 1, col_types = "text"))
CAPs_Data[, (72:ncol(CAPs_Data))] <- list(NULL)

CAPs_Data$`Account ID` <- gsub("/E", "", CAPs_Data$`Account ID`)
CAPs_Data$`Account ID` <- as.numeric(CAPs_Data$`Account ID`)

# Load Master Factory List #
Master <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", sep = ""), "Master Factory List")

# Clean up Master, remove unnecessary columns
Master = Master[complete.cases(Master$`Account ID`),]
Master = Master[, c("Account ID", "Active Brands", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "Confirmed Date of 5th RVV", "Confirmed Date of 6th RVV", "Confirmed Date of 7th RVV", "Confirmed Date of 8th RVV", "CCVV 1 Date", "CCVV 2 Date", "CCVV 3 Date")]
setcolorder(Master, c("Account ID", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "Confirmed Date of 5th RVV", "Confirmed Date of 6th RVV", "Confirmed Date of 7th RVV", "Confirmed Date of 8th RVV", "CCVV 1 Date", "CCVV 2 Date", "CCVV 3 Date", "Active Brands"))
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
CAPs = left_join(CAPs_Data, Master, by = "Account ID")

# Save the file
write.csv(CAPs, "CAP Data with RVV dates.csv", na="")

## Open the csv file, paste only the Question and Level columns and the join from Master into Remediation Workbook ##
## Refresh "Current Status - no new NCs" column ##
## Refresh all data in Remediation Workbook and save ##

#### Fetch the Remediation Workbook spreadsheets and put the results in dataframes ####
# CAPs <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/Member Dashboards/Dashboard Workbook/Remediation Workbook.xlsx", 2, skip = 1)
CAPs_pivot <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsm", sep = ""), "Current Status Pivot", skip = 3)
# Remove Grand Total row and change column names
CAPs_pivot = CAPs_pivot[1:nrow(CAPs_pivot)-1, ]

# Set names if not including "no level" Accord CAPs
#setnames(CAPs_pivot, c(2:16,18:32,34:48), 
#         c("Electrical High - Completed",	"Electrical High - In progress - on track",	"Electrical High - In progress - not on track",	"Electrical High Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - on track",	"Electrical Medium - In progress - not on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - on track",	"Electrical Low - In progress - not on track",	"Electrical Low - Not Started",	"Electrical Low Total",
#           "Fire High - Completed",	"Fire High - In progress - on track",	"Fire High - In progress - not on track",	"Fire High Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - on track",	"Fire Medium - In progress - not on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - on track",	"Fire Low - In progress - not on track",	"Fire Low - Not Started",	"Fire Low Total",
#           "Structural High - Completed",	"Structural High - In progress - on track",	"Structural High - In progress - not on track",	"Structural High Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - on track",	"Structural Medium - In progress - not on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - on track",	"Structural Low - In progress - not on track",	"Structural Low - Not Started",	"Structural Low Total"))

# Set names if including "no level" Accord CAPs !!! Double Check to make sure headers are correct !!!
# Old one with "on track" and "not on track"
#setnames(CAPs_pivot, c(2:65), 
#         c("Electrical High - Completed",	"Electrical High - In progress - not on track",	"Electrical High - In progress - on track",	"Electrical High - Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress - not on track",	"Electrical Medium - In progress - on track",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress - not on track",	"Electrical Low - In progress - on track",	"Electrical Low - Not Started",	"Electrical Low Total", "Electrical - No Level - Completed", "Electrical - No Level - In progress - not on track", "Electrical - No Level - In progress - on track", "Electrical - No Level - Not started", "Electrical - No Level - Total", "Electrical Total",
#           "Fire High - Completed",	"Fire High - In progress - not on track",	"Fire High - In progress - on track",	"Fire High - Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress - not on track",	"Fire Medium - In progress - on track",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress - not on track",	"Fire Low - In progress - on track",	"Fire Low - Not Started",	"Fire Low Total", "Fire - No Level - Completed", "Fire - No Level - In progress - not on track", "Fire - No Level - In progress - on track", "Fire - No Level - Not started", "Fire - No Level - Total", "Fire Total",
#           "Structural High - Completed",	"Structural High - In progress - not on track",	"Structural High - In progress - on track",	"Structural High - Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress - not on track",	"Structural Medium - In progress - on track",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress - not on track",	"Structural Low - In progress - on track",	"Structural Low - Not Started",	"Structural Low Total", "Structural - No Level - Completed", "Structural - No Level - In progress - not on track", "Structural - No Level - In progress - on track", "Structural - No Level - Not started", "Structural - No Level - Total", "Structural Total", "Grand Total"))

# Set names if including "no level" Accord CAPs !!! Double Check to make sure headers are correct !!!
setnames(CAPs_pivot, c(2:53), 
         c("Electrical High - Completed",	"Electrical High - In progress",	"Electrical High - Not Started",	"Electrical High Total",	"Electrical Medium - Completed",	"Electrical Medium - In progress",	"Electrical Medium - Not Started",	"Electrical Medium Total",	"Electrical Low - Completed",	"Electrical Low - In progress",	"Electrical Low - Not Started",	"Electrical Low Total", "Electrical - No Level - Completed", "Electrical - No Level - In progress", "Electrical - No Level - Not started", "Electrical - No Level - Total", "Electrical Total",
           "Fire High - Completed",	"Fire High - In progress",	"Fire High - Not Started",	"Fire High Total",	"Fire Medium - Completed",	"Fire Medium - In progress",	"Fire Medium - Not Started",	"Fire Medium Total",	"Fire Low - Completed",	"Fire Low - In progress",	"Fire Low - Not Started",	"Fire Low Total", "Fire - No Level - Completed", "Fire - No Level - In progress", "Fire - No Level - Not started", "Fire - No Level - Total", "Fire Total",
           "Structural High - Completed",	"Structural High - In progress",	"Structural High - Not Started",	"Structural High Total",	"Structural Medium - Completed",	"Structural Medium - In progress",	"Structural Medium - Not Started",	"Structural Medium Total",	"Structural Low - Completed",	"Structural Low - In progress",	"Structural Low - Not Started",	"Structural Low Total", "Structural - No Level - Completed", "Structural - No Level - In progress", "Structural - No Level - Not started", "Structural - No Level - Total", "Structural Total", "Grand Total"))


# Fetch statuses of each RVV and CCVV, remove Grand Total row, and change column names
CAPs_RVVs <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsm", sep = ""), "RVV Pivots", skip = 3)
RVV1 = CAPs_RVVs[,1:4]
setnames(RVV1, names(RVV1), c("Account ID", "Completed - RVV1", "In progress - RVV1", "Not started - RVV1"))
RVV1 = RVV1[1:nrow(RVV1)-1,]
RVV2 = CAPs_RVVs[,6:9]
setnames(RVV2, names(RVV2), c("Account ID", "Completed - RVV2", "In progress - RVV2", "Not started - RVV2"))
RVV2 = RVV2[complete.cases(RVV2$`Account ID`),]
RVV2 = RVV2[1:nrow(RVV2)-1,]
RVV3 = CAPs_RVVs[,11:14]
setnames(RVV3, names(RVV3), c("Account ID", "Completed - RVV3", "In progress - RVV3", "Not started - RVV3"))
RVV3 = RVV3[complete.cases(RVV3$`Account ID`),]
RVV3 = RVV3[1:nrow(RVV3)-1,]
RVV4 = CAPs_RVVs[,16:19]
setnames(RVV4, names(RVV4), c("Account ID", "Completed - RVV4", "In progress - RVV4", "Not started - RVV4"))
RVV4 = RVV4[complete.cases(RVV4$`Account ID`),]
RVV4 = RVV4[1:nrow(RVV4)-1,]
#RVV4$`Account ID` <- as.numeric(RVV4$`Account ID`)
RVV5 = CAPs_RVVs[,21:24]
setnames(RVV5, names(RVV5), c("Account ID", "Completed - RVV5", "In progress - RVV5", "Not started - RVV5"))
RVV5 = RVV5[complete.cases(RVV5$`Account ID`),]
RVV5 = RVV5[1:nrow(RVV5)-1,]
#RVV5$`Account ID` <- as.numeric(RVV5$`Account ID`)
RVV6 = CAPs_RVVs[,26:29]
setnames(RVV6, names(RVV6), c("Account ID", "Completed - RVV6", "In progress - RVV6", "Not started - RVV6"))
RVV6 = RVV6[complete.cases(RVV6$`Account ID`),]
RVV6 = RVV6[1:nrow(RVV6)-1,]
#RVV6$`Account ID` <- as.numeric(RVV6$`Account ID`)
RVV7 = CAPs_RVVs[,31:34]
setnames(RVV7, names(RVV7), c("Account ID", "Completed - RVV7", "In progress - RVV7", "Not started - RVV7"))
RVV7 = RVV7[complete.cases(RVV7$`Account ID`),]
RVV7 = RVV7[1:nrow(RVV7)-1,]
#RVV7$`Account ID` <- as.numeric(RVV7$`Account ID`)
RVV8 = CAPs_RVVs[,36:39]
setnames(RVV8, names(RVV8), c("Account ID", "Completed - RVV8", "In progress - RVV8", "Not started - RVV8"))
RVV8 = RVV8[complete.cases(RVV8$`Account ID`),]
RVV8 = RVV8[1:nrow(RVV8)-1,]
#RVV7$`Account ID` <- as.numeric(RVV7$`Account ID`)

# Join the CAP tables
CAPs = left_join(CAPs_pivot, RVV1, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV2, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV3, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV4, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV5, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV6, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV7, by = c("Row Labels" = "Account ID"))
CAPs = left_join(CAPs, RVV8, by = c("Row Labels" = "Account ID"))

# Fetch Highest Priority (Urgent Life Safety) NCs
HPNCs_all <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsm", sep = ""), "All Factories", skip = 4)
HPNCs_com <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Remediation Workbook.xlsm", sep = ""), "Completed", skip = 11)

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
Master <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/MASTER Factory Status.xlsx", sep = ""), "Master Factory List")
Master = Master[complete.cases(Master$`Account ID`),]
Master$`Recommended to Review Panel?` <- as.character(Master$`Recommended to Review Panel?`)
#Master = Master[, 1:82]

# Remove unnecessary columns !!! Double Check to make sure columns are correct !!!
#Master[, c("Working Comments", "Factory Closed", "Factory Closure Reason", "Date Added to FFC (Activated as Pending)", "Deactivated brands (Date)", "Building Expanded? \r\n(if yes, list date)", "Date Approved (for factories added after April 2015)", "Thermal Scan Report Sending Date", "Linked Factories Building", "Linked Factories Compound",
#           "Case Team Number", "Yet to assign team number", "QAF", "Worker Compensation Required?", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel)", "DEA / Core Test Status (Not Started, In Progress, Completed)", "FOS Status (OK, DEA Needed, Core Test Needed, Review Panel) [Second Round Review]", "Access Denied?", "1st RVV Done?", 
#           "Contact information", "Contact Email", "Phone Extension", "Contacts Management", "Email of Management", "Contacts Technical staff", "Address1", "Address2", "City", "Postal Code",
#           "Number of separate buildings belonging to production facility", "Number of stories of each building", "Floors of the building which the factory occupies", "Helpline Launched", "link check 15.10.4")] <- list(NULL)

Master = Master[, c("Account ID", "Account Name", "Active Brands", "Number of Active Members", "Remediation Factory Status", "Accord Shared/Alliance only info", "Inspected by Alliance / Accord / All&Acc", "Province", "RENTED", "Mixed Occupancy", "Factory housing in multi-factory building", "Case Group", "Escalation Date", "Escalation Status", 
                    "Recommended to Review Panel?", "CAP Approved by Alliance", "Date of Initial Inspection", "CAP Approval Date", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV" , "Confirmed Date of 4th RVV", "Confirmed Date of 5th RVV", "Confirmed Date of 6th RVV", "Confirmed Date of 7th RVV", "Confirmed Date of 8th RVV",
                    "CCVV 1 Date", "CCVV 1 Result", "CCVV 1 % of Completion", "CCVV 2 Date", "CCVV 2 Result", "CCVV 2 % of Completion", "CCVV 3 Date", "CCVV 3 Result", "CCVV 3 % of Completion", "Safety Management Visit Date", "Safety Management Visit Status", "Box Folder Link")]

# Change Subs. Completion to Initial CAP Completed
table(Master$`Remediation Factory Status`)
Master$`Remediation Factory Status` = ifelse(grepl("Subs", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE, "Initial CAP Completed", Master$`Remediation Factory Status`)

# Change from "CCVV Pending Awaiting CAP Closure of Unoccupied Expansion" to "On track"
Master$`Remediation Factory Status`[grepl("CCVV Pending", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE] = "On track"

# Master$`Remediation Factory Status` = ifelse(grepl("Suspended", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE, "Critical", Master$`Remediation Factory Status`)

# Clean up Review Panel data 
Master$`Review Panel` <- ifelse(!is.na(Master$`Recommended to Review Panel?`), Master$`CAP Approved by Alliance`, NA)

# Separate out Expansions to get Remediation Status
Master_Exp = Master[grepl("/E", Master$`Account ID`, ignore.case = TRUE) == TRUE, ]
Master$`Account ID` = as.numeric(Master$`Account ID`)
Master = Master[complete.cases(Master$`Account ID`),]

Master_Exp = Master_Exp[, c("Account ID", "Remediation Factory Status")]
setnames(Master_Exp, "Remediation Factory Status", "Expansion Remediation Status")

Master_Exp$`Expansion Remediation Status` = ifelse(grepl("building", Master_Exp$`Expansion Remediation Status`, ignore.case = TRUE) == TRUE, NA, Master_Exp$`Expansion Remediation Status`)

#Master_Exp$`Account ID` = arrange(Master_Exp, desc(`Account ID`))
Master_Exp$`Account ID` = gsub("/E.*", "", Master_Exp$`Account ID`)
Master_Exp$`Account ID` = gsub("/E", "", Master_Exp$`Account ID`)

Master_Exp = aggregate(Master_Exp["Expansion Remediation Status"], Master_Exp["Account ID"], paste, collapse = ", ")

Master_Exp$`Account ID` = as.numeric(Master_Exp$`Account ID`)

# Clean up Number of Active Members data
Master$`Number of Active Members.2` <- ifelse(grepl("*", Master$`Number of Active Members`, fixed = TRUE) == TRUE, 
                                              "*", "")
Master$`Number of Active Members` <- gsub(" *", "", Master$`Number of Active Members`, fixed = TRUE)
Master$`Number of Active Members` = as.numeric(Master$`Number of Active Members`)
Master$`Number of Active Members` <- ifelse(grepl("*", Master$`Number of Active Members.2`, fixed = TRUE) == TRUE,
                                            Master$`Number of Active Members` + 1, Master$`Number of Active Members`)

# Clean up Shared with Accord data
# Master$`Shared with Accord` <- ifelse(Master$`Shared with Accord` == "Yes", "Yes", "No")
Master$`Accord Shared/Alliance only info`[is.na(Master$`Accord Shared/Alliance only info`)] <- "Alliance Only"

# Join the tables
Master = left_join(Master, Master_Exp, by = "Account ID")
CAPs$`Row Labels` = as.numeric(CAPs$`Row Labels`)
Combined = left_join(Master, CAPs, by = c("Account ID" = "Row Labels"))


#### Plan Review Tracker ####
PR <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Master Tracker - Drawing Design_Without Suspended & Under Accord Factories.xlsx", sep = ""), "Master Tracker", skip = 1)
PR$`Account ID` <- as.numeric(PR$`Account ID`)
PR <- PR[complete.cases(PR$`Account ID`),]
# PR$`Account ID` <- gsub("/E", "", PR$`Account ID`)
setnames(PR, c("Status","Status__1","Status__2","Status__3","Status__4","Status__5","Status__6","Status__7"), c("DEA Status","Design Status","Central Fire Status","Hydrant Status","Sprinkler Status","Fire Door Status","Lightning Status","Single Line Diagram Status"))
PR = PR[, c("Account ID", "DEA Status","Design Status","Central Fire Status","Hydrant Status","Sprinkler Status","Fire Door Status","Lightning Status","Single Line Diagram Status")]
PR[is.na(PR)] <- "Not Required or N/A"

# Join the tables
Combined = left_join(Combined, PR, by = "Account ID")


#### DEA Tracker ####
DEA <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/DEA Tracker.xls", sep = ""), "DEA")
DEA$`Account ID` <- as.numeric(DEA$`Account ID`)
DEA <- DEA[complete.cases(DEA$`Account ID`),]
DEA = DEA[ , c("Account ID", "Retrofitting Status")]

# Join the tables
Combined = left_join(Combined, DEA, by = "Account ID")

# Change to Completed if Remediation has been completed
Combined$`Retrofitting Status`[grepl("completed", Combined$`Remediation Factory Status`, ignore.case = TRUE) == TRUE] = "Completed"

#### Basic Fire Safety Training ####
## Before loading, put in descending order with 4a and 3a first ##
Training <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Basic Fire Safety & Helpline Training Implementation.xlsx", sep = ""), "BFST Factory List", col_types = "text")
Training$`Account ID` = as.numeric(Training$`Account ID`)
Training = Training[complete.cases(Training$`Account ID`), ]
# Check to make sure 3a and 4a phases are included
table(Training$`Training Phase`, useNA = "ifany")

# Remove unnecessary columns
# Training[, c(4:25, 27:32, 37:43, 63:72)] <- list(NULL)
Training = Training[, c("Account ID", "Number of employees to be trained.", "Total number of employees trained so far.", "Training Phase", "Percentage of Workers Trained", "STATUS", "Final Training Status \r\n(CCVV)",
                        "Final Training Assessment (CCVV) \r\nCurrent Results \r\n(Pass or Fail)", "Support Visit Required?")]

# Subset initial training (phase 1, 2, 3a, 4a)
IT = Training[Training$`Training Phase` %in% c("1","2", "3a", "4a"),]
IT = IT[, c("Account ID", "Total number of employees trained so far.")]
setnames(IT, "Total number of employees trained so far.", "Initial Basic Fire Safety Workers Trained")

# If factory is in phase 3 or 4 and had phase 1, 2, 3a, or 4a remove phase 1, 2, 3a, or 4a rows
# Training$`Training Phase`[Training$`Training Phase` == "3a"] <- "3"
# Training$`Training Phase`[Training$`Training Phase` == "4a"] <- "4"
# Training$`Training Phase` <- as.numeric(Training$`Training Phase`)
# Training = arrange(Training, desc(`Training Phase`))
Training = Training %>%
  mutate(`Training Phase` =  factor(`Training Phase`, levels = c("4", "3", "4a", "3a", "2", "1"))) %>%
  arrange(`Training Phase`)
Training = distinct(Training, `Account ID`, .keep_all = TRUE)
Training$`Refresher Training` <- "No"
Training$`Refresher Training`[Training$`Training Phase` == "3"] <- "Yes"
Training$`Refresher Training`[Training$`Training Phase` == "4"] <- "Yes"

Training$`Total number of employees trained so far.`[Training$`Refresher Training` == "No"] <- 0

table(Training$STATUS, useNA = "ifany")

# Join the tables
Training = left_join(Training, IT, by = "Account ID")
Combined = left_join(Combined, Training, by = "Account ID")

#Combined$`Refresher Training` = ifelse(Combined$`Refresher Training` != "Yes", "No")


##### Security Guard Training ####
SG_Training <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Security Guard Training Implementation.xlsx", sep = ""), "SG Factory List")

# Remove unnecessary columns
# SG_Training[, c(4:30, 34:37, 47:59)] <- list(NULL)
SG_Training = SG_Training[, c("Account ID", "Total number of security staff trained so far", "Training Phase", "Percentage of security staff trained", "STATUS", "Spot Check Results (Pass or Fail)", "Support Visit")]

# Subset initial training (phase 1)
SG_IT = SG_Training[SG_Training$`Training Phase` == 1,]
SG_IT = SG_IT[, c("Account ID", "Total number of security staff trained so far")]
setnames(SG_IT, "Total number of security staff trained so far", "Initial Security Guards Trained")

# If factory is in phase 2 and had phase 1, remove phase 1 rows
SG_Training$`Training Phase` <- as.numeric(SG_Training$`Training Phase`)
SG_Training = arrange(SG_Training, desc(`Training Phase`))
SG_Training = distinct(SG_Training, `Account ID`, .keep_all = TRUE)
SG_Training$`Security Guard Refresher Training` <- "No"
SG_Training$`Security Guard Refresher Training`[!is.na(SG_Training$`Training Phase`) & SG_Training$`Training Phase` == 2] <- "Yes"

SG_Training$`Total number of security staff trained so far`[SG_Training$`Security Guard Refresher Training` == "No"] <- 0

# Join the tables
SG_Training = left_join(SG_Training, SG_IT, by = "Account ID")
SG_Training$`Account ID` = as.numeric(SG_Training$`Account ID`)
Combined = left_join(Combined, SG_Training, by = "Account ID")


#### Helpline ####
## Helpline calls ##
H_calls <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", sep = ""), "Caller Group breakdown", skip = 1)
H_calls$`Row Labels` <- as.numeric(H_calls$`Row Labels`)
H_calls = H_calls[complete.cases(H_calls$`Row Labels`),]

## Helpline urgent safety calls by reason ##
H_reasons <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", sep = ""), "Urgent Safety breakdown", skip = 3)
H_reasons$`Row Labels` = as.numeric(H_reasons$`Row Labels`)
H_reasons = H_reasons[complete.cases(H_reasons$`Row Labels`),]

## Helpline factories ##
H_factories <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Factory Profile Update_Helpline.xlsx", sep = ""), "Helpline List")
H_factories = H_factories[, c("Account ID", "Workers Trained")]
H_factories$`Account ID` = as.numeric(H_factories$`Account ID`)
H_factories = H_factories[complete.cases(H_factories$`Account ID`), ]
H_factories$Implemented <- "Yes"

# Join the tables
Helpline = left_join(H_factories, H_calls, by = c("Account ID" = "Row Labels"))
Helpline = left_join(Helpline, H_reasons, by = c("Account ID" = "Row Labels"))
Combined = left_join(Combined, Helpline, by = "Account ID")

# Add "No" to Helpline Implemented column
Combined$Implemented[is.na(Combined$Implemented)] <- "No"

#### Safety Committees ####
SC <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/SC Implementation.xlsx", sep = ""), "Factory List - SC", col_types = "text")

SC = SC[complete.cases(SC$`Account ID`),]
setnames(SC, names(SC), gsub("\\r\\n", " ", names(SC)))
# SC = SC[SC$`SC Formation  (Yes/No) %` > 0,]
# SC = SC[!is.na(SC$`SC Formation  (Yes/No) %`),]
SC = SC[, c("Account ID", "PC / CBA or TU / WWA (Yes/No) % (10)" , "SC Formation  (Yes/No) % (10)", "SC Formation Date", 
            "SC Formation Process", "Total SC members", "TtT Received from Alliance (Yes/No) % (10)", 
            "Number of Participants", "Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators", 
            "SC Activity Implementation Completion Date (Total - 100 Days)", "Action Plan Submitted (Yes/No) % (Total % - 5)", 
            "Introduction of SC (Yes/No) (5 %)", "Conducted Training for rest of SC members by  Factory Facilitators (Yes/No) % (10)", 
            "Safety Policy &  Emergency Response Procedure  (Yes/No) % (5)", "Risk Assessment by Safety Committee & Follow Up (Yes/No) % Total (10)", 
            "Formal SC Meeting Arrangement (Yes/No) % Total (10)", "Conduct Fire Or Evacuation Drill (Yes/No) % Total (10)", 
            "Issues Report (Yes/No) % Total (5)", "Safety Record Book Prepare (Yes/No) (5 %)", 
            "Final Percentage (Achievement) %", "Status")]

# setnames(SC, "TtT Received from Alliance (Yes/No) %", "TtT Received from Alliance")
setnames(SC, "SC Activity Implementation Completion Date (Total - 100 Days)", "SC Activity Implementation Completion Date")
# SC$Status <- gsub("^([a-z])", "\\U\\1", tolower(SC$Status), perl=TRUE)

# Change these columns to numeric
SC$`Total SC members` <- as.numeric(SC$`Total SC members`)
SC$`Number of Participants` <- as.numeric(SC$`Number of Participants`)
SC$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators` <- as.numeric(SC$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators`)
SC$`Account ID` <- as.numeric(SC$`Account ID`)

# Save number of factory facilitors and rest of SC members to re-add back in later
SC2 = SC[, c("Account ID", "Total SC members", "Number of Participants", "Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators")]

# Change time values to completed, in progress, or NA
SC[SC == "10"] <- "Completed"
SC[SC == "1"] <- "In progress"
SC[SC == "2"] <- "In progress"
SC[SC == "3"] <- "In progress"
SC[SC == "4"] <- "In progress"
SC[SC == "5"] <- "In progress"
SC[SC == "6"] <- "In progress"
SC[SC == "7"] <- "In progress"
SC[SC == "8"] <- "In progress"
SC[SC == "9"] <- "In progress"
SC[SC == "3.33"] <- "In progress"
SC[SC == "6.66"] <- "In progress"

# Re-add the above three columns
SC$`Total SC members` = SC2$`Total SC members`
SC$`Number of Participants` = SC2$`Number of Participants`
SC$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators` = SC2$`Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators`


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

# If in Master Factory List but not an expansion inspection in the FFC, remove Expansion Remediation Status
Combined$`Expansion Remediation Status` = ifelse(Combined$Expansion == "Yes", Combined$`Expansion Remediation Status`, NA)

# Remove duplicate rows
Combined = Combined[!duplicated(Combined[,"Account ID"]),]


# Save the combined data, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Combined, "Combined.csv", na="")



#### Reorder the columns of Combined to match the Dashboard Workbook more closely ####
someCol <- c("Account ID", "Account Name.x", "Active Brands", "Number of Active Members", "Remediation Factory Status", "Safety Management Visit Status", "Accord Shared/Alliance only info", "Inspected by Alliance / Accord / All&Acc", "Expansion Remediation Status", "Number of employees to be trained.", "Province", "RENTED", "Mixed Occupancy", "Factory housing in multi-factory building", "Case Group", 
             "Escalation Date", "Escalation Status", "Review Panel", 
             "Electrical High - Completed", "Electrical High - In progress", "Electrical High - Not Started", "Electrical High Total",
             "Electrical Medium - Completed", "Electrical Medium - In progress", "Electrical Medium - Not Started", "Electrical Medium Total",
             "Electrical Low - Completed", "Electrical Low - In progress", "Electrical Low - Not Started", "Electrical Low Total",
             "Electrical - No Level - Completed", "Electrical - No Level - In progress", "Electrical - No Level - Not started", "Electrical - No Level - Total",
             "Fire High - Completed", "Fire High - In progress", "Fire High - Not Started", "Fire High Total",
             "Fire Medium - Completed", "Fire Medium - In progress", "Fire Medium - Not Started", "Fire Medium Total",
             "Fire Low - Completed", "Fire Low - In progress", "Fire Low - Not Started", "Fire Low Total",
             "Fire - No Level - Completed", "Fire - No Level - In progress", "Fire - No Level - Not started", "Fire - No Level - Total",
             "Structural High - Completed", "Structural High - In progress", "Structural High - Not Started", "Structural High Total",
             "Structural Medium - Completed", "Structural Medium - In progress", "Structural Medium - Not Started", "Structural Medium Total",
             "Structural Low - Completed", "Structural Low - In progress", "Structural Low - Not Started", "Structural Low Total",
             "Structural - No Level - Completed", "Structural - No Level - In progress", "Structural - No Level - Not started", "Structural - No Level - Total",
             "# of Highest Priority NCs", "# of Highest Priority NCs completed", "% of Highest Priority NCs completed", 
             "Date of Initial Inspection", "CAP Approval Date", 
             "Actual Date of 1st RVV", "Completed - RVV1", "In progress - RVV1", "Not started - RVV1",
             "Confirmed Date of 2nd RVV", "Completed - RVV2", "In progress - RVV2", "Not started - RVV2",
             "Confirmed Date of 3rd RVV", "Completed - RVV3", "In progress - RVV3", "Not started - RVV3",
             "Confirmed Date of 4th RVV", "Completed - RVV4", "In progress - RVV4", "Not started - RVV4",
             "Confirmed Date of 5th RVV", "Completed - RVV5", "In progress - RVV5", "Not started - RVV5",
             "Confirmed Date of 6th RVV", "Completed - RVV6", "In progress - RVV6", "Not started - RVV6",
             "Confirmed Date of 7th RVV", "Completed - RVV7", "In progress - RVV7", "Not started - RVV7",
             "Confirmed Date of 8th RVV", "Completed - RVV8", "In progress - RVV8", "Not started - RVV8",
             "CCVV 1 Date", "CCVV 1 % of Completion", "CCVV 1 Result", "CCVV 2 Date", "CCVV 2 % of Completion", "CCVV 2 Result", "CCVV 3 Date", "CCVV 3 % of Completion", "CCVV 3 Result",
             "Safety Management Visit Date",
             "Retrofitting Status", "DEA Status", "Design Status", "Central Fire Status", "Hydrant Status", "Sprinkler Status", "Fire Door Status", "Lightning Status", "Single Line Diagram Status",
             "Initial Basic Fire Safety Workers Trained", "Refresher Training", "Total number of employees trained so far.", "Percentage of Workers Trained", "STATUS.x", "Final Training Status \r\n(CCVV)", "Final Training Assessment (CCVV) \r\nCurrent Results \r\n(Pass or Fail)", "Support Visit Required?",
             "Initial Security Guards Trained", "Security Guard Refresher Training", "Total number of security staff trained so far", "Percentage of security staff trained", "STATUS.y", "Spot Check Results (Pass or Fail)", "Support Visit",
             "PC / CBA or TU / WWA (Yes/No) % (10)", "SC Formation  (Yes/No) % (10)", "SC Formation Date", "SC Formation Process", "TtT Received from Alliance (Yes/No) % (10)", "Number of Participants", "Total Number of Participants in Factory Training for rest of SC members by Factory Facilitators", "SC Activity Implementation Completion Date", "Status",
             "Implemented", "Workers Trained", "General Inquiries", "No Category", "Non-urgent: Non-safety", "Non-urgent: Safety", "Urgent: Non-safety", "Urgent: Safety",
             "Fire - Active (factory)", "Fire - Danger (factory)", "Locked factory exit or blocked egress route", "Other", "Sparking / short circuit", "Structural - Cracks in beams, columns or walls", "Structural - walls or windows shaking", "Unattended / bare electric wires", "Unauthorized subcontracting",
             "Box Folder Link")

setcolorder(Combined, c(someCol, colnames(Combined)[!(colnames(Combined) %in% someCol)]))

#if the above failed, try this to see which of someCol is not in Combined. Then change the name to match.
setdiff(someCol, colnames(Combined))
names(Combined)

# Save the combined data, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Combined, "Combined.csv", na="")

# # Change values for factories with 100% CAP completion
# Combined$`Electrical High - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Electrical High Total`
# Combined$`Electrical High - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Electrical High - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Electrical Medium - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Electrical Medium Total`
# Combined$`Electrical Medium - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Electrical Medium - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Electrical Low - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Electrical Low Total`
# Combined$`Electrical Low - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Electrical Low - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Electrical - No Level - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Electrical - No Level - Total`
# Combined$`Electrical - No Level - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Electrical - No Level - Not started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Fire High - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Fire High Total`
# Combined$`Fire High - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Fire High - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Fire Medium - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Fire Medium Total`
# Combined$`Fire Medium - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Fire Medium - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Fire Low - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Fire Low Total`
# Combined$`Fire Low - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Fire Low - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Fire - No Level - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Fire - No Level - Total`
# Combined$`Fire - No Level - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Fire - No Level - Not started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Structural High - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Structural High Total`
# Combined$`Structural High - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Structural High - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Structural Medium - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Structural Medium Total`
# Combined$`Structural Medium - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Structural Medium - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Structural Low - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Structural Low Total`
# Combined$`Structural Low - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Structural Low - Not Started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`Structural - No Level - Completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`Structural - No Level - Total`
# Combined$`Structural - No Level - In progress`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# Combined$`Structural - No Level - Not started`[Combined$`CCVV 1 % of Completion` == 1] <- NA
# 
# Combined$`# of Highest Priority NCs completed`[Combined$`CCVV 1 % of Completion` == 1] <- Combined$`# of Highest Priority NCs`
# Combined$`% of Highest Priority NCs completed`[Combined$`CCVV 1 % of Completion` == 1] <- 1



#### Dashboard Workbook ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions

wd = dirname(getwd())
wd = sub("/OneDrive", "", wd)

Combined <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Dashboard Workbook.xlsm", sep = ""), "Combined", skip = 2)

Factory_Monthly <- read_excel(paste(wd,"/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Dashboard Workbook.xlsm", sep = ""), "Factory Monthly", skip = 1)

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
Combined = Combined[, c(1,2,5:10)]
Combined = Combined[complete.cases(Combined$`Account ID`),]
Factory_Monthly = Factory_Monthly[, -2]
Factory_Monthly = Factory_Monthly[complete.cases(Factory_Monthly$`Account ID_1`),]

# Join old Factory Monthly to Combined
Mon = left_join(Combined, Factory_Monthly, by = c("Account ID" = "Account ID_1"))

# Sort columns
allmisscols = apply(Mon, 2, function(x) {all(is.na(x))})
colswithallmiss = names(allmisscols[allmisscols>0])
colswithallmiss[3]

ind = grep(paste("^", gsub("_.*","",colswithallmiss[3]), sep = ""), names(Mon))

Mon[, ind[1]] = Mon$`Overall Status`
Mon[, ind[2]] = Mon$`Remediation Status`
Mon[, ind[3]] = Mon$`Worker Training Status`
Mon[, ind[4]] = Mon$`Security Guard Training Status`
Mon[, ind[5]] = Mon$`Helpline Status`
Mon[, ind[6]] = Mon$`Safety Committee Status`

# Remove unnecessary columns
Mon = Mon[, -c(3:8)]

# Save the new Factory Monthly list, then copy and paste into the appropriate columns in the Excel Dashboard Workbook
write.csv(Mon, "Factory Monthly.csv", na="")


