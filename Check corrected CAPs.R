# Created by Andrew Russell, 2017.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

## This script will match files that have been corrected by the FFC admins to the Audit Scope file to check for accuracy

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions


# Load raw CAP data #
CAP_Tracker <- read_excel("C:/Users/Andrew/Box Sync/Database/Excel/CAP Trackers.xlsm", "CAP Trackers", skip = 1, col_types = "text")
# CAP_Tracker$`FFC ID` <- gsub("/E", "", CAP_Tracker$`FFC ID`)
# CAP_Tracker$`FFC ID` = as.numeric(CAP_Tracker$`FFC ID`)
# CAP_Tracker = CAP_Tracker[ , 1:43]

# Load Audit Scope #
Electrical <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Electrical")
Fire <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Fire")
Structural <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Structural")

Audit_Scope <- rbindlist(list(Electrical, Fire, Structural))

setnames(Audit_Scope, names(Audit_Scope), gsub("\\r\\n", " ", names(Audit_Scope)))

# Load list of corrected factories #
Cor_Factories_1 <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/List of factories completed.xlsx", 1)
Cor_Factories_2 <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/List of factories completed.xlsx", 2)
Cor_Factories_3 <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/List of factories completed.xlsx", 3)
Cor_Factories_4 <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/List of factories completed.xlsx", 4)

# Subset factories into those that have been corrected
l = c(Cor_Factories_1$`Account ID`, Cor_Factories_2$`Account ID`, Cor_Factories_3$`Account ID`, Cor_Factories_4$`Account ID`)
l = as.character(l)

CAPs = CAP_Tracker[CAP_Tracker$`FFC ID` %in% l,]

# Add factory names
f = as.data.table(l)
f$`Factory Name` = c(Cor_Factories_1$`Factory Name`, Cor_Factories_2$`Factory Name`, Cor_Factories_3$`Factory Name`, Cor_Factories_4$`Factory Name`)

CAPs = left_join(CAPs, f, by = c("FFC ID" = "l"))

CAPs = CAPs[, c(1:2, ncol(CAPs), 3:(ncol(CAPs)-1))]

# Remove rows with no Question
CAPs = CAPs[CAPs$Question != 0,]

# Change 0 to NA to for matching to Audit Scope
CAPs$Level[CAPs$Level == 0] <- NA

# Change from Highest to original priority level
CAPs$Level[CAPs$Level == "Highest"] <- "High"
CAPs$Level[CAPs$Question == "Illuminated exit signs are provided with battery backup or emergency power and are continuously illuminated.\r\n"] <- "Medium"
CAPs$Level[CAPs$Question == "Means of egress have a minimum ceiling height of 2.3 m (7 ft 6 in.) with projections from the ceiling not less than 2.03 m (6 ft 8 in.)."] <- "Medium"
CAPs$Level[CAPs$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded? \r\n"] <- "Medium"
CAPs$Level[CAPs$Question == "Electrical connections at equipment, fixtures, etc are properly secured.\r\n"] <- "Medium"

# Add columns that check whether the Subheader, Question, and Priority Level match those in Audit Scope
CAPs$Question_check = CAPs$Question %in% Audit_Scope$Question

CAPs$Subheader_check = ifelse(CAPs$Subheader == Audit_Scope[match(CAPs$Question, Audit_Scope$Question), 1], TRUE, FALSE)
CAPs$Subheader_check = ifelse(CAPs$Subheader == "Fire Safety Programs" & CAPs$Question == "Are there additional areas of non-compliance to report?", TRUE, CAPs$Subheader_check)

CAPs$Level_check = ifelse(CAPs$Level == Audit_Scope[match(CAPs$Question, Audit_Scope$Question), 4], TRUE, FALSE)
#CAPs$Level_check = ifelse(is.na(Audit_Scope[match(CAPs$Question, Audit_Scope$Question), 4]), TRUE, CAPs$Level_check)

# Correct some misspellings of Completed, In-progress, Not Started, or N/A
table(CAPs$`Alliance Remarks 1`, useNA = "ifany")
table(CAPs$`Alliance Remarks 2`, useNA = "ifany")
table(CAPs$`Alliance Remarks 3`, useNA = "ifany")
table(CAPs$`Alliance Remarks 4`, useNA = "ifany")
table(CAPs$`Alliance Remarks 5`, useNA = "ifany")
table(CAPs$`Alliance Remarks 6`, useNA = "ifany")

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 6`[CAPs$`Alliance Remarks 6` == "completed"] <- "Completed"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 6`[CAPs$`Alliance Remarks 6` == "In-Progress"] <- "In-progress"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 6`[CAPs$`Alliance Remarks 6` == "Not started"] <- "Not Started"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 6`[CAPs$`Alliance Remarks 6` == "Not Applicable"] <- "N/A"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 6`[CAPs$`Alliance Remarks 6` == "NA"] <- "N/A"

# Add columns that check whether the Alliance Remarks statuses match Completed, In-progress, Not Started, or N/A
Statuses = c("Completed", "In-progress", "Not Started", "N/A", "0")

CAPs$RVV1_check = CAPs$`Alliance Remarks 1` %in% Statuses
CAPs$RVV2_check = CAPs$`Alliance Remarks 2` %in% Statuses
CAPs$RVV3_check = CAPs$`Alliance Remarks 3` %in% Statuses
CAPs$RVV4_check = CAPs$`Alliance Remarks 4` %in% Statuses
CAPs$RVV5_check = CAPs$`Alliance Remarks 5` %in% Statuses
CAPs$RVV6_check = CAPs$`Alliance Remarks 6` %in% Statuses

# Do not mark as FALSE if previous RVVs not conducted
CAPs$RVV2_check = ifelse(CAPs$`Alliance Remarks 1` != "0" & CAPs$`Alliance Remarks 3` != "0" & CAPs$`Alliance Remarks 2` == "0", FALSE, CAPs$RVV2_check)
CAPs$RVV3_check = ifelse(CAPs$`Alliance Remarks 2` != "0" & CAPs$`Alliance Remarks 4` != "0" & CAPs$`Alliance Remarks 3` == "0", FALSE, CAPs$RVV3_check)
CAPs$RVV4_check = ifelse(CAPs$`Alliance Remarks 3` != "0" & CAPs$`Alliance Remarks 5` != "0" & CAPs$`Alliance Remarks 4` == "0", FALSE, CAPs$RVV4_check)
CAPs$RVV5_check = ifelse(CAPs$`Alliance Remarks 4` != "0" & CAPs$`Alliance Remarks 6` != "0" & CAPs$`Alliance Remarks 5` == "0", FALSE, CAPs$RVV5_check)
CAPs$RVV6_check = ifelse(CAPs$`Alliance Remarks 5` != "0" & is.na(CAPs$`Alliance Remarks 6`) == TRUE, FALSE, CAPs$RVV6_check)

# Mark each RVV check as FALSE if all are blank
CAPs$RVV1_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV1_check)
CAPs$RVV2_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV2_check)
CAPs$RVV3_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV3_check)
CAPs$RVV4_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV4_check)
CAPs$RVV5_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV5_check)
CAPs$RVV6_check = ifelse(CAPs$`Alliance Remarks 1` == "0" & CAPs$`Alliance Remarks 2` == "0" & CAPs$`Alliance Remarks 3` == "0" & CAPs$`Alliance Remarks 4` == "0" & CAPs$`Alliance Remarks 5` == "0" & CAPs$`Alliance Remarks 6` == "0", FALSE, CAPs$RVV6_check)

# Remove unnecessary columns
CAPs[, c(1:2)] <- list(NULL)

# Remove rows with TRUE for all checks
CAPs = CAPs[CAPs$Question_check == FALSE | CAPs$Subheader_check == FALSE | CAPs$Level_check == FALSE 
            | CAPs$RVV1_check == FALSE | CAPs$RVV2_check == FALSE | CAPs$RVV3_check == FALSE 
            | CAPs$RVV4_check == FALSE | CAPs$RVV5_check == FALSE | CAPs$RVV6_check == FALSE, ]
CAPs = CAPs[complete.cases(CAPs$`FFC ID`),]

# Save the file in FFC > Data Migration
write.csv(CAPs, "/Users/Andrew/Box Sync/FFC/Data Migration/Errors in CAPs.csv", na="")

# Created list of validated CAPs
Validated_Factories = as.data.table(setdiff(l, CAPs$`FFC ID`))

setnames(Validated_Factories, "V1", "Validated Factories")

write.csv(Validated_Factories, "/Users/Andrew/Box Sync/FFC/Data Migration/Validated Factories.csv", na="")



#### Filter CAP Tracker for validated CAPs only ####
Validated_CAPs = CAP_Tracker[CAP_Tracker$`FFC ID` %in% Validated_Factories$`Validated Factories`,]

# Remove factories already uploaded to the FFC
Cor_Factories_1 = subset(Cor_Factories_1, is.na(`FFC Status`))
Cor_Factories_2 = subset(Cor_Factories_2, is.na(`FFC Status`))
Cor_Factories_3 = subset(Cor_Factories_3, is.na(`FFC Status`))
Cor_Factories_4 = subset(Cor_Factories_4, is.na(`FFC Status`))

l = c(Cor_Factories_1$`Account ID`, Cor_Factories_2$`Account ID`, Cor_Factories_3$`Account ID`, Cor_Factories_4$`Account ID`)
l = as.character(l)

Validated_CAPs = Validated_CAPs[Validated_CAPs$`FFC ID` %in% l,]

# Remove rows with no Question
Validated_CAPs = Validated_CAPs[Validated_CAPs$Question != 0,]

# Change 0 to NA for matching to Audit Scope
Validated_CAPs$Level[Validated_CAPs$Level == 0] <- NA

# Change from Highest to original priority level
Validated_CAPs$Level[Validated_CAPs$Level == "Highest"] <- "High"
Validated_CAPs$Level[Validated_CAPs$Question == "Illuminated exit signs are provided with battery backup or emergency power and are continuously illuminated.\r\n"] <- "Medium"
Validated_CAPs$Level[Validated_CAPs$Question == "Means of egress have a minimum ceiling height of 2.3 m (7 ft 6 in.) with projections from the ceiling not less than 2.03 m (6 ft 8 in.)."] <- "Medium"
Validated_CAPs$Level[Validated_CAPs$Question == "Is a program in place to ensure that the live loads for which a floor or roof is or has been designed will not be exceeded? \r\n"] <- "Medium"
Validated_CAPs$Level[Validated_CAPs$Question == "Electrical connections at equipment, fixtures, etc are properly secured.\r\n"] <- "Medium"

# Correct some misspellings of Completed, In-progress, Not Started, or N/A
Validated_CAPs$`Alliance Remarks 1`[Validated_CAPs$`Alliance Remarks 1` == "completed"] <- "Completed"
Validated_CAPs$`Alliance Remarks 2`[Validated_CAPs$`Alliance Remarks 2` == "completed"] <- "Completed"
Validated_CAPs$`Alliance Remarks 3`[Validated_CAPs$`Alliance Remarks 3` == "completed"] <- "Completed"
Validated_CAPs$`Alliance Remarks 4`[Validated_CAPs$`Alliance Remarks 4` == "completed"] <- "Completed"
Validated_CAPs$`Alliance Remarks 5`[Validated_CAPs$`Alliance Remarks 5` == "completed"] <- "Completed"
Validated_CAPs$`Alliance Remarks 6`[Validated_CAPs$`Alliance Remarks 6` == "completed"] <- "Completed"

Validated_CAPs$`Alliance Remarks 1`[Validated_CAPs$`Alliance Remarks 1` == "In-Progress"] <- "In-progress"
Validated_CAPs$`Alliance Remarks 2`[Validated_CAPs$`Alliance Remarks 2` == "In-Progress"] <- "In-progress"
Validated_CAPs$`Alliance Remarks 3`[Validated_CAPs$`Alliance Remarks 3` == "In-Progress"] <- "In-progress"
Validated_CAPs$`Alliance Remarks 4`[Validated_CAPs$`Alliance Remarks 4` == "In-Progress"] <- "In-progress"
Validated_CAPs$`Alliance Remarks 5`[Validated_CAPs$`Alliance Remarks 5` == "In-Progress"] <- "In-progress"
Validated_CAPs$`Alliance Remarks 6`[Validated_CAPs$`Alliance Remarks 6` == "In-Progress"] <- "In-progress"

Validated_CAPs$`Alliance Remarks 1`[Validated_CAPs$`Alliance Remarks 1` == "Not started"] <- "Not Started"
Validated_CAPs$`Alliance Remarks 2`[Validated_CAPs$`Alliance Remarks 2` == "Not started"] <- "Not Started"
Validated_CAPs$`Alliance Remarks 3`[Validated_CAPs$`Alliance Remarks 3` == "Not started"] <- "Not Started"
Validated_CAPs$`Alliance Remarks 4`[Validated_CAPs$`Alliance Remarks 4` == "Not started"] <- "Not Started"
Validated_CAPs$`Alliance Remarks 5`[Validated_CAPs$`Alliance Remarks 5` == "Not started"] <- "Not Started"
Validated_CAPs$`Alliance Remarks 6`[Validated_CAPs$`Alliance Remarks 6` == "Not started"] <- "Not Started"

Validated_CAPs$`Alliance Remarks 1`[Validated_CAPs$`Alliance Remarks 1` == "Not Applicable"] <- "N/A"
Validated_CAPs$`Alliance Remarks 2`[Validated_CAPs$`Alliance Remarks 2` == "Not Applicable"] <- "N/A"
Validated_CAPs$`Alliance Remarks 3`[Validated_CAPs$`Alliance Remarks 3` == "Not Applicable"] <- "N/A"
Validated_CAPs$`Alliance Remarks 4`[Validated_CAPs$`Alliance Remarks 4` == "Not Applicable"] <- "N/A"
Validated_CAPs$`Alliance Remarks 5`[Validated_CAPs$`Alliance Remarks 5` == "Not Applicable"] <- "N/A"
Validated_CAPs$`Alliance Remarks 6`[Validated_CAPs$`Alliance Remarks 6` == "Not Applicable"] <- "N/A"

Validated_CAPs$`Alliance Remarks 1`[Validated_CAPs$`Alliance Remarks 1` == "NA"] <- "N/A"
Validated_CAPs$`Alliance Remarks 2`[Validated_CAPs$`Alliance Remarks 2` == "NA"] <- "N/A"
Validated_CAPs$`Alliance Remarks 3`[Validated_CAPs$`Alliance Remarks 3` == "NA"] <- "N/A"
Validated_CAPs$`Alliance Remarks 4`[Validated_CAPs$`Alliance Remarks 4` == "NA"] <- "N/A"
Validated_CAPs$`Alliance Remarks 5`[Validated_CAPs$`Alliance Remarks 5` == "NA"] <- "N/A"
Validated_CAPs$`Alliance Remarks 6`[Validated_CAPs$`Alliance Remarks 6` == "NA"] <- "N/A"

table(Validated_CAPs$`Alliance Remarks 1`, useNA = "ifany")
table(Validated_CAPs$`Alliance Remarks 2`, useNA = "ifany")
table(Validated_CAPs$`Alliance Remarks 3`, useNA = "ifany")
table(Validated_CAPs$`Alliance Remarks 4`, useNA = "ifany")
table(Validated_CAPs$`Alliance Remarks 5`, useNA = "ifany")
table(Validated_CAPs$`Alliance Remarks 6`, useNA = "ifany")

# Remove unnecessary columns and rows
Validated_CAPs[, c(1:2)] <- list(NULL)
Validated_CAPs = Validated_CAPs[complete.cases(Validated_CAPs$`FFC ID`),]

# Add Audit IDs
Audit_IDs <- read.csv("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit IDs.csv")
Audit_IDs$Account.ID <- as.character(Audit_IDs$Account.ID)
Validated_CAPs = left_join(Validated_CAPs, Audit_IDs, by = c("FFC ID" = "Account.ID", "Sheet" = "Audit.Scope"))

# Move new columns to the front, delete unnecessary columns
Validated_CAPs = Validated_CAPs[c((ncol(Validated_CAPs)-1):ncol(Validated_CAPs), 1:(ncol(Validated_CAPs)-2))]
Validated_CAPs = Validated_CAPs[c(1, 3, 2, (4:ncol(Validated_CAPs)))]

Validated_CAPs = Validated_CAPs[, -c(5, 21, 24, 25, 27:29, 31:33, 35:37, 39:41, 43:45, 47)]

setnames(Validated_CAPs, "FFC ID", "Account ID")

# Load Master Factory List #
Master <- read_excel("C:/Users/Andrew/Box Sync/Alliance Factory info sheet/Master Factory Status/MASTER Factory Status.xlsx", "Master Factory List")

# Clean up Master, remove unnecessary columns
Master = Master[complete.cases(Master$`Account ID`),]
Master = Master[, c("Account ID", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "Confirmed Date of 5th RVV", "Confirmed Date of 6th RVV", "CCVV 1 Date")]
setcolorder(Master, c("Account ID", "Actual Date of 1st RVV", "Confirmed Date of 2nd RVV", "Confirmed Date of 3rd RVV", "Confirmed Date of 4th RVV", "Confirmed Date of 5th RVV", "Confirmed Date of 6th RVV", "CCVV 1 Date"))
# Master$`Account ID` <- as.numeric(Master$`Account ID`)
Master = Master[complete.cases(Master$`Account ID`),]

# Join the RVV dates to the Validated CAPs
Validated_CAPs = left_join(Validated_CAPs, Master, by = "Account ID")

# Reorder columns
Validated_CAPs = Validated_CAPs[c(1:21, 28, 22, 29, 23, 30, 24, 31, 25, 32, 26, 33, 27, 34)]


# setcolorder(Validated_CAPs, c("FFC ID", "Audit.ID", "Sheet", "Subheader", "Question", "Description", "Suggested Plan of Action", "Suggested Deadline Date", "Standard", "Factory CAP", "Factory CAP Deadline Date", "Factory Responsible Person", "Source of Findings", "Level"))


# Save the file in FFC > Data Migration
write.csv(Validated_CAPs, "/Users/Andrew/Box Sync/FFC/Data Migration/Validated_CAPs.csv", na="")
