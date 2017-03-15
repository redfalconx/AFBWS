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
CAP_Tracker <- read_excel("C:/Users/Andrew/Box Sync/Database/Excel/CAP Trackers.xlsm", "CAP Trackers", skip = 1)

# Load Audit Scope #
Electrical <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Electrical")
Fire <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Fire")
Structural <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/Audit Scope.xlsx", "Structural")

Audit_Scope <- rbindlist(list(Electrical, Fire, Structural))

setnames(Audit_Scope, names(Audit_Scope), gsub("\\r\\n", " ", names(Audit_Scope)))

# Load list of corrected factories #
Cor_Factories <- read_excel("C:/Users/Andrew/Box Sync/FFC/Data Migration/List of factories completed.xlsx", 1)

# Subset factories into those that have been corrected
l = c(Cor_Factories$`Account ID`)

CAPs = CAP_Tracker[CAP_Tracker$`FFC ID` %in% l,]

# Remove rows with no Question
CAPs = CAPs[CAPs$Question != 0,]

# Change 0 to NA to for matching to Audit Scope
CAPs$Level[CAPs$Level == 0] <- NA

# Add columns that check whether the Subheader, Question, and Priority Level match those in Audit Scope
CAPs$Question_check = CAPs$Question %in% Audit_Scope$Question

CAPs$Subheader_check = ifelse(CAPs$Subheader == Audit_Scope[match(CAPs$Question, Audit_Scope$Question), 1], TRUE, FALSE)

CAPs$Level_check = ifelse(CAPs$Level == Audit_Scope[match(CAPs$Question, Audit_Scope$Question), 4], TRUE, FALSE)

# Correct some misspellings of Completed, In-progress, Not Started, or N/A
CAPs$`Alliance Remarks 5` = as.character(CAPs$`Alliance Remarks 5`)

table(CAPs$`Alliance Remarks 1`)
table(CAPs$`Alliance Remarks 2`)
table(CAPs$`Alliance Remarks 3`)
table(CAPs$`Alliance Remarks 4`)
table(CAPs$`Alliance Remarks 5`)

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "completed"] <- "Completed"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "completed"] <- "Completed"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "In-Progress"] <- "In-progress"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "In-Progress"] <- "In-progress"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "Not started"] <- "Not Started"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "Not started"] <- "Not Started"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "Not Applicable"] <- "N/A"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "Not Applicable"] <- "N/A"

CAPs$`Alliance Remarks 1`[CAPs$`Alliance Remarks 1` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 2`[CAPs$`Alliance Remarks 2` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 3`[CAPs$`Alliance Remarks 3` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 4`[CAPs$`Alliance Remarks 4` == "NA"] <- "N/A"
CAPs$`Alliance Remarks 5`[CAPs$`Alliance Remarks 5` == "NA"] <- "N/A"

# Add columns that check whether the Alliance Remarks statuses match Completed, In-progress, Not Started, or N/A
Statuses = c("Completed", "In-progress", "Not Started", "N/A", "0")

CAPs$RVV1_check = CAPs$`Alliance Remarks 1` %in% Statuses
CAPs$RVV2_check = CAPs$`Alliance Remarks 2` %in% Statuses
CAPs$RVV3_check = CAPs$`Alliance Remarks 3` %in% Statuses
CAPs$RVV4_check = CAPs$`Alliance Remarks 4` %in% Statuses
CAPs$RVV5_check = CAPs$`Alliance Remarks 5` %in% Statuses

# Do not mark as FALSE if previous RVVs not conducted
CAPs$RVV2_check = ifelse(CAPs$`Alliance Remarks 1` != 0 & CAPs$`Alliance Remarks 3` != 0 & CAPs$`Alliance Remarks 2` == 0, FALSE, CAPs$RVV2_check)
CAPs$RVV3_check = ifelse(CAPs$`Alliance Remarks 2` != 0 & CAPs$`Alliance Remarks 4` != 0 & CAPs$`Alliance Remarks 3` == 0, FALSE, CAPs$RVV3_check)
CAPs$RVV4_check = ifelse(CAPs$`Alliance Remarks 3` != 0 & CAPs$`Alliance Remarks 5` != 0 & CAPs$`Alliance Remarks 4` == 0, FALSE, CAPs$RVV4_check)

# Remove unnecessary columns
CAPs[, c(1:2, 44)] <- list(NULL)

# Remove rows with TRUE for all checks
CAPs = CAPs[CAPs$Question_check == FALSE | CAPs$Subheader_check == FALSE | CAPs$Level_check == FALSE 
            | CAPs$RVV1_check == FALSE | CAPs$RVV2_check == FALSE | CAPs$RVV3_check == FALSE 
            | CAPs$RVV4_check == FALSE | CAPs$RVV5_check == FALSE, ]
CAPs = CAPs[complete.cases(CAPs$`FFC ID`),]

# Save the file in FFC > Data Migration
write.csv(CAPs, "/Users/Andrew/Box Sync/FFC/Data Migration/Errors in CAPs.csv", na="")
