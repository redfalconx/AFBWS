# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

# Load Active factories from FFC.R script
source('~/R/AFBWS/FFC.R', echo=TRUE)

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions


#### Master Factory List ####
# Fetch the Master table from the Excel spreadsheet and put the results in a dataframe
Master <- read_excel("C:/Users/Andrew/Box Sync/Alliance Factory info sheet/Master Factory Status 2016/MASTER Factory Status_2016-July 21.xlsx", "Master Factory List")
# Actives <- read_excel("C:/Users/Andrew/Desktop/FFC Actives.xlsx", 1)

# Remove (Active) from members' names
Actives$`Active Members (Display)` <- gsub(" (Active)", "", Actives$`Active Members (Display)`, fixed = TRUE)
# Remove " ," from members' names
Actives$`Active Members (Display)` <- gsub(" ,", ",", Actives$`Active Members (Display)`, fixed = TRUE)

# Join the tables
New_Master = left_join(Master, Actives, by = "Account ID")

## Do an anti-join to check if there are new factories
# Remove Li & Fung
New_Actives = subset(Actives, Actives$`Active Members (Display)` != "Li & Fung")
# Remove NAs
# New_Actives = New_Actives[complete.cases(Actives$`Active Members (Display)`),]
New_Actives = subset(Actives, Actives$`Active Members (Display)` != "")
new_factories = anti_join(New_Actives, Master, by = "Account ID")

write.csv(new_factories, "New_Factories.csv", na = "")

# Check if any of the factories have new members or if members left
New_Master$Diff_Members <- ifelse(New_Master$`Active Members (Display)` != New_Master$`Active Brands`, 1, 0)

# If there are differences, the sum will be more than zero
sum(New_Master$Diff_Members, na.rm = TRUE)


# If Li & Fung is one of the members, substract 1 from number of brands and add *
New_Master$`New_Number of Active Members` <- ifelse(grepl("Li & Fung", New_Master$`Active Members (Display)`) == TRUE, 
                                                   New_Master$`Number of Active Members.y` - 1, New_Master$`Number of Active Members.y`)
# Remove Li & Fung
New_Master$`Active Members (Display)` <- replace(New_Master$`Active Members (Display)`, New_Master$`Active Members (Display)` == "Li & Fung", NA)
# Add *
New_Master$`New_Number of Active Members` <- as.character(New_Master$`New_Number of Active Members`)
New_Master$`New_Number of Active Members` <- ifelse(grepl("Li & Fung", New_Master$`Active Members (Display)`) == TRUE, 
                                                   paste(New_Master$`New_Number of Active Members`, "*"), New_Master$`New_Number of Active Members`)

# Update Accord Shared Factories
# New_Master$`ACTIVE Accord Shared Factories` <- replace(New_Master$`ACTIVE Accord Shared Factories`, New_Master$`ACTIVE Accord Shared Factories` == "No", "Yes")

# Save the file
write.csv(New_Master, file = "New_Master.csv", na = "")


#### Suspended Factories ####
# Fetch the Master table from the Excel spreadsheet and put the results in a dataframe
Suspended <- read_excel("C:/Users/Andrew/Box Sync/Alliance Factory info sheet/Master Factory Status 2016/MASTER Factory Status_2016-July 21.xlsx", "Suspended Factories")

# Join the tables
New_Suspended = left_join(Suspended, Actives, by = "Account ID")

# Check if any of the factories have new members or if members left
New_Suspended$Diff_Members <- ifelse(New_Suspended$`Active Members (Display)` != New_Suspended$`Active Brands`, 1, 0)

# If there are differences, the sum will be more than zero
sum(New_Suspended$Diff_Members, na.rm = TRUE)


# If Li & Fung is one of the members, substract 1 from number of brands and add *
New_Suspended$`New_Number of Active Members` <- ifelse(grepl("Li & Fung", New_Suspended$`Active Members (Display)`) == TRUE, 
                                                    New_Suspended$`Number of Active Members.y` - 1, New_Suspended$`Number of Active Members.y`)
# Remove Li & Fung
New_Suspended$`Active Members (Display)` <- replace(New_Suspended$`Active Members (Display)`, New_Suspended$`Active Members (Display)` == "Li & Fung", NA)
# Add *
New_Suspended$`New_Number of Active Members` <- as.character(New_Suspended$`New_Number of Active Members`)
New_Suspended$`New_Number of Active Members` <- ifelse(grepl("Li & Fung", New_Suspended$`Active Members (Display)`) == TRUE, 
                                                    paste(New_Suspended$`New_Number of Active Members`, "*"), New_Suspended$`New_Number of Active Members`)

# Save the file
write.csv(New_Suspended, file = "New_Suspended.csv", na = "")


#### Training ####
# Fetch the Train the Trainer table from the Excel spreadsheet and put the results in a dataframe
Training <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/Basic Fire Safety & Helpline Training Implementation_Imran.xlsx", 1)
# Actives = subset(Actives, Actives$`Active Members (Display)` != "Li & Fung")

# Join the tables
New_Training = left_join(Training, Actives, by = "Account ID")

# Do an anti-join to check if there are new factories
new_factories = anti_join(New_Actives, Training, by = "Account ID")

# Add row names column (because arrange removes it), then arrange by Training Phase
New_Training$rn = as.numeric(rownames(New_Training))
New_Training = arrange(New_Training, `Training Phase`)

# Check if any of the factories have new members or if members left
New_Training$Diff_Members <- ifelse(New_Training$`Active Members (Display)` != New_Training$`Active Members`, 1, 0)
sum(New_Training$Diff_Members, na.rm = TRUE)

New_Training$Members <- ifelse(duplicated(New_Training$`Account ID`, fromLast = TRUE) == TRUE, New_Training$`Active Members`, New_Training$`Active Members (Display)`)

New_Training$Diff_Members <- ifelse(New_Training$Members != New_Training$`Active Members`, 1, 0)
sum(New_Training$Diff_Members, na.rm = TRUE)

# Reorder by original order when it was imported
New_Training <- New_Training[order(New_Training$rn),]

# Save the file
write.csv(New_Training, file = "New_Training.csv", na = "")


#### Security Guard Training ####
# Fetch the Security Guard Training table from the Excel spreadsheet and put the results in a dataframe
SG_Training <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/Security Guard Training Implementation_Imran.xlsx", 1)

# SG_Training <- subset(SG_Training, `Factory Name` != "NA")

# Join the tables
New_SG_Training = left_join(SG_Training, Actives, by = "Account ID")

# Do an anti-join to check if there are new factories
new_factories = anti_join(New_Actives, SG_Training, by = "Account ID")

# Check if any of the factories have new members or if members left
New_SG_Training$Diff_Members <- ifelse(New_SG_Training$`Active Members (Display)` != New_SG_Training$`Active Members`, 1, 0)

sum(New_SG_Training$Diff_Members, na.rm = TRUE)

# Save the file
write.csv(New_SG_Training, file = "New_SG_Training.csv", na = "")


#### Safety Committees ####
# Fetch the Safety Committees table from the Excel spreadsheet and put the results in a dataframe
SC <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/SC Implementation_MM.xlsx", 1)

# Join the tables
New_SC = left_join(SC, Actives, by = "Account ID")

# Check if any of the factories have new members or if members left
New_SC$Diff_Members <- ifelse(New_SC$`Active Members (Display)` != New_SC$`Active Members`, 1, 0)

sum(New_SC$Diff_Members, na.rm = TRUE)

# Save the file
write.csv(New_SC, file = "New_SC.csv", na = "")


#### Plan Review Tracker ####
# Fetch the Plan Review Tracker table from the Excel spreadsheet and put the results in a dataframe
PR <- read_excel("C:/Users/Andrew/Box Sync/Alliance Design Approval/Master Tracker/Master Tracker.xls", 1, skip = 1)

# Join the tables
New_PR = left_join(PR, Actives, by = c("ID" = "Account ID"))

# Check if any of the factories have new members or if members left
New_PR$Diff_Members <- ifelse(New_PR$`Active Members (Display)` != New_PR$`BRAND`, 1, 0)

sum(New_PR$Diff_Members, na.rm = TRUE)

# Save the file
write.csv(New_PR, file = "New_PR.csv", na = "")


#### Inactives ####
# Read the Inactive in FFC file
Inactive_in_FFC <- read_excel("C:/Users/Andrew/Desktop/Inactive in FFC.xlsx", 1)
Actives <- read_excel("C:/Users/Andrew/Desktop/FFC Master.xlsx", 1)

# Remove (Active) from members' names
Actives$`Active Members (Display)` <- gsub(" (Active)", "", Actives$`Active Members (Display)`, fixed = TRUE)

# Subset to get a list of inactive factories
# Inactives = anti_join(Master, Inactives, by = "Account ID")
Inactives = subset(Actives, !complete.cases(Actives$`Active Members (Display)`))
Inactives = inner_join(Master, Inactives, by = "Account ID")

# Subset lists to remove unnecessary columns
# Inactives = subset(Inactives, Inactives$`Account ID` != "NA" & Inactives$`Account ID` > 999)
setnames(Inactives, 42, "Date Notification Sent (to factory & brands) 2")
# Inactive_in_FFC = Inactive_in_FFC[, ]
New_Training = select(New_Training, `Account ID`, `Training Phase`, `TtT date`, `First Day of TtT Training`, `Number of employees to be trained.`, `Total number of employees trained so far (AL).`:`STATUS`)
New_SG_Training = New_SG_Training[, -4]
New_SG_Training = select(New_SG_Training, `Account ID`:`Factory Name`, `Start date (first session)`, `Number of security staff to be trained.`, ` Total number of security staff trained so far`:`STATUS `)
Inactives = Inactives[, 1:65]

# Remove duplicate rows in Inactive_in_FFC
Inactive_in_FFC <- Inactive_in_FFC[order(Inactive_in_FFC$`Status Last Modified Date`, decreasing = TRUE),]
Inactive_in_FFC <- Inactive_in_FFC[!duplicated(Inactive_in_FFC$`Account ID`),]

# Join the table with Training spreadsheet and Security Guard Training spreadsheet
Inactives = left_join(Inactives, New_Training, by = "Account ID")
Inactives = left_join(Inactives, New_SG_Training, by = "Account ID")
Inactives = left_join(Inactives, Inactive_in_FFC, by = "Account ID")

# Remove unnecessary columns
# Inactives = select(Inactives, `Account Name`, `Account ID`, `Deactivated brands (Date)`:`Pilot Final Inspection % of Completion`, `Dropbox Folder Link`, `Training Phase`, `TtT date`:`Action Plan Submission Date`, `Number of employees to be trained.`, `Last Online Submission`, `Total number of employees trained so far (AL).`:`Date when the status is changed. `)
Inactives = Inactives[, c(-3:-9, -12, -18, -22, -25:-30, -35, -38, -42:-64, -73)]

# Save the file
write.csv(Inactives, file = "Inactives.csv", na = "")


#### Case Managers ####
# Fetch the Case Manager tables from the Excel spreadsheet and put the results in a dataframe
CM1 <- read_excel("C:/Users/Andrew/Box Sync/Alliance - Reports and Cap/Case Manager-1 (momtaz ala shibbir ahmed)/3. CAP Tracker/Factory Status of Case Manager-1 till 26 August 15.xls", 1)
CM2 <- read_excel("C:/Users/Andrew/Box Sync/Alliance - Reports and Cap/Case Manager-2 (M M Motiur Rahman)/3. CAP Tracker/Tracker Case Manager 2_02.xlsx", 1)
CM3 <- read_excel("C:/Users/Andrew/Box Sync/Alliance - Reports and Cap/Case Manager-3/3. CAP Tracker/06 Sep Factory Status of Case Manager 3 till September, 2015- Nusrat .xls", 1, skip = 3)

# CM3 <- CM3[-1:-2,]
# CM3 <- data.frame(CM3, header = T)
  
# Add all 3 together
All_CM = rbind(CM1[, 1:4], CM2[, 1:4], CM3[, 1:4])

# Join the tables
New_CM1 = left_join(CM1, Actives, by = "Account ID")
New_CM2 = left_join(CM2, Actives, by = "Account ID")
New_CM3 = left_join(CM3, Actives, by = "Account ID")
Actives = left_join(Actives, Master, by = "Account ID")
Unassigned_CM = left_join(Actives, All_CM, by = "Account ID")

# Check if any of the factories have new members or if members left
New_CM1$Diff_Members <- ifelse(New_CM1$`Active Members (Display)` != New_CM1$`Active Members`, 1, 0)
New_CM2$Diff_Members <- ifelse(New_CM2$`Active Members (Display)` != New_CM2$`Active Members`, 1, 0)
New_CM3$Diff_Members <- ifelse(New_CM3$`Active Members (Display)` != New_CM3$`Active Members`, 1, 0)

sum(New_CM1$Diff_Members, na.rm = TRUE)
sum(New_CM2$Diff_Members, na.rm = TRUE)
sum(New_CM3$Diff_Members, na.rm = TRUE)

# Save the file
write.csv(New_CM1, file = "New_CM1.csv")
write.csv(New_CM2, file = "New_CM2.csv")
write.csv(New_CM3, file = "New_CM3.csv")
write.csv(Unassigned_CM, file = "Unassigned_CM.csv")


#### Compare remediation to training for factories ####
# Join the tables
comb_Master_Training <- left_join(New_Training, New_Master, by = "Account ID")
comb_Master_SG_Training <- left_join(New_SG_Training, New_Master, by = "Account ID")
comb_Training_SG_Training <- left_join(New_SG_Training, New_Training, by = "Account ID")
comb_All <- left_join(comb_Master_SG_Training, comb_Training_SG_Training, by = "Account ID")

# Correlation of security staff trained to remediation percentage
cor(comb_All$`Percentage of security staff trained.x`, comb_All$`% Overall Remediation Complete in 1st RVV`, use = "complete.obs")
# Correlation of staff trained to remediation percentage
cor(comb_All$`12. Percentage of workers trained`, comb_All$`% Overall Remediation Complete in 1st RVV`, use = "complete.obs")
# Correlation of security staff trained to staff trained
cor(comb_All$`Percentage of security staff trained.x`, comb_All$`12. Percentage of workers trained`, use = "complete.obs")

