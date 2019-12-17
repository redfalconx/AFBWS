# Created by Andrew Russell, 2017.
# These packages are used at various points: 
# install.packages("data.table", "readxl", "dplyr", "tidyr")

## This script will pull statuses from each program area to automate the website statuses

#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions
library(RCurl) #  push to ftp server

wd = dirname(getwd())

#### Fetch Data #### 
### Master Factory List ###
# Fetch the Master table from the Excel spreadsheet and put the results in a dataframe
Master <- read_excel("C:/Users/Andrew/Box Sync/Alliance Factory info sheet/Master Factory Status/MASTER Factory Status.xlsx", "Master Factory List", col_types = "text")
Master = Master[complete.cases(Master$`Account ID`),]
Master = Master[, c("Account Name", "Account ID", "Remediation Factory Status", "Inspected by Alliance / Accord / All&Acc")]


### Training ###
# Fetch the Train the Trainer table from the Excel spreadsheet and put the results in a dataframe
Training <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/Basic Fire Safety & Helpline Training Implementation.xlsx", 1, col_types = "text")
Training = Training[complete.cases(Training$`Account ID`),]
Training = Training[complete.cases(Training$`Training Phase`),]
Training = Training[, c("Account ID", "Training Phase", "STATUS", "Helpline Comment")]


### Security Guard Training ###
# Fetch the Security Guard Training table from the Excel spreadsheet and put the results in a dataframe
SG_Training <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/Security Guard Training Implementation.xlsx", 1, col_types = "text")
SG_Training = SG_Training[complete.cases(SG_Training$`Account ID`),]
SG_Training = SG_Training[, c("Account ID", "Training Phase", "STATUS")]


### Safety Committees ###
# Fetch the Safety Committees table from the Excel spreadsheet and put the results in a dataframe
SC <- read_excel("C:/Users/Andrew/Box Sync/Training and Worker Empowerment Programs/7. Training Implementation Trackers/SC Implementation.xlsx", 1, col_types = "text")
SC = SC[complete.cases(SC$`Account ID`),]
SC = SC[, c("Account ID", "Status")]


#### Rename Statuses ####
### Master Factory List ###
# Change "Subs. Complete" to "Completed"
Master$`Remediation Factory Status`[grepl("Subs", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE] = "Completed"

# Change from "Completed Under Accord" to "Completed"
Master$`Remediation Factory Status`[grepl("completed", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE] = "Completed"

# Change from "CCVV Pending Awaiting CAP Closure of Unoccupied Expansion" to "On track"
Master$`Remediation Factory Status`[grepl("CCVV Pending", Master$`Remediation Factory Status`, ignore.case = TRUE) == TRUE] = "On track"

table(Master$`Remediation Factory Status`, useNA = "ifany")


### Helpline ###
# Sort by Training Phase
Training = Training %>%
  mutate(`Training Phase` =  factor(`Training Phase`, levels = c("4", "3", "4a", "3a", "2", "1"))) %>%
  arrange(`Training Phase`)

# Create Helpline column with "Not started" as default
Helpline = Training
Helpline$Helpline = "Not started"

# If in Helpline pilot, then status is "Completed"
Helpline$Helpline[grepl("pilot", Training$`Helpline Comment`, ignore.case = TRUE) == TRUE] = "Completed"

# If in anything but phase 1 or blank, then status is "Completed"
Helpline$Helpline[Helpline$`Training Phase` != "1"] = "Completed"

# Subset, and remove duplicates
Helpline = distinct(Helpline, `Account ID`, .keep_all = TRUE)
Helpline = Helpline[, c("Account ID", "Helpline")]
table(Helpline$Helpline, useNA = "ifany")


### Training ###
# If factory is in phase 3 or 4 and had phase 1 or 2, remove phase 1 or 2 rows
Training = distinct(Training, `Account ID`, .keep_all = TRUE)
Training$`Refresher Training` = "No"
Training$`Refresher Training`[Training$`Training Phase` == "3"] = "Yes"
Training$`Refresher Training`[Training$`Training Phase` == "4"] = "Yes"

Training = Training[, c("Account ID", "STATUS", "Refresher Training")]
setnames(Training, "STATUS", "Training Status")

# If status is "critical" or "unwilling", then "Critical"
Training$`Training Status`[grepl("critical", Training$`Training Status`, ignore.case = TRUE) == TRUE | grepl("unwilling", Training$`Training Status`, ignore.case = TRUE) == TRUE] = "Critical"

# If status is "not on track" or "Needs intervention", then "Needs intervention"
Training$`Training Status`[grepl("not on track", Training$`Training Status`, ignore.case = TRUE) == TRUE | grepl("needs intervention", Training$`Training Status`, ignore.case = TRUE) == TRUE] = "Needs intervention"

# If status is "on track", then "On track"
Training$`Training Status`[grepl("on track", Training$`Training Status`, ignore.case = TRUE) == TRUE] = "On track"

# If status is "schedule" or "not started", then "Not started"
Training$`Training Status`[grepl("schedule", Training$`Training Status`, ignore.case = TRUE) == TRUE | grepl("not started", Training$`Training Status`, ignore.case = TRUE) == TRUE] = "Not started"

# If status is "complete", then "Completed"
Training$`Training Status`[grepl("complete", Training$`Training Status`, ignore.case = TRUE) == TRUE] = "Completed"

table(Training$`Training Status`, useNA = "ifany")


### SG Training ###
# If factory is in phase 2 and had phase 1 , remove phase 1 rows
SG_Training$`Training Phase` <- as.numeric(SG_Training$`Training Phase`)
SG_Training = arrange(SG_Training, desc(`Training Phase`))
SG_Training = distinct(SG_Training, `Account ID`, .keep_all = TRUE)
SG_Training$`SG Refresher Training` = "No"
SG_Training$`SG Refresher Training`[SG_Training$`Training Phase` == 2] = "Yes"

SG_Training = SG_Training[, c("Account ID", "STATUS", "SG Refresher Training")]
setnames(SG_Training, "STATUS", "SG Training Status")

# If status is "critical" or "unwilling", then "Critical"
SG_Training$`SG Training Status`[grepl("critical", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE | grepl("unwilling", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE] = "Critical"

# If status is "not on track", then "Needs intervention"
SG_Training$`SG Training Status`[grepl("not on track", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE] = "Needs intervention"

# If status is "on track", then "On track"
SG_Training$`SG Training Status`[grepl("on track", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE] = "On track"

# If status is "schedule" or "not started", then "Not started"
SG_Training$`SG Training Status`[grepl("schedule", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE | grepl("not started", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE] = "Not started"

# If status is "complete", then "Completed"
SG_Training$`SG Training Status`[grepl("complete", SG_Training$`SG Training Status`, ignore.case = TRUE) == TRUE] = "Completed"

table(SG_Training$`SG Training Status`, useNA = "ifany")


### SC ###
# If "elect" in status, change to "Needs WPC election"
SC$Status[grepl("elect", SC$Status, ignore.case = TRUE) == TRUE] = "Needs WPC election"

# If "Ready for next TtT" in status, change to "Ready for training"
SC$Status[grepl("Ready for next TtT", SC$Status, ignore.case = TRUE) == TRUE] = "Ready for training"

# If "on track" in status, change to "On track"
SC$Status[grepl("on track", SC$Status, ignore.case = TRUE) == TRUE] = "On track"

# If "not assessed", blank, or starts with "Formal" in status, change to "Not assessed"
SC$Status[grepl("not assessed", SC$Status, ignore.case = TRUE) == TRUE | grepl("Formal", SC$Status, ignore.case = FALSE) == TRUE | is.na(SC$Status) == TRUE] = "Not assessed"

# If "Not on track", "critical", or "form" in status, change to "Needs intervention"
SC$Status[grepl("not on track", SC$Status, ignore.case = TRUE) == TRUE | grepl("critical", SC$Status, ignore.case = TRUE) == TRUE | grepl("form", SC$Status, ignore.case = TRUE) == TRUE] = "Needs intervention"

# If "Accord" in status, change to "Under Accord"
SC$Status[grepl("accord", SC$Status, ignore.case = TRUE) == TRUE] = "Under Accord"

# If "completed" in status, change to "Completed"
SC$Status[grepl("completed", SC$Status, ignore.case = TRUE) == TRUE] = "Completed"

# If not one of the above statuses, change to "Not assessed"
SC$Status[grepl("Needs WPC election", SC$Status, ignore.case = TRUE) == FALSE & 
            grepl("Ready for training", SC$Status, ignore.case = TRUE) == FALSE & 
            grepl("On track", SC$Status, ignore.case = TRUE) == FALSE & 
            grepl("Not assessed", SC$Status, ignore.case = TRUE) == FALSE &
            grepl("Needs intervention", SC$Status, ignore.case = TRUE) == FALSE & 
            grepl("Under Accord", SC$Status, ignore.case = TRUE) == FALSE &
            grepl("Completed", SC$Status, ignore.case = TRUE) == FALSE] = "Not assessed"

setnames(SC, "Status", "SC Status")

#### Join tables ####
Statuses = Master %>%
  left_join(Training, by = "Account ID") %>%
  left_join(SG_Training, by = "Account ID") %>%
  left_join(Helpline, by = "Account ID") # %>%
  # left_join(SC, by = "Account ID")


#### Add new columns ####
# Inspection column
Statuses$Inspection = "Completed"
Statuses$Inspection[Statuses$`Inspected by Alliance / Accord / All&Acc` == "New"] = "Not started"

# Overall Status column
Statuses$`Overall Status` = "Participating"
Statuses$`Overall Status`[Statuses$`Remediation Factory Status` == "Suspended"] = "Suspended"
Statuses$`Overall Status`[Statuses$`Remediation Factory Status` == "Removed"] = "Removed"


#### Reorder columns and missing values = "Not started" ####
Statuses = Statuses[, c("Account ID", "Account Name", "Inspection", "Remediation Factory Status", "Training Status", "SG Training Status", "Helpline", #"SC Status", 
                        "Overall Status", "Refresher Training", "SG Refresher Training")]
Statuses = Statuses[complete.cases(Statuses$`Account ID`), ]
Statuses[is.na(Statuses)] = "Not started"
Statuses$`Refresher Training`[Statuses$`Refresher Training` == "Not started"] = "No"
Statuses$`SG Refresher Training`[Statuses$`SG Refresher Training` == "Not started"] = "No"


#### Save the txt file ####
write.table(Statuses, "C:/Users/Andrew/Box Sync/FFC/Factory List/Monthly Website Lists/Status Lists/factory-statuses.txt", sep="\t", row.names=FALSE)


#### Upload to FTP Server ####
Sys.sleep(5)

# ftpUpload("C:/Users/Andrew/Box Sync/FFC/Factory List/Monthly Website Lists/Status Lists/factory-statuses.txt", "ftp://www.afbws.org/public_html/alliance/files/factory-lists/factory-statuses.txt", userpwd = "infactor:A!FBWS007$%2018")

ftpUpload("C:/Users/Andrew/Box Sync/FFC/Factory List/Monthly Website Lists/Status Lists/factory-statuses.txt", "sftp://www.afbws.org/home/infactor/public_html/alliance/files/factory-lists/factory-statuses.txt", userpwd = "root:ELEVATE2017@usa")
