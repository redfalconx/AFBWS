# Created by Andrew Russell, 2015.
# Randomize Helpline phone numbers for UC Berkeley analysis project

library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation

wd = dirname(getwd())

# Read in raw Helpline data, then assign unique ID to Phone #s
H <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", 1)
D = as.data.table(H[, "Phone"])
D$Phone = as.character(H$Phone)
D$Source = "Helpline Data"

comb <- with(D, Phone)
D <- within(D, Phone_ID <- match(comb, unique(comb)))

H$Phone = D$Phone_ID
setnames(H, "Phone", "Phone_ID")

write.csv(D, "Helpline Data with Unique ID for Phone Numbers.csv", na="")

# Copy Phone_ID column into Helpline Excel file in Dashboard Workbook folder, then save in UC Berkeley Box folder
# and remove Phone number column

write.xlsx2(H, "C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/UCB - Helpline Data.xlsx", sheetName = "Comprehensive", col.names=TRUE, row.names=FALSE, append=FALSE)

wb1 <- createWorkbook()
addWorksheet(wb1, sheetName = "Comprehensive")
writeData(wb1, sheet = "Comprehensive", H, colNames = T)

wb <- loadWorkbook("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx")
writeData(wb, sheet = "Comprehensive", H, colNames = T)
saveWorkbook(wb, "C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/UCB - Helpline Data.xlsx", overwrite = T)

# Read in raw Helpline Satisfaction Survey data
HS <- as.data.table(read.csv("C:/Users/Andrew/Desktop/Helpline Satisfaction - Raw Data w Num.csv", header=TRUE))
HS = HS[, .(Phone.Number)]
HS$Source = "Helpline Satisfaction Survey"
HS$Phone = substr(HS$Phone.Number, 4, 13)
HS = HS[, .(Phone)]

# Read in raw Worker Perception Survey data
WP <- as.data.table(read.csv("C:/Users/Andrew/Desktop/Worker Perception - Raw Data w Num.csv", header=TRUE))
WP = WP[, .(Phone.Number)]
WP$Source = "Worker Perception Survey"
WP$Phone = substr(HS$Phone.Number, 4, 13)
WP = WP[, .(Phone)]

# Append all datasets
C <- bind_rows(HS, WP)
C <- bind_rows(C, H)

comb <- with(C, Phone)
D <- within(C, Phone_ID <- match(comb, unique(comb)))

write.csv(D, "Helpline Unique ID for Phone Numbers.csv", na="")

# Open csv and use Offset function to match phone numbers to new Unqiue IDs in each dataset (like below)
# Copy Phone_ID column into Helpline Excel file in Dashboard Workbook folder, then save in UC Berkeley Box folder
# and remove Phone number column