# Created by Andrew Russell, 2016.
# These packages are used at various points: 
# install.packages("RSelenium", "data.table")

#### Load packages ####
library(RSelenium)
library(XML)
library(data.table) # converts to data tables

wd = dirname(getwd())

#### RSelenium code ####
# RSelenium::checkForServer() # install server if needed
#RSelenium::startServer() # if needed
eCaps = list(chromeOptions = list( #args = c('--headless', '--disable-gpu', '--window-size=1280,800'), 
  prefs = list("profile.default_content_settings.popups" = 0L, "download.prompt_for_download" = FALSE,
  "download.default_directory" = paste(wd,"/Documents/", sep = ""))))
rD <- rsDriver(extraCapabilities = eCaps)
remDr <- rD[["client"]]
Sys.sleep(3)
# remDr <- remoteDriver(browserName = "chrome")
# remDr$open()
remDr$getStatus() # Checks the status

# Navigate to the FFC and log in
appURL <- "http://sfi.fairfactories.org/ffcweb/Web/Common/LogonPage.aspx"
username <- "arussell@afbws.org"
password <- "Master00"
remDr$navigate(appURL)

userName <- remDr$findElement("id", "txtEmail")
userName$sendKeysToElement(list(username))
passWord <- remDr$findElement("id", "txtPassword")
passWord$sendKeysToElement(list(password))
logIn <- remDr$findElement("id", "btnLogin")
logIn$clickElement()

# Navigate to the saved queries page
remDr$navigate("http://sfi.fairfactories.org/ffcweb/Web/Reports/SavedConfiguredReportsList.aspx?id=PUBLIC")

# Click on the Monthly Factory List for website query
Monthly_Factory_List <- remDr$findElement("link text", 'Monthly Factory List for website')
Monthly_Factory_List$highlightElement()
remDr$executeScript("arguments[0].click();", list(Monthly_Factory_List))

# Click on the Export button
Export <- remDr$findElement("id", "btnExport")
Export$highlightElement()
Export$clickElement()
Sys.sleep(5)

# Download the file
Alink <- remDr$findElement("xpath", "//*[@id='flexExportStatus']/a")

Aurl <- Alink$getElementAttribute("href")[[1]]
Alink$highlightElement()
Alink$clickElement()
Sys.sleep(5)

# Open the file
Afile <- sub(".*filename=", "", Aurl)
Sys.sleep(5)

while (file.exists(paste(wd,"/Documents/", Afile, sep = "")) == FALSE) {
  Sys.sleep(5)
}

file.rename(paste(wd,"/Documents/", Afile, sep = ""), paste(wd,"/Documents/Monthly Factory List.xls", sep = ""))
MFL <- as.data.frame(readHTMLTable(paste(wd,"/Documents/Monthly Factory List.xls", sep = "")))
Sys.sleep(5)

# Close the browser and stop server
remDr$close()
rD[["server"]]$stop()

# Change column names
setnames(MFL, names(MFL), gsub("NULL.", "", names(MFL)))
setnames(MFL, names(MFL), gsub("\\.", " ", names(MFL)))
setnames(MFL, "Active Members  Display ", "Active Members (Display)")

# Remove Blanks and Li & Fung factories
MFL = MFL[MFL$`Active Members (Display)` != "" ,]
# MFL = MFL[MFL$`Active Members (Display)` != "Li & Fung (Active)" ,]

# Change Shared with Accord to "*"
MFL$`ACTIVE Accord Shared Factories` = ifelse(MFL$`ACTIVE Accord Shared Factories` != "", "*", "")

# Merge address columns, then remove them, then change order of columns
MFL$Address = paste(MFL$Address1, MFL$Address2, sep = " ")
MFL = MFL[, -c(5:6)]
MFL <- subset(MFL, select=c(1:4, 16, 5:15))

# Save as Confidential
write.csv(MFL, "Factory List_Confidential.csv", na = "")

# Remove Designation, ID, City, Active Members columns
MFL = MFL[, -c(2:3, 6, 16)]

# Save as Public
write.csv(MFL, "Factory List_Public.csv", na = "")

