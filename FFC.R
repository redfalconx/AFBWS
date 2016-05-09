# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("RSelenium")

#### Load packages ####
library(RCurl)
library(rvest)
library(RSelenium)
library(data.table) # converts to data tables
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions


#### RSelenium code ####
# RSelenium::checkForServer() # install server if needed
RSelenium::startServer() # if needed
remDr <- remoteDriver(browserName = "chrome")
remDr$open()
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

# Click on the Alliance Factory Report 2 query
Alliance_Factory_Report_2 <- remDr$findElement("css selector", '[href*="ctl00$ContentPlaceHolder1$dgReport$ctl14$ctl00"]')
Alliance_Factory_Report_2$highlightElement()
remDr$executeScript("arguments[0].click();", list(Alliance_Factory_Report_2))

# Click on the Export button
Export <- remDr$findElement("id", "btnExport")
Export$clickElement()

# Download the file
Alink <- remDr$findElement("xpath", "//*[@id='flexExportStatus']/a")

Aurl <- Mlink$getElementAttribute("href")[[1]]
Alink$clickElement()

# Open the file
Afile <- sub(".*filename=", "", Aurl)
Afile <- "C:/Users/Andrew/Downloads/" & Actives
Actives <- as.data.frame(readHTMLTable(paste("C:/Users/Andrew/Downloads/", Afile, sep = "")))

# Change column names
setnames(Actives, names(Actives), gsub("NULL.", "", names(Actives)))
setnames(Actives, names(Actives), gsub("\\.", " ", names(Actives)))

# Change ID column to numeric
Actives$`Account ID` <- as.numeric(Actives$`Account ID`)
