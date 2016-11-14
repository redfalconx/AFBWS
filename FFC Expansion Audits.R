# Created by Andrew Russell, 2016.
# These packages are used at various points: 
# install.packages("RSelenium", "data.table")

#### Load packages ####
library(RSelenium)
library(data.table) # converts to data tables


#### RSelenium code ####
# RSelenium::checkForServer() # install server if needed
RSelenium::startServer() # if needed
Sys.sleep(3)
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

# Click on the Monthly Factory List for website query
Expansions <- remDr$findElement("css selector", '[href*="ctl00$ContentPlaceHolder1$dgReport$ctl22$ctl00"]')
Expansions$highlightElement()
remDr$executeScript("arguments[0].click();", list(Expansions))

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

# Open the file
Afile <- sub(".*filename=", "", Aurl)
Sys.sleep(5)
Expansions <- as.data.frame(readHTMLTable(paste("C:/Users/Andrew/Downloads/", Afile, sep = "")))
Sys.sleep(5)

# Close the browser
remDr$close()

# Change column names
setnames(Expansions, names(Expansions), gsub("NULL.", "", names(Expansions)))
setnames(Expansions, names(Expansions), gsub("\\.", " ", names(Expansions)))

