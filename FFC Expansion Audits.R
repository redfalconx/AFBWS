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
               "download.default_directory" = paste(wd,"/Downloads/", sep = ""))))
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
Expansions <- remDr$findElement("link text", 'Expansion audits')
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
Sys.sleep(5)

# Open the file
Afile <- sub(".*filename=", "", Aurl)
Sys.sleep(5)

while (file.exists(paste(wd,"/Downloads/", Afile, sep = "")) == FALSE) {
  Sys.sleep(5)
}

file.rename(paste(wd,"/Downloads/", Afile, sep = ""), paste(wd,"/Downloads/Expansions.xls", sep = ""))
Expansions <- as.data.frame(readHTMLTable(paste(wd,"/Downloads/Expansions.xls", sep = "")))
Sys.sleep(5)

# Close the browser and stop server
remDr$close()
rD[["server"]]$stop()

# Change column names
setnames(Expansions, names(Expansions), gsub("NULL.", "", names(Expansions)))
setnames(Expansions, names(Expansions), gsub("\\.", " ", names(Expansions)))

