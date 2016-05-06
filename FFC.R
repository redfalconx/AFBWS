# Created by Andrew Russell, 2015.
# These packages are used at various points: 
# install.packages("RSelenium", "readxl")

#### Load packages ####
library(RCurl)
library(readxl) # reads Excel files
library(rvest)
library(RSelenium)

#### RCurl code ####
curl = getCurlHandle()
curlSetOpt(cookiejar = 'cookies.txt', followlocation = TRUE, autoreferer = TRUE, curl = curl)

html <- getURL('http://sfi.fairfactories.org/ffcweb/Web/Common/LogonPage.aspx', curl = curl)

viewstate <- as.character(sub('.*id="__VIEWSTATE" value="([0-9a-zA-Z+/=]*).*', '\\1', html))

params <- list(
  'ctl00$ContentPlaceHolder3$Login1$UserName'    = '<arussell@afbws.org>',
  'ctl00$ContentPlaceHolder3$Login1$Password'    = '<Master00>',
  'ctl00$ContentPlaceHolder3$Login1$LoginButton' = 'Log In',
  '__VIEWSTATE'                                  = viewstate
)

html = postForm('http://sfi.fairfactories.org/ffcweb/Web/Common/LogonPage.aspx', .params = params, curl = curl)

# grepl('Logout', html)

#### rvest code ####
s <- html_session("http://sfi.fairfactories.org/ffcweb/Web/Utilities/Dispatcher.aspx?responsepage=CONFIGURED_REPORT_INDEX_PAGE")
s1 <- s %>% jump_to("http://sfi.fairfactories.org/ffcweb/Web/Reports/SavedConfiguredReportsList.aspx?id=PUBLIC")
s1 %>% follow_link("Alliance Factory Report 2")
s2 <- s1 %>% jump_to("javascript:__doPostBack('ctl00$ContentPlaceHolder1$dgReport$ctl11$ctl00','')")
s1 %>% follow_link(url = "javascript:__doPostBack('ctl00$ContentPlaceHolder1$dgReport$ctl11$ctl00','')", css = ".mainContent a")



s <- html_session("http://sfi.fairfactories.org/ffcweb/Web/Common/HomePage.aspx")
s1 <- s %>% jump_to("http://sfi.fairfactories.org/ffcweb/Web/Utilities/Dispatcher.aspx?responsepage=CONFIGURED_REPORT_INDEX_PAGE")
s2 <- s1 %>% jump_to("http://sfi.fairfactories.org/ffcweb/Web/Reports/SavedConfiguredReportsList.aspx?id=PUBLIC")
s3 <- s2 %>% jump_to("javascript:__doPostBack('ctl00$ContentPlaceHolder1$dgReport$ctl11$ctl00','')")
s4 <- 
s1 %>% follow_link("Alliance Factory Report 2")
s2 <- s1 %>% jump_to("javascript:__doPostBack('ctl00$ContentPlaceHolder1$dgReport$ctl11$ctl00','')")

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
Master <- remDr$findElement("xpath", "//*[@id='flexExportStatus']/a")
file <- as.data.frame(Master$clickElement())

rawHTML <- paste(readLines(Master$clickElement()), collapse="\n")

file <- readHTMLTable(Master$clickElement())

file <- read_excel(Master$clickElement(), "Sheet1")

