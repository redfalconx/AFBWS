#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions
library(plyr) # data manipulation
library(zoo) # time-series functions
library(ggplot2) # plotting
library(scales) # works with ggplot2 to properly label axes on plots
library(dygraphs) # for making interactive graphs
library(xts) # time-series functions

#### Load Helpline data ####
Helpline <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", 1)
# Helpline_PreDec <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/AK_Full Dataset/Comprehensive Helpline Issues (Updated Monthly)/Amader Kotha - Pre Dec 2014 (Alliance).xlsx", "Workers")
RS <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", "Resolution Status", skip = 1)

# H2 <- read_excel("C:/Users/Andrew/Box Sync/Member Reporting/Dashboards/Dashboard Workbook/Amader Kotha - Helpline Data.xlsx", 1, skip = 7500, col_names = FALSE)

# Create new column that combines Caller Group and Reason
Helpline = mutate(Helpline, `Caller Group: Reason` = paste(`Caller Group`, Reason, sep = ": "))


# Get sum totals for reasons and resoltuions by month by category

H_caller_group_reason_count = Helpline %>%
  group_by(`Factory FFC Number`, year = year(`Call Date`), month = month(`Call Date`), `Caller Group: Reason`) %>%
  summarise(count = n_distinct(Sl.)) %>%
  spread(`Caller Group: Reason`, count, fill = 0) %>%
  arrange(`Factory FFC Number`, year, month)

H_caller_group_reason_count = mutate(H_caller_group_reason_count, Date = as.POSIXct(paste(year, month, "01", sep = "-")))

# H_caller_group_resolution_count = Helpline %>%
#  group_by(`Factory FFC Number`, year = year(`Call Date`), month = month(`Call Date`), `Resolution Status`) %>%
#  summarise(count = n_distinct(Sl.)) %>%
#  spread(`Resolution Status`, count, fill = 0) %>%
#  arrange(`Factory FFC Number`, year, month)

# H_caller_group_resolution_count = mutate(H_caller_group_resolution_count, Date = as.POSIXct(paste(year, month, "01", sep = "-")))

# Convert to data table
H_caller_group_reason_count <- as.data.table(H_caller_group_reason_count)

# All of the relevant dates
ts <- seq.POSIXt(as.POSIXlt("2014-12-01"), as.POSIXlt(Sys.Date() - 30), by="month")
dates.all = H_caller_group_reason_count[, ts, by = `Factory FFC Number`]

# Set the key and merge filling in the blanks with missing dates
setkey(H_caller_group_reason_count, `Factory FFC Number`, Date)
H_caller_group_reason_count <- H_caller_group_reason_count[dates.all, roll = T]

# Format Timestamp to yyyy-mm
H_caller_group_reason_count$Date <- format(H_caller_group_reason_count$Date, "%b-%Y")

# Remove year and month columns
H_caller_group_reason_count[, 2:3] <- list(NULL)

# Move Date to second column
H_caller_group_reason_count <- subset(H_caller_group_reason_count, select=c(1, ncol(H_caller_group_reason_count), 2:(ncol(H_caller_group_reason_count)-1)))

# Join with Resolution Statuses
RS$`Row Labels` <- as.character(RS$`Row Labels`)
H_caller_group_reason_count = left_join(H_caller_group_reason_count, RS, by = c("Factory FFC Number" = "Row Labels"))

# Save the file
write.csv(H_caller_group_reason_count, "Helpline calls by factory by month.csv", na = "0")



# Breakdown data by call types
calls <- Helpline %>% group_by(`Factory FFC Number`) %>% summarise(count = n_distinct(Sl.))
calls <- calls[-nrow(calls),]
calls <- calls[-nrow(calls),]

hist(calls$count, breaks = 100, main = "Histogram of all calls", xlab = "Number of calls per factory", ylab = "Number of factories")
plot(density(calls$count))
hist(calls$count[calls$count < 200 ], breaks = 100, main = "Histogram of all calls, excluding outliers", xlab = "Number of calls per factory", ylab = "Number of factories")

substantive_calls <- subset(Helpline, `Caller Group` != "No Category")
substantive_calls <- subset(substantive_calls, `Caller Group` != "General Inquiries")

substantive_calls_count <- substantive_calls %>% group_by(`Factory FFC Number`) %>% summarise(count = n_distinct(Sl.))
substantive_calls_count <- substantive_calls_count[-nrow(substantive_calls_count),]
substantive_calls_count <- substantive_calls_count[-nrow(substantive_calls_count),]

hist(substantive_calls_count$count, breaks = 100, main = "Histogram of substantive calls", xlab = "Number of calls per factory", ylab = "Number of factories")
plot(density(substantive_calls_count$count))
hist(substantive_calls_count$count[substantive_calls_count$count < 200 ], breaks = 100, main = "Histogram of substantive calls, excluding outliers", xlab = "Number of calls per factory", ylab = "Number of factories")

# Get sum totals for substantive calls by month by category
Helpline_sub_sum = Helpline %>% 
  group_by(year = year(`Call Date`), month = month(`Call Date`), `Caller Group`) %>% 
  summarise(count = n_distinct(Sl.)) %>%
  spread(`Caller Group`, count, fill = 0) %>%
  arrange(year, month)

Helpline_sub_sum = Helpline %>% 
  group_by(`Call Date`, `Caller Group`) %>% 
  summarise(count = n_distinct(Sl.)) %>%
  spread(`Caller Group`, count, fill = 0)

# Plot it!
substantive_calls$Month <- as.Date(cut(substantive_calls$`Call Date`, breaks = "month"))

ggplot(substantive_calls, aes(Month, Sl.)) +
  ggtitle("Substantive calls by category by month") + labs(y = "Number of substantive calls") +
  stat_summary(fun.y = n_distinct, geom = "line", aes(colour = `Caller Group`)) + 
  scale_x_date(breaks = pretty_breaks(10))
  

# Interactive charts!
dygraph(xts(Helpline_sub_sum[,2:7], order.by = Helpline_sub_sum$`Call Date`, unique = FALSE, tzone = "")) %>% 
  dyRangeSelector()
dygraph(xts(Helpline_sub_sum[,2:7], order.by = Helpline_sub_sum$`Call Date`, unique = FALSE, tzone = "")) %>% 
  dyRangeSelector() %>% dyRoller(rollPeriod = 7) %>% dyOptions(stepPlot = TRUE)
dygraph(xts(Helpline_sub_sum[,2:7], order.by = to.monthly(Helpline_sub_sum$`Call Date`), unique = FALSE, tzone = "")) %>% 
  dyRangeSelector()

ts(pcp, frequency = 12, start = 2001)

lungDeaths <- cbind(ldeaths, mdeaths, fdeaths)
View(lungDeaths)
ld <- as.xts(lungDeaths)

data <- xts(Helpline_sub_sum[,2:7], order.by = Helpline_sub_sum$`Call Date`, unique = FALSE, tzone = "")

