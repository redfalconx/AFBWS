#### Load packages ####
library(data.table) # converts to data tables
library(readxl) # reads Excel files
library(dplyr) # data manipulation
library(tidyr) # a few pivot-table functions
library(ggplot2) # plotting
library(scales) # works with ggplot2 to properly label axes on plots
library(dygraphs)
library(xts)

#### Load Helpline data ####
Helpline <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/AK_Full Dataset/Comprehensive Helpline Issues (Updated Monthly)/Amader Kotha - Dec1 - July31 (Alliance) - AR.xlsx")
# Helpline_PreDec <- read_excel("C:/Users/Andrew/Dropbox (AFBWS.org)/AK_Full Dataset/Comprehensive Helpline Issues (Updated Monthly)/Amader Kotha - Pre Dec 2014 (Alliance).xlsx", "Workers")


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

