# Name: Bui Nguyen Thuy Nhu
# Student ID: K194141737

###-----------------------------END-OF-COURSE PROJECT-------------------------------

### Import library------------------------------------------------------------------

library(tidyverse)
library(readxl)
library(dplyr)
library(forcats)
library(ggplot2)
library(scales)
library(zoo)
library(pastecs)
library(forecast) #forecast
library(tseries) #adf.test
library(lmtest) #coeftest
library(stats) #Box.test
library(Metrics) #RMSE

### Import data---------------------------------------------------------------------

data = read_excel('K194141737.xlsx')

### Preprocess data-----------------------------------------------------------------

data$Time = as.yearqtr(data$Time,format = "%Y-%m-%d") # Convert dates to quarterly
print(data$Time)
options(scipen = 999) # Convert exponential notation to specific display
data[data == "NULL"] = NA # Transform string "NULL" to NA value
view(data) # Check data


# Count NA values in the variables I will choose
sum(is.na(data$TotalAssets))
sum(is.na(data$Cash))
sum(is.na(data$TotalDebt))
sum(is.na(data$TotalRevenue))
sum(is.na(data$NetWorkingCapital))

data = data[order(data$Time), ] # Reverse data transfer follow quarters
view(data) # Check data

# Create function calculate percentage change
pct <- function(x) {x / lag(x) - 1}

# Create variable Columns
as_tibble(data)
data = data %>%
  mutate(Cash_Holding = Cash/TotalAssets) %>% # Create Cash Holding Column
  mutate(Leverage = TotalDebt/TotalAssets) %>% # Create Leverage Column
  mutate(Firm_Size = log(TotalAssets)) %>% # Create Firm_Size Column
  mutate(Growth = pct(TotalRevenue)) # Create Liquidity Column

data$Growth[is.na(data$Growth)] = 0 # Fill the first cell of the Growth column as 0
view(data)

# Create variables to distinguish the period with Covid19
data$Covid19 = 0
data$Covid19[as.numeric(format(data$Time, format="%Y"))<2020] = "BEFORE" 
data$Covid19[as.numeric(format(data$Time, format="%Y"))>=2020] = "AFTER" 

data = data %>%
  select (Time,Cash_Holding,Leverage,Firm_Size,Growth,NetWorkingCapital,Covid19)
view(data)


### 3. Descriptive statistics ------------------------------------------------------

# Descriptive statistics all variable before Covid19--------------------------------
Before_Covid19 = data %>% 
  filter(Covid19 == "BEFORE") %>%
  select(Cash_Holding,Leverage,Firm_Size,Growth,NetWorkingCapital) %>%
  stat.desc() %>%
  t() %>%
  subset(select = c(min,max,median,mean,std.dev))

print(round(Before_Covid19,3))

## Detail descriptive statistics

# Descriptive statistics all variable after Covid19---------------------------------
After_Covid19 = data %>% 
  filter(Covid19 == "AFTER") %>%
  select(Cash_Holding,Leverage,Firm_Size,Growth,NetWorkingCapital) %>%
  stat.desc() %>%
  t() %>%
  subset(select = c(min,max,median,mean,std.dev))

print(round(After_Covid19,3))

# Descriptive statistics: Cash holding ---------------------------------------------
data %>% 
  group_by(Covid19) %>% 
  summarize(Min = min(Cash_Holding),
            Max = max(Cash_Holding),
            Mean = mean(Cash_Holding),
            Median = median(Cash_Holding),
            Standard_deviation = sd(Cash_Holding))

# Descriptive statistics: Leverage -------------------------------------------------
data %>% 
  group_by(Covid19) %>% 
  summarize(Min = min(Leverage),
            Max= max(Leverage),
            Mean=mean(Leverage),
            Median = median(Leverage),
            Standard_deviation = sd(Leverage))

# Descriptive statistics: Firm_Size ------------------------------------------------
data %>% 
  group_by(Covid19) %>% 
  summarize(Min = min(Firm_Size),
            Max= max(Firm_Size),
            Mean=mean(Firm_Size),
            Median = median(Firm_Size),
            Standard_deviation = sd(Firm_Size))

# Descriptive statistics: Growth ---------------------------------------------------
data %>% 
  group_by(Covid19) %>% 
  summarize(Min = min(Growth),
            Max= max(Growth),
            Mean=mean(Growth),
            Median = median(Growth),
            Standard_deviation = sd(Growth))

# Descriptive statistics: Change in Working Capital --------------------------------
data %>% 
  group_by(Covid19) %>% 
  summarize(Min = min(NetWorkingCapital),
            Max= max(NetWorkingCapital),
            Mean=mean(NetWorkingCapital),
            Median = median(NetWorkingCapital),
            Standard_deviation = sd(NetWorkingCapital))

### 4. The box & whisker plot and histogram of Cash holding-------------------------

# Box & whisker plot----------------------------------------------------------------
# For entire 
data %>% ggplot(aes(y = Cash_Holding)) + 
  labs(title = "Boxplot of Cash holding") +
  geom_boxplot(color = 'sandybrown')

# For 2 periods
data %>% ggplot(aes(x=Covid19, y = Cash_Holding,fill = Covid19)) + 
  geom_boxplot()

# Histogram plot---------------------------------------------------------------------
ggplot(data, aes(x= Cash_Holding, fill=Covid19)) + 
  labs(title = "Histogram plot of Cash holding ") +
  geom_histogram(binwidth=0.02, color = 'palegreen4') 

### 5. Perform multiple regression--------------------------------------------------

# 5.1. With the usual individual variables------------------------------------------
df = data %>% select(Cash_Holding,Leverage,Firm_Size,Growth,NetWorkingCapital)
cor(df)

model0 = lm(Cash_Holding~Leverage+Firm_Size+Growth+NetWorkingCapital, data=data)
summary(model0)

model1 = lm(Cash_Holding~Leverage+Firm_Size, data=data)
summary(model1)

# 5.2. With the interaction between Covid-19 dummy variable-------------------------

# Create dummy variable

data$Covid = 0
data$Covid[as.numeric(format(data$Time, format="%Y"))<2020] = 0
data$Covid[as.numeric(format(data$Time, format="%Y"))>=2020] = 1

data = data %>%
  mutate(Leverage_Covid = Leverage*Covid) %>% 
  mutate(Firm_Size_Covid = Firm_Size*Covid) %>% 
view(data)

# Perform multiple regression with dummy variable

model2 = lm(Cash_Holding~Leverage+Firm_Size
            +Leverage_Covid+Firm_Size_Covid, data=data)
summary(model2)

# 5.3 Predict for all the quarters of the sample using Model 1-----------------------

# Make predict
predict_model = predict.lm(model1)

# Create result_predicted data frame 
result_predicted = data.frame(matrix(ncol = 3, nrow = 61)) 
colnames(result_predicted) = c("Time",'Actual', 'Predicted')
result_predicted$Time = data$Time
result_predicted$Actual = data$Cash_Holding
result_predicted$Predicted = predict_model

# Convert factor to numeric
result_predicted$Actual = as.numeric(result_predicted$Actual)
result_predicted$Predicted = as.numeric(result_predicted$Predicted)
result_predicted

# RMSE 
rmse(result_predicted$Actual,result_predicted$Predicted)

# Plot predicted values
ggplot(result_predicted,                                     
       aes(x = Predicted,
           y = Actual)) +
  geom_point() +
  geom_abline(intercept = 0,
              slope = 1,
              color = "red",
              size = 2)



# 6. Perform ARIMA model to predict Cash holding------------------------------------
as.Date(as.yearqtr(data$Time, format = "Q%q/%y"), frac = 1) # Convert quarters to date
tsData = ts(data$Cash_Holding, start = c(2007, 1), frequency = 4) #Set timeseries data

# Find out the components of this time series
components.ts = decompose(tsData)
plot(components.ts)

# Check stationary 
par(mfrow=c(1,1))
acf(tsData,main='ACF for Cash holding')
pacf(tsData,main='PACF for Cash holding')
adf.test(tsData)

timeseriesseasonallyadjusted <- tsData- components.ts$seasonal
tsstationary <- diff(timeseriesseasonallyadjusted)

# Now check whether Cash_Holding is stationary
acf(tsstationary)
pacf(tsstationary)
adf.test(tsstationary)

# Use auto.arima function to determine best P, D, Q
auto = auto.arima(tsstationary,trace = T,max.order=4,
                  ic='aic')
summary(auto)

# Ljung-Box test
coeftest(auto.arima(tsstationary,seasonal=F))
acf(auto$residuals)
pacf(auto$residuals)
Box.test(auto$residuals,lag=20,type='Ljung-Box')

# Prediction
term= 4
predict(auto,h=term)
fcastauto=forecast(auto,h=term)
fcastauto
plot(fcastauto)
