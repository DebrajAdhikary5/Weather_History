# Install and load necessary packages
install.packages('openxlsx')
library(openxlsx)
install.packages('plyr')
library(plyr)
install.packages('party')
library(party)
install.packages('caTools')
library(caTools)

# Read the weather data
weatherHistory = read.xlsx('E:\\OneDrive\\IVY\\Stats+R\\Materials\\Assigments\\Multiple Linear Regression Assignment\\weatherHistory.xlsx', na.strings = c("", " ", "NA", "NULL", "null"))

# Basic Exploration of the data
sum(is.na(weatherHistory))
head(weatherHistory)
str(weatherHistory)
summary(weatherHistory)

# Extract day, month, year, hour from the date
weatherHistory$Day = as.numeric(sapply(weatherHistory$Formatted.Date, substring, 9, 10))
weatherHistory$Month = as.numeric(sapply(weatherHistory$Formatted.Date, substring, 6, 7))
weatherHistory$Year = as.numeric(sapply(weatherHistory$Formatted.Date, substring, 1, 4))
weatherHistory$Hour = as.numeric(sapply(weatherHistory$Formatted.Date, substring, 12, 13))

# Drop unused columns
weatherHistory = weatherHistory[, !(colnames(weatherHistory) %in% c("Formatted.Date", "Loud.Cover", "Daily.Summary", "Apparent.Temperature..C."))]

# Handling missing values
weatherHistory$Precip.Type[is.na(weatherHistory$Precip.Type)] <- 'rain'

# Visualizations
library(RColorBrewer)

# Histograms for continuous variables
conHist = c("Temperature.(C)", "Humidity", "Wind.Speed.(km/h)", "Wind.Bearing.(degrees)", "Visibility.(km)")
par(mfrow = c(2, 3))
for (contCol in conHist) {
  hist(weatherHistory[, contCol], main = paste('Histogram of:', contCol), col = brewer.pal(8, "Paired"))
}

# Bar plots for categorical variables
catBar = c("Summary", "Precip.Type")
par(mfrow = c(1, 2))
for (catCol in catBar) {
  barplot(table(weatherHistory[, catCol]), main = paste('Barplot of:', catCol), col = brewer.pal(8, "Paired"))
}

# Correlation analysis
ContinuousCols = c("Temperature.(C)", "Humidity", "Wind.Speed.(km/h)", "Wind.Bearing.(degrees)", "Visibility.(km)")
cor(weatherHistory[, ContinuousCols], use = "complete.obs")

# Continuous Vs Categorical correlation strength: ANOVA
ColsForANOVA = c("Summary", "Precip.Type")
for (catCol in ColsForANOVA) {
  anovaData = weatherHistory[, c("Temperature.(C)", catCol)]
  print(str(anovaData))
  print(summary(aov(`Temperature.(C)` ~ ., data = anovaData)))
}

# Map categorical variables to numerical values
weatherHistory$Precip.Type <- mapvalues(weatherHistory$Precip.Type, from = c("rain", "snow"), to = c(1, 2))
weatherHistory$Summary <- mapvalues(weatherHistory$Summary, from = as.vector(unique(weatherHistory$Summary)), to = seq(1, length(as.vector(unique(weatherHistory$Summary))), 1))

# Generate the data frame for machine learning
InputData = weatherHistory
TargetVariableName <- c('Temperature.(C)')
BestPredictorName <- c("Humidity", "Wind.Speed.(km/h)", "Wind.Bearing.(degrees)", "Visibility.(km)", "Precip.Type", "Summary", "Day", "Month", "Year", "Hour")

# Extracting target and predictor variables
TargetVariable <- InputData[, TargetVariableName]
PredictorVariable <- InputData[, BestPredictorName]
DataForML <- data.frame(TargetVariable, PredictorVariable)
DataForML$Summary <- as.numeric(DataForML$Summary)
DataForML$Precip.Type <- as.numeric(DataForML$Precip.Type)

# Sampling | Splitting data into 70% for training 30% for testing
set.seed(123)
TrainingSampleIndex <- sample(1:nrow(DataForML), size = 0.7 * nrow(DataForML))
DataForMLTrain <- DataForML[TrainingSampleIndex, ]
DataForMLTest <- DataForML[-TrainingSampleIndex, ]

# Linear Regression Model
startTime <- Sys.time()
Model_Reg <- lm(TargetVariable ~ ., data = DataForMLTrain)
summary(Model_Reg)
hist(Model_Reg$residuals)
endTime <- Sys.time()
endTime - startTime

# Checking Accuracy of model on Testing data
DataForMLTest$Pred_LM <- predict(Model_Reg, DataForMLTest)
LM_APE <- 100 * (abs(DataForMLTest$Pred_LM - DataForMLTest$TargetVariable) / DataForMLTest$TargetVariable)
DataForMLTest$APE <- LM_APE
print(paste('### Mean Accuracy of Linear Regression Model is: ', 100 - mean(LM_APE)))
print(paste('### Median Accuracy of Linear Regression Model is: ', 100 - median(LM_APE)))

# Decision Tree Model
startTime <- Sys.time()
Model_CTREE <- ctree(TargetVariable ~ ., data = DataForMLTrain, controls = ctree_control(maxdepth = 5))
plot(Model_CTREE)
endTime <- Sys.time()
endTime - startTime

# Checking Accuracy of model on Testing data
DataForMLTest$Pred_CTREE <- as.numeric(predict(Model_CTREE, DataForMLTest))
CTREE_APE <- 100 * (abs(DataForMLTest$Pred_CTREE - DataForMLTest$TargetVariable) / DataForMLTest$TargetVariable)
print(paste('### Mean Accuracy of ctree Model is: ', 100 - mean(CTREE_APE)))
print(paste('### Median Accuracy of ctree Model is: ', 100 - median(CTREE_APE)))
