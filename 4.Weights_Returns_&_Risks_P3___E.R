#Required Packages 
library(fBasics)
library(fPortfolio)
library(openxlsx2)
library(PerformanceAnalytics)
library(readxl)
library(timeSeries)

#Data
Data <- data.frame(read_excel("~/Proyecto_3_Portafolio_Eficiente___English/6.Data_1_P3___E.xlsx"))
Data$Date <- as.Date(Data$Date)#Transforming the "Date" column of the "Data" data frame to data type "Date"

#Returns
Returns <- na.omit(Return.calculate(Data), type= "Discrete")

#Transforming Returns to Time Series
ReturnsTS <- as.timeSeries(Returns)

#Efficient Frontier
Efficient_Frontier <- portfolioFrontier(ReturnsTS, constraints= "LongOnly")

#Efficient Frontier Chart
plot(Efficient_Frontier, c(1))

#Efficient Frontier Weights
Weights_EF <- getWeights(Efficient_Frontier)

#Efficient Frontier Returns
Returns_EF <- getTargetReturn(Efficient_Frontier)

#Efficient Frontier Risks
Risks_EF <- getTargetRisk(Efficient_Frontier)

#Export Weights to Excel
write_xlsx(Weights_EF, "Weights.xlsx")

#Export Returns to Excel
write_xlsx(Returns_EF, "Returns.xlsx")

#Export Risks to Excel
write_xlsx(Risks_EF, "Risks.xlsx")