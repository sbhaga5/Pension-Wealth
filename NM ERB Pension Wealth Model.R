rm(list = ls())
library("readxl")
library(tidyverse)
setwd(getwd())
FileName <- 'Model Inputs.xlsx'
#Test
#
#
model_inputs <- read_excel(FileName, sheet = 'Main')
MeritIncreases <- model_inputs[,5]
payroll_growth <- as.double(model_inputs[which(model_inputs[,1] == 'Payroll Growth Rate (Basic)'),2])
salary_growth <- as.double(model_inputs[which(model_inputs[,1] == 'Salary Increase (Wage Inflation)'),2])
StartingSalary <- as.double(model_inputs[which(model_inputs[,1] == 'Starting Salary'),2])
Vesting <- as.double(model_inputs[which(model_inputs[,1] == 'Vesting (years)'),2])
HiringAge <- as.double(model_inputs[which(model_inputs[,1] == 'Hiring Age'),2])
YOS <- model_inputs[,4]

IRSCompLimit <- as.double(model_inputs[which(model_inputs[,1] == 'IRS Compensation Limit'),2])
ER_Contrib <- as.double(model_inputs[which(model_inputs[,1] == 'ER Contribution Rate'),2])
EE_Contrib <- as.double(model_inputs[which(model_inputs[,1] == 'EE Contribution Rate'),2])
ARR <- as.double(model_inputs[which(model_inputs[,1] == 'ARR/DR'),2])
Interest <- as.double(model_inputs[which(model_inputs[,1] == 'Credited Interest'),2])
COLA <- as.double(model_inputs[which(model_inputs[,1] == 'COLA'),2])
COLACompound <- as.double(model_inputs[which(model_inputs[,1] == 'COLA Compounds (1=Yes)'),2])
ScaleMultiple <- as.double(model_inputs[which(model_inputs[,1] == 'Mortality Rate Scale Multiple'),2])
SetBackYear <- as.double(model_inputs[which(model_inputs[,1] == 'Mortality Rate Set Back (year)'),2])
#
#
#Linear Interpolation Function. Use for later
LinearInterpolation <- function(Data){
  StartValue <- Data[1,1]
  StartIndex <- 2
  count <- 2
  while(count <= nrow(Data)){
    if(!is.na(Data[count,1])){
      EndValue <- Data[count,1]
      EndIndex <- count - 1
      
      for(i in StartIndex:EndIndex){
        Data[i,1] <- Data[i-1,1] - (StartValue - EndValue)/(EndIndex - StartIndex + 2)
      }
      
      StartValue <- Data[count,1]
      StartIndex <- count + 1
    } else if((count == nrow(Data))) {
      EndIndex <- count
      for(i in StartIndex:EndIndex){
        Data[i,1] <- Data[i-1,1] - (StartValue - EndValue)/(EndIndex - StartIndex + 2)
      }
    }
    count <- count + 1
  }
  
  return(Data)
}

InitializeDataFieldName <- function (Data, ColumnName) {
  NewData <- Data
  colnames(NewData) <- ColumnName
  return(NewData)
}
#
#Salary increases and other
#MeritIncreases <- LinearInterpolation(MeritIncreases)
Salary <- InitializeDataFieldName(MeritIncreases, 'Salary Growth')
FinalAvgSalary <- InitializeDataFieldName(MeritIncreases,'Final Average Salary Growth')
IRSSalaryCap <- InitializeDataFieldName(MeritIncreases,'Salary Growth Subject to IRS Cap')
CumulativeWage <- InitializeDataFieldName(MeritIncreases,'Cumulative Current Wage (CCW)')

Salary[1,1] <- StartingSalary
CumulativeWage[1,1] <- 0

for(i in 1:nrow(MeritIncreases)){
  if(i > 1){
    Salary[i,1] <- Salary[i-1,1]*(1 + (MeritIncreases[i-1,1] + salary_growth)) 
    CumulativeWage[i,1] <- CumulativeWage[i-1,1]*(1 + ARR) + Salary[i-1,1]
  }
  IRSSalaryCap[i,1] <- min(Salary[i,1],IRSCompLimit)
    
  FinalAvgSalary[i,1] <- 0
  if(YOS[i,1] >= Vesting){
    FinalAvgSalary[i,1] <- sum(Salary[(i-Vesting):(i-1),1])/Vesting
  }
}
#
#
#Employee Contribution
Employee_Contrib <- EE_Contrib*InitializeDataFieldName(Salary,'Employee Contribution')
Employee_Contrib_CI <- EE_Contrib*InitializeDataFieldName(Salary,'Employee Contribution (with Compounded Interest)')
PV_Employee_Contrib_CI <- InitializeDataFieldName(Employee_Contrib_CI,'PV of Employee Contribution (with Compounded Interest)')
Employer_Contrib <- ER_Contrib*InitializeDataFieldName(Salary,'Employer Contribution')

for(i in 1:nrow(PV_Employee_Contrib_CI)){
  if((YOS[i,1] > 0) && (i > 1)){
    Employee_Contrib_CI[i,1] <- (Employee_Contrib_CI[i-1,1]*(1+Interest) + Employee_Contrib[i-1,1])
    PV_Employee_Contrib_CI[i,1] <- Employee_Contrib_CI[i,1]/((1+ARR)^(as.double(YOS[i,1])))
  } else {
    Employee_Contrib_CI[i,1] <- 0
    PV_Employee_Contrib_CI[i,1] <- 0
  }
}
#
#
#Mortality Rates
#SurvivalModel <- model_inputs[which(model_inputs[,1] == 'Survival Model'),2]
MaleMortality <- read_excel(FileName, sheet = 'MP-2017_Male')
FemaleMortality <- read_excel(FileName, sheet = 'MP-2017_Female')
SurvivalRates <- read_excel(FileName, sheet = 'Mortality Rates')

Age <- SurvivalRates[,1]
Pub2010NonDisMaleTeacher <- SurvivalRates[,2]
Pub2010DisMaleTeacher <- SurvivalRates[,3]
Pub2010MaleTeacher <- SurvivalRates[,4]
Pub2010NonDisFemaleTeacher <- SurvivalRates[,5]
Pub2010DisFemaleTeacher <- SurvivalRates[,6]
Pub2010FemaleTeacher <- SurvivalRates[,7]

RP2014Male <- SurvivalRates[,8]
RP2014HealthyMale <- SurvivalRates[,9]
RP2014DisMale <- SurvivalRates[,10]
RP2014Female <- SurvivalRates[,11]
RP2014HealthyFemale <- SurvivalRates[,12]
RP2014DisFemale <- SurvivalRates[,13]

# Pub2010NonDisMaleTeacher <- LinearInterpolation(Pub2010NonDisMaleTeacher)
# Pub2010DisMaleTeacher <- LinearInterpolation(Pub2010DisMaleTeacher)
# #Pub2010MaleTeacher <- LinearInterpolation(Pub2010MaleTeacher)
# Pub2010NonDisFemaleTeacher <- LinearInterpolation(Pub2010NonDisFemaleTeacher)
# Pub2010DisFemaleTeacher <- LinearInterpolation(Pub2010DisFemaleTeacher)
# #Pub2010FemaleTeacher <- LinearInterpolation(Pub2010FemaleTeacher)

YearCol <- (2129 - 2009 + 1)
AgeRow <- (120 - 25 + 1)
SurvivalMaleMatrix <- matrix(0, nrow = AgeRow, ncol = YearCol)
SurvivalFemaleMatrix <- matrix(0, nrow = AgeRow, ncol = YearCol)

for(i in 1:96){
  SurvivalAge <- i + 24
  for(j in 2:121){
    SurvivalYear <- j + 2008
    #Column Check is for all of the special cases such as RP 2014 and Pub 2010
    ColumnCheck <- 0
    
    if((SurvivalAge == 120)){
      SurvivalMaleMatrix[i,j] <- 1
      SurvivalFemaleMatrix[i,j] <- 1
      ColumnCheck <- 1
    } else if((SurvivalAge <= 57) && (SurvivalYear == 2010)){
      SurvivalMaleMatrix[i,j] <- as.double(Pub2010MaleTeacher[which(SurvivalRates[,1] == SurvivalAge),1])
      SurvivalFemaleMatrix[i,j] <- as.double(Pub2010FemaleTeacher[which(SurvivalRates[,1] == SurvivalAge),1])
      ColumnCheck <- 1
    } else if((SurvivalAge > 57) && (SurvivalYear == 2010)){
      SurvivalMaleMatrix[i,j] <- as.double(Pub2010NonDisMaleTeacher[which(SurvivalRates[,1] == SurvivalAge),1])*ScaleMultiple
      SurvivalFemaleMatrix[i,j] <- as.double(Pub2010NonDisFemaleTeacher[which(SurvivalRates[,1] == SurvivalAge),1])
      ColumnCheck <- 1
    }
    
    #If ColumnCheck is 0, then do the normal calc
    if(ColumnCheck == 0){
      if(SurvivalYear <= 2033){
        YearRef <- SurvivalYear
      } else {
        YearRef <- 2033
      }
      
      RowIndex <- which(MaleMortality[,1] == SurvivalAge)
      ColIndex <- which(colnames(MaleMortality) == 2033)
      SurvivalMaleMatrix[i,j] <- SurvivalMaleMatrix[i,j-1]*(1 - as.double(MaleMortality[RowIndex,ColIndex]))
      
      RowIndex <- which(FemaleMortality[,1] == SurvivalAge)
      ColIndex <- which(colnames(FemaleMortality) == 2033)
      SurvivalFemaleMatrix[i,j] <- SurvivalFemaleMatrix[i,j-1]*(1 - as.double(FemaleMortality[RowIndex,ColIndex]))
    }
    
    if(SurvivalYear == 2020){
      SurvivalMaleMatrix[i,1] <- SurvivalMaleMatrix[i,12]
      SurvivalFemaleMatrix[i,1] <- SurvivalFemaleMatrix[i,12]
    }
  }
}
#
#
#RP2014 and Adjusted Mortality Rates
#RP Rates and Adjusted Mortality
RPRatesMale <- RP2014Male
RPRatesFemale <- RP2014Female
RPRatesAvg <- RPRatesMale

AdjustedMale <- RP2014Male
AdjustedFemale <- RP2014Female
AdjustedAvg <- RPRatesMale

for(i in 1:nrow(RPRatesMale)){
  if(is.na(RP2014Male[i,1])){
    RPRatesMale[i,1] <- RP2014HealthyMale[i,1]
  } else {
    RPRatesMale[i,1] <- RP2014Male[i,1]
  }
  
  if(is.na(RP2014Female[i,1])){
    RPRatesFemale[i,1] <- RP2014HealthyFemale[i,1]
  } else {
    RPRatesFemale[i,1] <- RP2014Female[i,1]
  }
  
  RPRatesAvg[i,1] <- (RPRatesMale[i,1] + RPRatesFemale[i,1]) /2
  
  if(Age[i,1] < HiringAge){
    AdjustedMale[i,1] <- SurvivalMaleMatrix[i,11]
    AdjustedFemale[i,1] <- SurvivalFemaleMatrix[i,11]
  } else {
    AdjustedMale[i,1] <- SurvivalMaleMatrix[i,11+i]
    AdjustedFemale[i,1] <- SurvivalFemaleMatrix[i,11+i]
    AdjustedAvg[i,1] <- (AdjustedMale[i,1] + AdjustedFemale[i,1]) / 2
  }
}
#
#
#Survival Probability and Annuity Factor
SurvivalRates <- AdjustedAvg
InitializeDataFieldName
ProbNextYear <- InitializeDataFieldName(SurvivalRates,'Probability of Survival to Next Year')
ProbNextYear[1,1] <- 1
DiscProbNextYear <- InitializeDataFieldName(ProbNextYear,'Discount Probability of Survival to Next Year')
LifeExpectancy <- InitializeDataFieldName(ProbNextYear,'Life Expectancy')

for(i in 1:nrow(ProbNextYear)){
  if(i == 1){
    ProbNextYear[i,1] <- (1-SurvivalRates[i,1]) * 1
  } else {
    ProbNextYear[i,1] <- (1-SurvivalRates[i-1,1]) * ProbNextYear[i-1,1]
  }
  DiscProbNextYear[i,1] <- ProbNextYear[i,1]/(1+ARR)^(i-1)
}

for(i in 1:nrow(ProbNextYear)){
  LifeExpectancy[i,1] <- sum(ProbNextYear[i:nrow(ProbNextYear),]) / ProbNextYear[i,1]
}

AnnuityFactor <- DiscProbNextYear
colnames(AnnuityFactor) <- 'Annuity Factor'
AnnuityFactor[1:20,1] <- 1

for(j in 45:120){
  sum <- 0
  RowIndex <- j - 25 + 1
  TempValue <- 0
  for(i in RowIndex:nrow(DiscProbNextYear)){
    if((as.double(Age[i,1]) >= j) && (COLACompound == 1)){
      TempValue <- as.double(DiscProbNextYear[i,1])*((1+COLA)^(as.double(Age[i,1]) - j))
      sum <- sum + TempValue
    } else {
      TempValue <- as.double(DiscProbNextYear[i,1])*(1+(COLA*(as.double(Age[i,1]) - j)))
      sum <- sum + TempValue
    }
    
    if(as.double(Age[i,1]) == j){
      FactorDiv <- TempValue
    }
  }
  
  AnnuityFactor[RowIndex,1] <- sum / FactorDiv
}
