## =====================================================================
## FORECASTING FINAL PROJECT
## Minh PHAN - May 2016
## =====================================================================
rm(list=ls(all=T))
dev.off()
par(mfrow=c(1,1))

## =====================================================================
## Project description
## =====================================================================

## The purpose of this final project is to apply all forecasting methods
## that have been studied in the course, compare the results and select
## the best one.

## The datasets for the project were created from real data, as follow:
## [1] Monthly Crude Oil Spot Price FOB (01/1986-04/2016)
## [2] Monthly Avg Temperature for Washington Area, DC (01/1900-05/2016)

# install.packages("fpp")
# install.packages("readxl")
library(fpp)
library(readxl)

setwd("C:/Users/Administrator/Desktop/Forecasting/Final project")



## =====================================================================
## PART I - Master function to test all ETS models and Auto-ARIMA
## =====================================================================
## The purpose of this part is to created automatic functions that
## provide a framwork to choose the best ETS or Auto ARIMA model.
##
## The list of functions is as follow:
## (1) ETS.all(...) : scan all possible ETS methods, return a list of 
##                    trained ETS models
## (2) model.all(...) : combine ETS.all(...) and Auto ARIMA, return a
##                      list of trained models
## (3) forecast.all(...) : apply the whole list of models to forecast 
##                         data, return a list of forecasting objects
## (4) accuracy.fit(...) : evaluate the accuracy of models on training
##                         datasets, return a data.frame of RMSE, MAE,
##                         MAPE, MASE values for each forecast model
## (5) accuracy.fc(...) : evaluate the accurary of models on testing
##                        datasets, return a data.frame of RMSE, MAE,
##                        MAPE, MASE values for each forecast model
## (6) acc.rank(...) : rank the accurary tables to provide a guide to
##                     select the best model
## =====================================================================

## ---------------------------------------------------------------------
## Function to train all possible ETS models
## ---------------------------------------------------------------------

ETS.all <- function(data.ts,error,trend,season) {
  
  ## List of all useless ETS models
  model.bad <- c("MMA","ANM","AAM","AMN","AMA","AMM")
  
  ## List of ETS model, will be returned in the end of function
  model.list <- list()
  model.name <- c()
  
  for (e in error) # Scan through all possible error options
    for (t in trend) # Scan through all possible trend options
      for (s in season) { # Scan through all possible season options
        
        model <- paste0(e,t,s) # Form a model name
        if (model %in% model.bad) next # If model in bad models list, skip the loop

        fit <- ets(data.ts,model=model)
        model.name <- c(model.name,model)
        model.list[[length(model.list)+1]] <- fit
        
        if (t!="N") { # Check for damped model
          fit.damped <- ets(data.ts,model=model,damped=T)
          model.name <- c(model.name,paste(model,"+ Damped"))
          model.list[[length(model.list)+1]] <- fit.damped
        }
      }
  
  return(list(model.name,model.list))
}

## ---------------------------------------------------------------------
## Function to combine all ETS models with Auto ARIMA model
## ---------------------------------------------------------------------

model.all <- function(data.ts,error,trend,season) {
  
  ## Create all possible ETS models
  model.list <- ETS.all(data.ts,error,trend,season)
  
  ## Auto ARIMA model
  fit <- auto.arima(data.ts)
  model.list[[1]] <- c(model.list[[1]],"Auto ARIMA")
  model.list[[2]][[length(model.list[[2]])+1]] <- fit
  
  return(model.list)
}

## ---------------------------------------------------------------------
## Function to test accuracy of all models (i.e. ETS and Auto ARIMA)
## Return a table with RMSE, MAE, MAPE, MASE values of all models
## ---------------------------------------------------------------------

accuracy.fit <- function(model.list) {
  
  acc <- data.frame()
  for (i in 1:length(model.list[[2]]))
    acc <- rbind(acc,accuracy(model.list[[2]][[i]]))
  
  acc <- acc[,c(2,3,5,6)] # RMSE, MAE, MAPE, MASE
  rownames(acc) <- model.list[[1]]
  # colnames(acc) <- c("RMSE", "MAE", "MAPE", "MASE")
  return(acc)
}

## ---------------------------------------------------------------------
## Function to create forecasting list
## ---------------------------------------------------------------------

forecast.all <- function(model.list,h) {
  
  forecast.list <- list()
  for (i in 1:length(model.list[[2]])) {
    fc <- forecast(model.list[[2]][[i]],h=h)
    forecast.list[[length(forecast.list)+1]] <- fc
  }
  
  return(list(model.list[[1]],forecast.list))
}

## ---------------------------------------------------------------------
## Function to evaluation accurary of forecasting list
## ---------------------------------------------------------------------

accuracy.fc <- function(data.ts,forecast.list) {
  
  acc <- data.frame()
  for (i in 1:length(forecast.list[[2]]))
    acc <- rbind(acc,accuracy(forecast.list[[2]][[i]],data.ts)[2,])
  
  acc <- acc[,c(2,3,5,6)] # RMSE, MAE, MAPE, MASE
  rownames(acc) <- forecast.list[[1]]
  colnames(acc) <- c("RMSE", "MAE", "MAPE", "MASE")
  return(acc)
}

## ---------------------------------------------------------------------
## Function to rank the accurary results
## ---------------------------------------------------------------------

acc.rank <- function(acc.table) {
  r <- data.frame(sapply(acc.table,rank,ties.method=c("min")))
  r <- transform(r, Sum=rowSums(r))
  
  TotalRank <- rank(r$Sum,ties.method=c("min"))
  r <- cbind(r,TotalRank)
  r$Sum <- NULL
  
  rownames(r) <- rownames(acc.table)
  return(r)
}



## =====================================================================
## Part II - Analyse Dataset 1
## =====================================================================
## Data: Monthly Crude Oil Spot Price FOB (USD/Barrel)
## Period: 01/1986-04/2016
## =====================================================================

## ---------------------------------------------------------------------
## Read, clean and plot data
## ---------------------------------------------------------------------

## Input data from Excel file
crudeOil <- read_excel("datasets.xlsx",sheet=4)

## Convert to time series data type
oil.ts <- ts(crudeOil$`Cushing, OK WTI Spot Price FOB (Dollars per Barrel)`,
             start=c(1986,1),frequency=12)

## Plot the whole period
plot(oil.ts,type="l",ylab="Price (USD/Barrel)",xlab="Year",
     main="Monthly Crude Oil Spot Price FOB \n Jan 1986 - Apr 2016")

## STL decomposition data
plot(stl(oil.ts,s.window="periodic"))

## Seasonal plot by month
seasonplot(oil.ts,ylab="Price (USD/Barrel)",xlab="Month", 
           main="Monthly Crude Oil Spot Price FOB \n Jan 1986 - Apr 2016",
           year.labels=T,year.labels.left=T,col=1:200,pch=19)

## ---------------------------------------------------------------------
## Selecting forecast models
## ---------------------------------------------------------------------

## Train and test datasets
oil.train <- window(oil.ts,start=c(1986,1),end=c(2009,12))
oil.test <- window(oil.ts,start=c(2010,1),end=c(2016,4))
oil.h <- length(oil.test) # Period [76]

## Training all possible ETS models and Auto ARIMA
## Put them in a list of model
oil.model.list <- model.all(data=oil.train,
                            error=c("A","M"),
                            trend=c("N","A","M"), # Full options for trend
                            season=c("N","A","M")) # Full options for season

## Forecasting all above models, put the results in a list
oil.forecast.list <- forecast.all(oil.model.list,h=oil.h)

## Make some fancy plotting: all models
plot(oil.ts,type="l",ylab="Price (USD/Barrel)",xlab="Year", 
     main="Monthly Crude Oil Spot Price FOB \n Jan 1986 - Apr 2016")

for (i in 1:length(oil.forecast.list[[2]]))
  lines(oil.forecast.list[[2]][[i]]$mean,col=i,lwd=2.5)

legend("topleft",oil.forecast.list[[1]],lty=c(1,1),lwd=c(2.5,2.5),
       col=c(1:length(oil.forecast.list[[2]])),cex=0.8,ncol=2,
       x.intersp=0.5,y.intersp=0.5,text.width=5)

## Evaluating the accurary of all models
oil.fit.acc <- accuracy.fit(oil.model.list)
oil.fit.rank <- acc.rank(oil.fit.acc)
oil.fit.acc
oil.fit.rank

## Evaluating the accurary based on testing datasets
oil.fc.acc <- accuracy.fc(oil.test,oil.forecast.list)
oil.fc.rank <- acc.rank(oil.fc.acc)
oil.fc.acc
oil.fc.rank
# Top best models:
# - In term of RMSE: AAA and AAA + Damped
# - In term of MASE: MMM and MMM + Damped

## Make some fancy plotting: selected models
plot(oil.ts,type="l",ylab="Price (USD/Barrel)",xlab="Year", 
     main="Monthly Crude Oil Spot Price FOB \n Jan 1986 - Apr 2016")

oil.top <- c(5,6,18,19) # Selected models
for (i in oil.top)
  lines(oil.forecast.list[[2]][[i]]$mean,col=i,lwd=2.5)

legend("topleft",oil.forecast.list[[1]][oil.top],lty=c(1,1),lwd=c(2.5,2.5),
       col=oil.top,cex=0.8,ncol=1,
       x.intersp=0.5,y.intersp=0.5,text.width=5)

## ---------------------------------------------------------------------
## Final forecast models: MMM + Damped (best in RMSE)
## Forecast: Oil prices of May-Dec 2016
## ---------------------------------------------------------------------

oil.fit <- ets(oil.ts,model="MMM",damped=T)
summary(oil.fit)
oil.fc <- forecast(oil.fit,h=8)
oil.fc

plot(oil.fc,ylab="Price (USD/Barrel)",xlab="Year", 
     main="Forecast Crude Oil Spot Price FOB \n for May 2016 - Dec 2016 periods")

## ---------------------------------------------------------------------
## Some improvements: log transformation
## ---------------------------------------------------------------------

plot(log(oil.ts),type="l",col="blue",ylab="Log(Price (USD/Barrel))",xlab="Year", 
     main="Monthly Crude Oil Spot Price FOB \n Jan 1986 - Apr 2016")

par(new=T)
plot(oil.ts,type="l",col="red",axes=F,xlab="",ylab="")

legend("topleft",c("Log(Oil Price)","Oil Price"),lty=c(1,1),lwd=c(2.5,2.5),
       col=c("blue","red"),cex=0.8,ncol=1,
       x.intersp=0.5,y.intersp=0.5,text.width=5)



## =====================================================================
## Part III - Analyse Dataset 2
## =====================================================================
## Data: Monthly Avg Temperature for Washington Area, DC (F) 
## Period: 01/1900-05/2016
## =====================================================================

## ---------------------------------------------------------------------
## Read, clean and plot data
## ---------------------------------------------------------------------

## Input data from Excel file
wdcTemp <- read_excel("datasets.xlsx",sheet=2,na="M")
wdcTemp <- wdcTemp[,-c(1,14)] # Drop unused columns

## Convert to time series data type
temp <- na.omit(c(t(as.matrix(wdcTemp)))) # Transpose and convert matrix to vector
temp.ts <- ts(temp,start=1900,frequency=12)

## Plot the whole period
plot(temp.ts,type="l",ylab="Temperature (F)",xlab="Year",
     main="Monthly Avg Temperature for Washington Area, DC \n Jan 1900 - May 2016")

## Plot by month
seasonplot(temp.ts,ylab="Temperature (F)",xlab="Month", 
           main="Monthly Avg Temperature for Washington Area, DC \n Jan 1900 - May 2016",
           year.labels=T,year.labels.left=T,col=1:200,pch=19)

## STL decomposition data
plot(stl(temp.ts,s.window="periodic"))

## ---------------------------------------------------------------------
## Selecting forecast models
## ---------------------------------------------------------------------

## Train and test datasets
temp.train <- window(temp.ts,start=c(1900,1),end=c(1989,12))
temp.test <- window(temp.ts,start=c(1990,1),end=c(2016,5))
temp.h <- length(temp.test) # Period [317]

## Training all possible ETS models and Auto ARIMA
## Put them in a list of model
temp.model.list <- model.all(data=temp.train,
                             error=c("A","M"),
                             trend=c("N","A","M"),
                             season=c("N","A","M"))

## Forecasting all above models, put the results in a list
temp.forecast.list <- forecast.all(temp.model.list,h=temp.h)

## Make some fancy plotting: all models
plot(temp.ts,type="l",ylab="Temperature (F)",xlab="Year", 
     main="Monthly Avg Temperature for Washington Area, DC \n Jan 1900 - May 2016")

for (i in 1:length(temp.forecast.list[[2]]))
  lines(temp.forecast.list[[2]][[i]]$mean,col=i,lwd=2.5)

legend("topleft",temp.forecast.list[[1]],lty=c(1,1),lwd=c(2.5,2.5),
       col=c(1:length(temp.forecast.list[[2]])),cex=0.8,ncol=2,
       x.intersp=0.5,y.intersp=0.5,text.width=15)

## Evaluating the accurary of all models
temp.fit.acc <- accuracy.fit(temp.model.list)
temp.fit.rank <- acc.rank(temp.fit.acc)
temp.fit.acc
temp.fit.rank

## Evaluating the accurary based on testing datasets
temp.fc.acc <- accuracy.fc(temp.test,temp.forecast.list)
temp.fc.rank <- acc.rank(temp.fc.acc)
temp.fc.acc
temp.fc.rank
# Top best models: AAA and AAA + Damped

## Make some fancy plotting: selected models
plot(temp.ts,type="b",ylab="Temperature (F)",xlab="Year", 
     main="Monthly Avg Temperature for Washington Area, DC \n Jan 1900 - May 2016")

temp.top <- c(5,6) # Selected models
for (i in temp.top)
  lines(temp.forecast.list[[2]][[i]]$mean,col=i,lwd=2.5)

legend("topleft",temp.forecast.list[[1]][temp.top],lty=c(1,1),lwd=c(2.5,2.5),
       col=temp.top,cex=0.8,ncol=1,
       x.intersp=0.5,y.intersp=0.5,text.width=15)

## ---------------------------------------------------------------------
## Final forecast models: AAA + Damped (best in all measures)
## Forecast: Washington DC Avg Temperature (F) of Jun-Dec 2016
## ---------------------------------------------------------------------

temp.fit <- ets(temp.ts,model="AAA",damped=T)
summary(temp.fit)
temp.fc <- forecast(temp.fit,h=7)
temp.fc

plot(temp.fc,ylab="Temperature (F)",xlab="Year", 
     main="Forecast Monthly Avg Temperature for Washington Area, DC \n June 2016 - Dec 2016")

## =====================================================================
## Last edited on 31 May 2016
## =====================================================================