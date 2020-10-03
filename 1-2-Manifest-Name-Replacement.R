# ---- Program Header ----
# Project: 
# Program Name: 
# Author: 
# Created: 
# Purpose: 
# Revision History:
# Date        Author        Revision
# 

# ---- Initialize Libraries ----
library(dplyr)
library(lubridate)
library(haven)
library(openxlsx)
library(writexl)
library(tidyverse)
library(knitr)
library(labelled)
library("WriteXLS")

# ---- Load Functions ----
source("../11.3.1  R Production Programs/Global_Functions.R")
cleaner()
source("../11.3.1  R Production Programs/Global_Functions.R")

#source("Project_Functions.R")

# ---- Runtime Parameters ----
run_time <- Sys.time()
file_name <- paste("Subject_Listing", format(run_time,"%Y%b%d-%H%M%OS"), sep="")
report_date <- paste("Report Run:", format(run_time,"%Y-%b-%d %H:%M:%OS", sep=" "))
#dbCon <- dbConnect(odbc(),"DSN_name")

# ---- Load Raw Data ----
allManifests <- readAllExcelFiles("Manifests","xls",TRUE)
listOfFileNames <- allFileNames("Manifests","xls")
# ---- Apply Data Transformations ----
#Replaces all values in the column visit name wiht the values from "CSF visit type"
for(i in 1:length(allManifests)){
  allManifests[[i]][,13]<- allManifests[[i]][,15]
}


# ---- Generate Output ----
for(i in 1:length(allManifests)){
  WriteXLS(allManifests[[i]],paste("Correct Visit Name Manifests/",listOfFileNames[i],sep = ""))
}

write.xlsx(allManifests[[1]],paste("Correct Visit Name Manifests/Corrected",listOfFileNames[1]))
write
WriteXLS()
#Arrange columns for final output


#Export to Excel


#Export R Object


#Generate DTD


# ---- Send Notifications ----
#Email output to recipients


# ---- Clean Environment ----
rm(list = ls())
gc()
