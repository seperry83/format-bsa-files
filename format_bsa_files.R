# BSA Data Format
# extracting data from Excel workbook(s) and export as .csv files
# output is a folder named after the workbook populated by the indvidual .csv files
# questions: sarah.perry@water.ca.gov

#~~~~~~~~~~~~~~~~~~~~~~~~
#~~~Variables to Edit~~~#
#~~~~~~~~~~~~~~~~~~~~~~~~
# name of excel workbook(s)
filePath = 'C:/R/format-bsa-files/' # make sure final slash is included
excelFile = c('DataTemplateForR.xlsx') # keep parentheses; add multiple excel files if wanted

# choose output path
outputPath = filePath # will be created if it doesn't exist

#~~~~~~~~~~~~~~~~~~~~~~
#~~~CODE STARTS HERE~~~#
#~~~~~~~~~~~~~~~~~~~~~~
# import packages
library(readxl)
library(janitor)
library(tidyverse)
source('functions/import_df_funcs.R')
source('functions/import_meta_funcs.R')

# Extract the Data
# create full file
for (wkbk in excelFile) {
  excelBook <- paste(filePath,wkbk, sep='') # full path
  excelName <- strsplit(wkbk, '.', fixed=TRUE)[[1]][1] # remove extension from name
  
  # grab all the sheet names in a map
  sheetMap <- excel_sheets(excelBook) 
  
  # iterate over all the sheets
  for (sheetName in sheetMap) { 
    # ----
    # Intial Imports
    # ----
    # import df
    dfAbund <- import_sheet(wkbk, sheetName)
    
    # populate the category col
    dfAbund <- pop_cat_col(dfAbund) 
    
    # ----
    # Extract the Metadata
    # ----
    # add in date/time
    dfAbund <- add_date_time(wkbk, dfAbund) 
    
    # add station name
    dfAbund <- add_station_name(wkbk, dfAbund)
    
    # add v1 and v2
    dfAbund <- add_vols(wkbk, dfAbund)
    
    # add sub1 and sub2
    dfAbund <- add_subs(wkbk, dfAbund)

    # ----
    # Export Dataframe 
    # ----
    # reorganize columns
    dfAbund <- dfAbund[,c('sample','date','time','category','taxon','count','subsample','v1_ml','sub1_ml','v2_ml','sub2_ml')]
    
    # export CSV files into own folder
    fullPath <- paste(outputPath,excelName,'/', sep = '')
    fileName <- paste(fullPath,sampName,'_',dateDate,'.csv', sep = '')
    
    # create directory if it doesn't exist
    dir.create(file.path(fullPath), showWarnings = FALSE)
    
    # write CSV
    write.csv(dfAbund,fileName,row.names=F)
  }
}

print('Done! :)')