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
    
    # add project name
    dfAbund <- add_proj_name(wkbk, sheetName, dfAbund)
    
    # add station name
    dfAbund <- add_station_name(wkbk, sheetName, dfAbund)
    
    # add in date/time
    dfAbund <- add_date_time(wkbk, sheetName, dfAbund) 
    
    # add sample number
    dfAbund <- add_samp_num(wkbk, sheetName, dfAbund) 
    
    # add tow
    dfAbund <- add_tow(wkbk, sheetName, dfAbund)
    
    # add numer of jars
    dfAbund <- add_jar_num(wkbk, sheetName, dfAbund)
    
    # add numer of vials
    dfAbund <- add_vial_num(wkbk, sheetName, dfAbund)
    
    # add BSA metadata
    dfAbund <- add_bsa_meta(wkbk, sheetName, dfAbund)
    
    # add v1, v2, and vsed
    dfAbund <- add_vols(wkbk, sheetName, dfAbund)
    
    # add sub1 and sub2
    dfAbund <- add_subs(wkbk, sheetName, dfAbund)
    
    # ----
    # Export Dataframe 
    # ----
    # reorganize columns
    dfAbund <- dfAbund[,c('project','station','date','time','samp_num','tow',
                          'id_by','scope','mag','id_date','num_of_jars','num_of_vials',
                          'category','taxon','count','subsample','v1_ml','sub1_ml','v2_ml','sub2_ml','vsed_ml')]
    
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