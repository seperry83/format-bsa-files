# ----
# add datetime cols
# ----
add_date_time <- function(wkbk, dfAbund) {
  dfDate <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'F6:M6'
      ,col_types = 'numeric'
      ,col_names = F
    )
  )
  
  # collapse date to string
  dateStr <- unlist(dfDate, use.names = FALSE) %>% paste(.,collapse = '')
  
  # convert to datetime and add to df
  dateDate <- as.Date(dateStr, '%m%d%Y')
  dfAbund$date <- dateDate
  
  # extract time from sheet
  dfTime <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'F9'
      ,col_types = 'date'
      ,col_names = F
    )
  )
  
  # format timestamp
  timeDate <- strptime(dfTime[[1]], "%Y-%m-%d %H:%M:%S") %>% format(., '%H:%M')

  # add time to df
  dfAbund$time <- timeDate
  
  return(dfAbund)
}

# ----
# extract sample name
# ----
add_station_name <- function(wkbk, dfAbund){
  statName <- as.character(
    suppressMessages(
      read_excel(
        wkbk
        ,sheet = sheetName
        ,range = 'V6'
        ,col_names = F
      )
    )
  )
  
  # changes spaces to underscores
  if (' ' %in% statName) {
    statName <- gsub(" ", "_", statName)  
  }
  
  # add samp name to df; remove whitespace
  dfAbund$station <- trimws(statName)

  return(dfAbund)
}

# ----
# add volumes
# ----
add_vols <- function(wkbk, dfAbund){
  # import v1 df
  vOneDf <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'M9:Q9'
      ,col_types = 'numeric'
      ,col_names = F
    )
  )
  
  # grab value/append to df
  if (sum(vOneDf, na.rm = TRUE) == 0) {
    print(paste('Check sheet',sheetName)) # error in sheet
    dfAbund$v1_ml <- NA # append NA
  } else {
    vOne <- vOneDf[,colSums(is.na(vOneDf)) == 0][[1]] # collapse to value
    dfAbund$v1_ml <- vOne # append
  }
  
  # import v2 df
  vTwoDf <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'C12:D12'
      ,col_types = 'numeric'
      ,col_names = F
    )
  )
  
  # grab value/append to df
  if (sum(vTwoDf, na.rm = TRUE) == 0) {
    print(paste('Check sheet',sheetName)) # error in sheet
    dfAbund$v2_ml <- NA # append NA
  } else {
    vTwo <- vTwoDf[,colSums(is.na(vTwoDf)) == 0][[1]] # collapse to value
    dfAbund$v2_ml <- vTwo # append
  }
  
  return(dfAbund)
}

# ----
# add subsamples
# ----
add_subs <- function(wkbk, dfAbund){
  # import sub1 df
  subOneDf <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'U9:W9'
      ,col_types = 'numeric'
      ,col_names = F
    )
  )
  
  # grab value/append to df
  if (sum(subOneDf, na.rm = TRUE) == 0) {
    print(paste('Check sheet',sheetName)) # error in sheet
    dfAbund$sub1_ml <- NA # append NA
  } else {
    subOne <- subOneDf[,colSums(is.na(subOneDf)) == 0][[1]] # collapse to value
    dfAbund$sub1_ml <- subOne # append
  }
  
  # import sub2 df
  subTwoDf <- suppressMessages(
    read_excel(
      wkbk
      ,sheet = sheetName
      ,range = 'H12:I12'
      ,col_types = 'numeric'
      ,col_names = F
    )
  )
  
  # grab value/append to df
  if (sum(subTwoDf, na.rm = TRUE) == 0) {
    print(paste('Check sheet',sheetName)) # error in sheet
    dfAbund$sub2_ml <- NA # append NA
  } else {
    subTwo <- subTwoDf[,colSums(is.na(subTwoDf)) == 0][[1]] # collapse to value
    dfAbund$sub2_ml <- subTwo # append
  }
  
  return(dfAbund)
}