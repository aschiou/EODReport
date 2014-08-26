Consolidate<-function(filename=character()) {
  require(xlsx)
  rpttype <- c('')
  rowIndex <- numeric(0)
  if (grepl(" Con ", filename)) {
    rpttype<-c("Con")
    rowIndex <- c(7,8,11,12)
  }
  else if (grepl(" Lab ", filename)) {
    rpttype<-c("Lab")
    rowIndex <- c(9,10,13,14,21,22,25,26)
  }
  else if (grepl(" Rad ", filename)) {
    rpttype<-c("Rad")
    rowIndex <- c(7,8,11,12)
  }

  rptdate <- read.xlsx(filename, sheetIndex=2, colIndex=3, rowIndex=1, header=F, colClasses='Date')
  
  totals <- read.xlsx(filename, sheetIndex=2, colIndex=3, rowIndex=rowIndex, header=F, colClasses="numeric")

  df<-data.frame(rpttype=as.character(rpttype), rptdate=as.Date(rptdate[1,1]), 
                 va_ord_tot=as.numeric(totals[1,1]), va_ord_pas=as.numeric(totals[2,1]), 
                 dod_ord_tot=as.numeric(totals[3,1]),dod_ord_pas=as.numeric(totals[4,1]),
                 va_res_tot=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[5,1]), 
                 va_res_pas=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[6,1]), 
                 dod_res_tot=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[7,1]), 
                 dod_res_pas=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[8,1]))
  # filename="EOD Con Report 20140818.xlsx"
  df$rpttype <- as.character(df$rpttype)

  df
}