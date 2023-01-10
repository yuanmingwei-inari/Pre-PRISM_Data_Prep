##Script Author: Mingwei Yuan
##Time: 2/1/2022

install.packages("openxlsx")
install.packages("tidyverse")
install.packages("tidyxl")
library(openxlsx)
library(tidyverse)
library(tidyxl)

cols.mzh.gt <- c('Local Range','Local Row','STNDCT_PLOT_EARLY', 'GTDeadC_+7-10','GTInjureC_+7-10',
                 'GTDeadC_+18-21','GTInjureC_+18-21','STNDCT_PLOT_LATE','PHT_Est','GTStuntC_L',
                 'LRTLC','STLC', 'PLOTWT','PCTHOH','TWT','FLAG_CRO','CMNTS')
cols.sb.dpf <- c('Local Range','Local Row','STNDCT_METER','FC','PLLODG','DATE_R8','PLOTWT','PCTHOH','FLAG_CRO','CMNTS')
cols.sb.dpa <- c('Local Range','Local Row','STNDCT_METER','PLLODG','PLOTWT','PCTHOH','FLAG_CRO','CMNTS')

cols.mzh.dmpop <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','ERDRP','PLOTWT','PCTHOH','TWT','FLAG_CRO','CMNTS')

cols.mzh.dpf <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','ERTLC','GSPCE','P50_D','S50_D','PHT1','PHT2','PHT3','PHT4',
                  'EHT1','EHT2','EHT3','EHT4','LRTLC','STLC',
                  'ERDRP','PLOTWT','PCTHOH','TWT','FLAG_CRO','CMNTS')
cols.mzh.dpa <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','ERDRP','PLOTWT','PCTHOH','TWT','FLAG_CRO','CMNTS')
cols.mzh.ss <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','LRTLC','STLC','PLOTWT','PCTHOH','TWT','FLAG_CRO','CMNTS')
cols.mzh.sl <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','PLOTWT','PCTHOH_SL','FLAG_CRO','CMNTS')
cols.mzi.d1 <- c('Local Range','Local Row','STNDCT_PLOT_EARLY','P10_D','P50_D','P90_D','S10_D','S50_D','S90_D',
                 'FLAG_CRO','CMNTS')
cols.mzi.d2 <- c('Local Range','Local Row','ES','STNDCT_PLOT_EARLY','ERTLC','GSPCE','P10_D','P50_D','P90_D','S10_D',
                 'S50_D','S90_D','TPS','TSZ','PHT1', 'PHT2', 'PHT3', 'PHT4','EHT1', 'EHT2', 'EHT3', 'EHT4','STNDCT_PLOT_LATE',
                 'LRTLC','STLC','ML','FLAG_CRO','CMNTS')
cols.sb.ss <- c('Local Range','Local Row','EPS','PLLODG','DATE_R8','PLOTWT','PCTHOH','FLAG_CRO','CMNTS')
cols.sb.edf = c('Local Range','Local Row','SV','GS_17DAP','Date_17DAP','STNDCT_METER','GS_42DAP',	'Date_42DAP',	'DATE_R1',	'GS_67DAP',	'Date_67DAP',	'GS_98DAP',	'Date_98DAP','PLLODG','DATE_R8',
                'STNDCT_METER_LATE','PHT1','PHT2','PHT3','PHT4','PHT5','SHATT','PLOTWT','PCTHOH','FLAG_CRO','CMNTS')
cols.sb.eda = c('Local Range','Local Row','SV','STNDCT_METER','PLLODG','SHATT','PLOTWT','PCTHOH','FLAG_CRO','CMNTS')


###CRO DT functions
mzh.gt <- function(prot){
  #prot <- "22NA_MZH_GT"
  ###create measobs based on protocol
  cols <- cols.mzh.gt
  nc <- length(cols)
  nr <- nrow(measobs.mzh)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ##mzh.dpf
  #'STNDCT_PLOT_EARLY','GTDeadC_+7-10','GTInjureC_+7-10','GTDeadC_+18-21','GTInjureC_+18-21','STNDCT_PLOT_LATE','GTStuntC_L','LRTLC','STLC','PLOTWT': decimal, 0-200
  #'PHT_Est':decimal: 0-200
  #PCTHOH': decimal 0-100
  #TWT: decimal 0-80
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 8:9, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 10:20, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 21, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 22, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs", col = 23, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 28, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 31, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 34, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_GT_Example.xlsx", overwrite = TRUE)
}

sb.dp <- function(prot){
  #prot <- "22NA_SB_DP"
  ###create measobs based on protocol
  if (prot.t == "SB_DPa"){
    cols = cols.sb.dpa
  } else{
    cols = cols.sb.dpf
  }
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  #list_options[3,3] <- "2.3"
  #list_options[10,8] <- "2.3"
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  
  
  ##STNDCT_METER: decimal, 0:80
  #FC: white, purple, mixed
  ##PLLODG: whole, 1:9
  #DATE_R8: date
  #PLOTWT:decimal, 0:200
  #PCTHOH: decimal 0-100
  #FLAG_CRO: Include/Exclude
  if (prot.t == "SB_DPf"){
    dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                   operator = "greaterThanOrEqual", value = c(0))
    dataValidation(wb, "MeasObs",
                   col = 8, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 80))
    dataValidation(wb, "MeasObs", col = 9, rows = 2:50000, type = "list", 
                   value = "'list_options'!$I$2:$I$4")
    dataValidation(wb, "MeasObs",
                   col = 10, rows = 2:50000, type = "whole",
                   operator = "between", value = c(1,9))
    dataValidation(wb, "MeasObs",
                   col = 11, rows = 2:50000, type = "date",
                   operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
    dataValidation(wb, "MeasObs",
                   col = 12, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 200))
    dataValidation(wb, "MeasObs",
                   col = 13, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 100))
    dataValidation(wb, "MeasObs", col = 14, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  } else{
    dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                   operator = "greaterThanOrEqual", value = c(0))
    dataValidation(wb, "MeasObs",
                   col = 8, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 80))
    dataValidation(wb, "MeasObs",
                   col = 9, rows = 2:50000, type = "whole",
                   operator = "between", value = c(1,9))
    dataValidation(wb, "MeasObs",
                   col = 10, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 200))
    dataValidation(wb, "MeasObs",
                   col = 11, rows = 2:50000, type = "decimal",
                   operator = "between", value = c(0, 100))
    dataValidation(wb, "MeasObs", col = 12, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  }
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB_DP_Example.xlsx", overwrite = TRUE)
}

mzh.dmpop <- function(prot){
  prot <- "22NA_MZH_DMPop"
  ###create measobs based on protocol
  cols <- cols.mzh.dmpop
  nc <- length(cols)
  nr <- nrow(measobs.mzh)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  
  
  ##mzh.dpf
  #'STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','ERDRP','PLOTWT': decimal, 0-200
  #PCTHOH': decimal 0-100
  #TWT: decimal 0-80
  #FLAG_CRO: Include/Exclude
  dataValidation(wb, "MeasObs", col = 8:9, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 10:16, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 17, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 18, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs", col = 19, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:11, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH-DMPop_Example.xlsx", overwrite = TRUE)
}

mzh.dpf <- function(prot){
  prot <- "22NA_MZH_DPf"
  ###create measobs based on protocol
  cols <- cols.mzh.dpf
  nc <- length(cols)
  nr <- nrow(measobs.mzh)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ###DV for MeasObs
  #'STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','ERDRP','PLOTWT': decimal, 0-200
  #'P50_D','S50_D': date
  #'PHT1','PHT2','PHT3','PHT4': decimal: 0-180
  #'EHT1','EHT2','EHT3','EHT4':decimal: 0-80
  #PCTHOH': decimal 0-100
  #TWT: decimal 0-70
  #FLAG_CRO: Include/Exclude
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8:10, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 11:12, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 13:16, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 17:20, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 21:24, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 25, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 26, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs", col = 27, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_DPf_Example.xlsx", overwrite = TRUE)
}

mzh.dpa <- function(prot){
  prot <- "22NA_MZH_DPa"
  ###create measobs based on protocol
  cols <- cols.mzh.dpa
  nc <- length(cols)
  nr <- nrow(measobs.mzh)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  
  
  ##mzh.dpf
  #'STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','ERDRP','PLOTWT': decimal, 0-200
  #PCTHOH': decimal 0-100
  #TWT: decimal 0-80
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8:14, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 15, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 16, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs", col = 17, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH-DPa_Example.xlsx", overwrite = TRUE)
}

mzh.ss <- function(prot){
  prot <- "22NA_MZH_SS"
  ###create measobs based on protocol
  cols <- cols.mzh.ss
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ###Add data validation
  ##mzh.ss
  #'STNDCT_PLOT_EARLY','LRTLC','STLC','PLOTWT': decimal, 0-200: decimal, 0-200
  #PCTHOH': decimal 0-100
  #TWT: decimal 0-80
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8:11, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 12, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 13, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs", col = 14, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "~/Documents/21NA_MZH_SS_Test.xlsx", overwrite = TRUE)
}

mzh.sl <- function(prot){
  prot <- "22NA_MZH_SL"
  ###create measobs based on protocol
  cols <- cols.mzh.sl
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ###Add data validation
  #'STNDCT_PLOT_EARLY','ERTLC','GSPCE','LRTLC','STLC','PLOTWT': decimal, 0-200
  #PCTHOH_SL': decimal 0-100
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8:13, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 14, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs", col = 15, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "~/Documents/21NA_MZH_SS_Test.xlsx", overwrite = TRUE)
}

mzi.d1 <- function(prot){
  prot <- "22NA_MZI_ObsD1"
  ###create measobs based on protocol
  cols <- cols.mzi.d1
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ###Add data validation
  ##mzi
  #'STNDCT_PLOT_EARLY': decimal, 0-200
  #'P10_D','P50_D','P90_D','S10_D','S50_D','S90_D': date
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 9:14, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs", col = 15, rows = 2:50000, 
                 type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZI_Example.xlsx", overwrite = TRUE)
  
}

mzi.d2 <- function(prot){
  prot <- "22NA_MZI_ObsD2"
  ###create measobs based on protocol
  cols <- cols.mzi.d2
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ###Add data validation
  ##mzi
  #'ES','TPS','TSZ': whole, 1-9
  #'STNDCT_PLOT_EARLY','ERTLC','GSPCE': decimal, 0-200
  #'P10_D','P50_D','P90_D','S10_D','S50_D','S90_D': date
  #'PHT1','PHT2','PHT3','PHT4': decimal: 0-200
  #'EHT1','EHT2','EHT3','EHT4':decimal: 0-100
  #'STNDCT_PLOT_LATEE','LRTLC','STLC': decimal, 0-200
  #'ML': whole, 1-11
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 9:11, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 12:17, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2021-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 18:19, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 20:23, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 24:27, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs",
                 col = 28:30, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 31, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,11))
  dataValidation(wb, "MeasObs", col = 32, rows = 2:50000, type = "list", 
                 value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZI_Example.xlsx", overwrite = TRUE)
  
}

sb.ss <- function(prot){
  prot <- "22NA_SB_SS"
  ###create measobs based on protocol
  cols <- cols.sb.ss
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
  #list_options[3,3] <- "2.3"
  #list_options[10,8] <- "2.3"
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  ##EPS,PLLODG: whole, 1:9
  #DATE_R8: date
  #PLOTWT:decimal, 0:200
  #PCTHOH: decimal 0-100
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8:9, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 10, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 11, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 12, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs", col = 13, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB_SS_Example.xlsx", overwrite = TRUE)
}

sb.edf <- function(prot){
  prot <- "22NA_SB_EDf"
  ###create measobs based on protocol
  cols <- cols.sb.edf
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
 # list_options[3,3] <- "2.3"
 # list_options[10,8] <- "2.3"
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  
  ##STNDCT_METER, STNDCT_METER_LATE: decimal, 0:80
  ##SV, PLLODG: whole, 1:9
  #DATE_R8: date
  #PLOTWT:decimal, 0:200
  #PCTHOH: decimal 0-100
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 10, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 11, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs",
                 col = 13:14, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 16, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 18, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 19, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 20, rows = 2:50000, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-03-01"))
  dataValidation(wb, "MeasObs",
                 col = 21, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs",
                 col = 22:26, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 27, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 28, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 29, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs", col = 30, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))

    return(wb)
  #saveWorkbook(wb, "21NA_SB_Ed_Example.xlsx", overwrite = TRUE)
}

sb.eda <- function(prot){
  prot <- "22NA_SB_EDa"
  ###create measobs based on protocol
  cols <- cols.sb.eda
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/MgtTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  hist_mgt <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  current_mgt <- readxl::read_excel(mgt_fn,sheet = 4, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 5, guess_max = 1000000)
 # list_options[3,3] <- "2.3"
 # list_options[10,8] <- "2.3"
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Historical MgtPrac")
  writeDataTable(wb, "Historical MgtPrac", x = hist_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "Current MgtPrac")
  writeDataTable(wb, "Current MgtPrac", x = current_mgt, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  
  
  ##STNDCT_METER, STNDCT_METER_LATE: decimal, 0:80
  ##SV, PLLODG: whole, 1:9
  #DATE_R8: date
  #PLOTWT:decimal, 0:200
  #PCTHOH: decimal 0-100
  #FLAG_CRO: Include/Exclude
  
  dataValidation(wb, "MeasObs", col = 6:7, rows = 2:50000, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "MeasObs",
                 col = 8, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 9, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 80))
  dataValidation(wb, "MeasObs",
                 col = 10:11, rows = 2:50000, type = "whole",
                 operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 12, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 200))
  dataValidation(wb, "MeasObs",
                 col = 13, rows = 2:50000, type = "decimal",
                 operator = "between", value = c(0, 100))
  dataValidation(wb, "MeasObs", col = 14, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$B$2:$B$7")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$C$2:$C$4")
  
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 8:12, type = "decimal",
                 operator = "between", value = c(0, 300000))
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 13:15, type = "date",
                 operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  ##DV for Hist Mgt
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 3:7, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 9:13, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Historical MgtPrac", col = 2, rows = 15:19, type = "decimal", 
                 operator = "greaterThanOrEqual", value = c(0))
  
  ##DV for Current Mgt
  ##DV for Current Mgt
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 2, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 4, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 6, type = "list", value = "'list_options'!$E$2:$E$6")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 8, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 11, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 14, type = "list", value = "'list_options'!$F$2:$F$10")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 18, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 21, type = "list", value = "'list_options'!$G$2:$G$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 25, type = "list", value = "'list_options'!$H$2:$H$5")
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 28, type = "list", value = "'list_options'!$H$2:$H$5")
  
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 3, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 5, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 7, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 10, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 54, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 58, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 61, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 63, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 65, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 67, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 69, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 71, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 73, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  dataValidation(wb, "Current MgtPrac",col = 2, rows = 75, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-10-01"))
  
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 60, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 62, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 64, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 66, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 68, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 70, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 72, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 74, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  dataValidation(wb, "Current MgtPrac", col = 2, rows = 76, type = "decimal",operator = "greaterThanOrEqual", value = c(0))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB_Ed_Example.xlsx", overwrite = TRUE)
}

##ResAg DT functions
cols.ag.mzh.dp <- c('ES','SV','ANTNOSR','BLS','CR','EYSPT','GLS','GW','NCLB','SR','TrSpt','StyGrn','FAp','KRN1','KRN2','KRN3',
                    'Cob_COL1','Cob_COL2','Cob_COL3','EarAp','FUSER','DIPER','GIBER','Qual_K','AdvRc_Ag','Flag_Ag','CMNTS')
cols.ag.mzh.sl = c('ES','SV','ANTNOSR','BLS','CR','EYSPT','GLS','GW','NCLB','SR','TrSpt','Flag_Ag','CMNTS')
cols.ag.mzi = c('Flag_Ag','CMNTS')
#FLX and Flag_Ag mandatory
cols.ag.mzh.dmpop <- c('FLX','Flag_Ag','CMNTS') 
cols.ag.mzh.gt <- c('Flag_Ag','CMNTS')

cols.ag.sb.ss <- c('BRSTRT','FEYLP','IDC','PHYTPH','PYTHM','RHIZTO','SDS','WHITMO','Flag_Ag','CMNTS')
cols.ag.sb.ed <- c('EPS','BRSTRT','FEYLP','IDC','PHYTPH','PYTHM','RHIZTO','SDS','WHITMO','Flag_Ag','CMNTS')
cols.ag.sb.dp = c('EPS','BRSTRT','FEYLP','IDC','PHYTPH','PYTHM','RHIZTO','SDS', 'WHITMO','Flag_Ag','CMNTS')



ag.mzi <- function(prot){
  prot <- "21NA_MZI"
  ###create measobs based on protocol
  cols <- cols.ag.mzi
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/21NA_Protocols/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:40, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ###DV for MeasObs
  dataValidation(wb, "MeasObs",
                 col = 17:28, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  #dataValidation(wb, "MeasObs", col = 29, rows = 2:50000, type = "list", value = "'list_options'!$C$2:$C$3")
  dataValidation(wb, "MeasObs",
                 col = 29, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 30:32, rows = 2:50000, type = "decimal", operator = "between", value = c(0,30))
  dataValidation(wb, "MeasObs", col = 33:35, rows = 2:50000, type = "list", value = "'list_options'!$D$2:$D$5")
  dataValidation(wb, "MeasObs",
                 col = 36, rows = 2:50000, type = "decimal", operator = "between", value = c(0,200))
  dataValidation(wb, "MeasObs", col = 37, rows = 2:50000, type = "list", value = "'list_options'!$E$2:$E$5")
  dataValidation(wb, "MeasObs", col = 38, rows = 2:50000, type = "list", value = "'list_options'!$C$2:$C$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$A$2:$A$2")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$B$2:$B$8")
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 9:10, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11:12, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 14, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 18, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 22, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 24, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 26, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 28, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 32, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 36, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 40, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 44, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 48, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZI_ResAg.Example.xlsx", overwrite = TRUE)
}

ag.mzh.dp <- function(prot){
  prot <- "22NA_MZH_DP"
  ###create measobs based on protocol
  cols <- cols.ag.mzh.dp
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:44, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ###DV for MeasObs
  dataValidation(wb, "MeasObs",
                 col = 17:29, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs",
                 col = 30:32, rows = 2:50000, type = "decimal", operator = "between", value = c(0,30))
  dataValidation(wb, "MeasObs", col = 33:35, rows = 2:50000, type = "list", value = "'list_options'!$B$2:$B$5")
  dataValidation(wb, "MeasObs",
                 col = 36:41, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 42, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_DPf_ResAg.Example.xlsx", overwrite = TRUE)
}
ag.mzh.sl <- function(prot){
  prot <- "22NA_MZH_SL"
  ###create measobs based on protocol
  cols <- cols.ag.mzh.sl
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:44, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ###DV for MeasObs
  dataValidation(wb, "MeasObs",
                 col = 17:27, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 28, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_DPf_ResAg.Example.xlsx", overwrite = TRUE)
}


ag.mzh.dmpop <- function(prot){
  prot <- "22NA_MZH_DMPop"
  ###create measobs based on protocol
  cols <- cols.ag.mzh.dmpop
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:39, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ###DV for MeasObs
  dataValidation(wb, "MeasObs",
                 col = 17, rows = 2:50000, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 18, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_DMPop_ResAg.Example.xlsx", overwrite = TRUE)
}
ag.mzh.gt <- function(prot){
  prot <- "22NA_MZH_GT"
  ###create measobs based on protocol
  cols <- cols.ag.mzh.gt
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:39, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ###DV for MeasObs
  dataValidation(wb, "MeasObs", col = 17, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  ##DV for In-season meta
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_DMPop_ResAg.Example.xlsx", overwrite = TRUE)
}

ag.mzh.ss <- function(prot){
  prot <- "22NA_MZH_SS"
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  
  ##DV for In-season meta
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_MZH_SS_ResAg.Example.xlsx", overwrite = TRUE)
}

ag.sb.dp <- function(prot){
  prot <- "22NA_SB_DP"
  ###create measobs based on protocol
  cols <- cols.ag.sb.dp
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  list_options[3,2] <- "2.3"
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:25, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ##DV for MeasObs  
  dataValidation(wb, "MeasObs",
                 col = 17:25, rows = 2:50000, type = "whole",operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 26, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  
  
  ##DV for Current Mgt
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB-DP_ResAg.Example.xlsx", overwrite = TRUE)
}
ag.sb.dm <- function(prot){
  prot <- "21NA_SB_DM"
  ###create measobs based on protocol
  cols <- cols.ag.sb.dm
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/21NA_Protocols/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  list_options[3,2] <- "2.3"
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:21, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ##DV for MeasObs  
  dataValidation(wb, "MeasObs",
                 col = 19, rows = 2:50000, type = "whole",operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 20, rows = 2:50000, type = "list", value = "'list_options'!$C$2:$C$3")
  
  ##DV for Current Mgt
  dataValidation(wb, "In-season metadata", col = 2, rows = 2, type = "list", value = "'list_options'!$A$2:$A$2")
  dataValidation(wb, "In-season metadata", col = 2, rows = 3, type = "list", value = "'list_options'!$B$2:$B$14")
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 9:10, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11:12, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 14, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 16, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 18, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 20, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 22, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 24, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 26, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 28, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 30, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 32, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 34, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 36, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 38, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 40, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 42, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 44, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 46, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 48, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 50, type = "date",operator = "greaterThanOrEqual", value = as.Date("2021-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB-DM_ResAg.Example.xlsx", overwrite = TRUE)
}

ag.sb.ed <- function(prot){
  prot <- "22NA_SB_Ed"
  ###create measobs based on protocol
  cols <- cols.ag.sb.ed
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  list_options[3,2] <- "2.3"
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:25, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ##DV for MeasObs  
  dataValidation(wb, "MeasObs",
                 col = 17:25, rows = 2:50000, type = "whole",operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 26, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  
  
  ##DV for Current Mgt
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB-DP_ResAg.Example.xlsx", overwrite = TRUE)
}
ag.sb.ss <- function(prot){
  prot <- "22NA_SB_SS"
  ###create measobs based on protocol
  cols <- cols.ag.sb.ss
  nc <- length(cols)
  nr <- nrow(measobs)
  traits <- data.frame(matrix(ncol = nc, nrow = nr))
  colnames(traits) <- cols
  
  measobs.f <- cbind(measobs, traits)
  
  ###read in mgt sheets
  mgt_fn <- paste0("~/Documents/22_Analyses/Data_Templates/AgTemplates/",prot, ".xlsx")
  readme <- readxl::read_excel(mgt_fn,sheet = 1, guess_max = 1000000)
  meta <- readxl::read_excel(mgt_fn,sheet = 2, guess_max = 1000000)
  list_options <- readxl::read_excel(mgt_fn,sheet = 3, guess_max = 1000000)
  list_options[3,2] <- "2.3"
  
  ###
  style <- createStyle(
    fontColour = "black", bgFill = "yellow"
  )
  
  ####
  wb <- createWorkbook()
  addWorksheet(wb, "Readme_Metadata definitions")
  writeDataTable(wb, "Readme_Metadata definitions", x = readme, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "Readme_Metadata definitions", style, rows = 1, cols = 1:6, gridExpand = TRUE)
  
  addWorksheet(wb, "In-season metadata")
  writeDataTable(wb, "In-season metadata", x = meta, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "In-season metadata", style, rows = 1, cols = 1:2, gridExpand = TRUE)
  
  addWorksheet(wb, "MeasObs")
  writeDataTable(wb, "MeasObs", x = measobs.f, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "MeasObs", style, rows = 1, cols = 1:21, gridExpand = TRUE)
  
  addWorksheet(wb, "list_options")
  writeDataTable(wb, "list_options",  x = list_options, startRow = 1, startCol = 1, rowNames = FALSE)
  addStyle(wb, "list_options", style, rows = 1, cols = 1:5, gridExpand = TRUE)
  
  ##DV for MeasObs  
  dataValidation(wb, "MeasObs",
                 col = 17:24, rows = 2:50000, type = "whole",operator = "between", value = c(1,9))
  dataValidation(wb, "MeasObs", col = 25, rows = 2:50000, type = "list", value = "'list_options'!$A$2:$A$3")
  
  
  
  ##DV for Current Mgt
  dataValidation(wb, "In-season metadata",
                 col = 2, rows = 6:7, type = "whole", operator = "between", value = c(1,9))
  dataValidation(wb, "In-season metadata",col = 2, rows = 8:9, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 11, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 13, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 15, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 17, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 19, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 21, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 23, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 25, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 27, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 29, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 31, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 33, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 35, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 37, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 39, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 41, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 43, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 45, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  dataValidation(wb, "In-season metadata",col = 2, rows = 47, type = "date",operator = "greaterThanOrEqual", value = as.Date("2022-04-01"))
  
  return(wb)
  #saveWorkbook(wb, "21NA_SB-SS_ResAg.Example.xlsx", overwrite = TRUE)
}
