install.packages("openxlsx")
install.packages("tidyverse")
install.packages("tidyxl")
library(openxlsx)
library(tidyverse)
library(tidyxl)
library(stringr)

source("~/Documents/RCodes/22DataTemplatesFunctions.R")

###Loc Info: Mapping Bridge Sites + Protocol
bridge <- readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=1, guess_max = 1000000)
colnames(bridge)[5] <- "prot"
bridge$`Book Name` <- paste0(bridge$CROShort,"_", bridge$Loc, "_", bridge$State)
bridge$Protocol <- paste0("22NA_",bridge$prot, "_", bridge$EntrySet)
#bridge.mz.ss <- as.vector(unlist(bridge[bridge$prot == "MZH_SS", 'Book Name']))
#bridge.mz.dp <- as.vector(unlist(bridge[bridge$prot %in% c("MZH_DPf","MZH_DPa"), 'Book Name']))

bridge.sb.dp <- as.vector(unlist(bridge[bridge$prot %in% c("SB_DPa","SB_DPf"), 'Book Name']))
#bridge.sb.ss <- as.vector(unlist(bridge[bridge$prot == "SB_SS", 'Book Name']))

mapping.sb <- readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=2, guess_max = 1000000)
mapping.sb$`Book Name` <- paste0(mapping.sb$CROShort,"_", mapping.sb$Loc, "_", mapping.sb$State)
colnames(mapping.sb)[5] = "prot"
mapping.sb$Protocol = paste0("22NA_",mapping.sb$prot, "_", mapping.sb$EntrySet)

mapping.mz <- readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=4, guess_max = 1000000)
mapping.mz$`Book Name` <- paste0(mapping.mz$CROShort,"_", mapping.mz$Loc, "_", mapping.mz$State)
colnames(mapping.mz)[5] = "prot"
mapping.mz$Protocol = paste0("22NA_",mapping.mz$prot, "_", mapping.mz$EntrySet)




##MZH_GT
ms.gt = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22_GT.xlsx")
mz.gt_list <-  split(ms.gt, ms.gt$`Book Name`)
mz.gt_list.f <-  mz.gt_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(mz.gt_list.f)){
  #i=2
  print(i)
  measobs.mzh = mz.gt_list.f[[i]]
  prot <- "22NA_MZH_GT"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'FldTrtApp','RecId','Row','Range','RepNo')
  
  wb = mzh.gt(prot)
  
  bookname <- names(mz.gt_list.f)[i]
  filename <- paste0(bookname, "_22NA_MZH_GT_M+L")
  
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZH_GT/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_DMPop
ms.dm = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22_DMPop.xlsx")
mz.dm_list <-  split(ms.dm, ms.dm$`Book Name`)

for (i in 1:length(mz.dm_list)){
  #i=2
  print(i)
  measobs.mzh = mz.dm_list[[i]]
  prot <- "22NA_MZH_DMPop"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'FldTrtApp','RecId','Row','Range','RepNo')
  
  wb = mzh.dmpop(prot)
  
  bookname <- names(mz.dm_list)[i]
  filename <- paste0(bookname, "_22NA_MZH_DMPop_M+L")
  
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZH_DMPop/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

###SB templates
##SB_DPf and SB_DPa
ms.sb = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22_SB.xlsx")
ms.sb$prot.t <- sapply(strsplit(as.character(ms.sb$`Entry Book Name`),'_'), "[", 3)
ms.sb.dp <- ms.sb %>% filter(ms.sb$prot.t == "DP")

ms.sb.dp.f = ms.sb.dp %>% left_join(mapping.sb[,c(5,8:9)]) %>% select(-c("prot.t"))

#for (i in bridge.sb.dp){
#  #bridge.t <- bridge[bridge$prot == "SB_DP",]
#  protocol <- bridge[bridge$`Book Name` == i, "Protocol"][[1]]
#  ms.sb.dp[ms.sb.dp$`Book Name` == i, "Protocol"] <- protocol
#}



sb.dp_list <-  split(ms.sb.dp.f, list(ms.sb.dp.f$`Book Name`, ms.sb.dp.f$Protocol))
sb.dp_list.f <-  sb.dp_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(sb.dp_list.f)){
  #i=58
  print(i)
  prot <- "22NA_SB_DP"
  measobs.mzh <- sb.dp_list.f[[i]]
  prot.t = measobs.mzh$prot[1]
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb <- sb.dp(prot)
  
  filename <- names(sb.dp_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/SB_DP/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_DPa/f
ms = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_3.24.22.xlsx")
ms.mz = ms %>% filter(`Book Project` == "IC")
ms.mz$prot.t <- sapply(strsplit(as.character(ms.mz$`Entry Book Name`),'_'), "[", 3)
ms.mz.dp <- ms.mz %>% filter(prot.t %in% c("D1","D1.1","D1.2","D2", "PC"))

ms.mz.dp.f = ms.mz.dp %>% left_join(mapping.mz[,c(5,8:9)]) %>% select(-c("prot.t"))

mz.dp_list <-  split(ms.mz.dp.f, list(ms.mz.dp.f$`Book Name`, ms.mz.dp.f$Protocol))
mz.dp_list.f <-  mz.dp_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(mz.dp_list.f)){
  #i=58
  print(i)
  measobs.mzh <- mz.dp_list.f[[i]]
  prot.t = measobs.mzh$prot[1]
  prot = paste0("22NA_",prot.t)
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  if (prot == "22NA_MZH_DPa"){
    wb = mzh.dpa(prot)
  } else{
    wb = mzh.dpf(prot)
  }
  
  filename <- names(mz.dp_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZH_DP/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_SL
ms.mz.sl <- ms.mz %>% filter(prot.t == "SL")
bridge.sl <- readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=5, guess_max = 1000000)
colnames(bridge.sl)[5] <- "prot"
bridge.sl$`Book Name` <- paste0(bridge.sl$CROShort,"_", bridge.sl$Loc, "_", bridge.sl$State)
bridge.sl$Protocol <- paste0("22NA_",bridge.sl$prot, "_", bridge.sl$EntrySet)
bridge.sl.list <- as.vector(bridge.sl$`Book Name`)

for (i in bridge.sl.list){
  #bridge.t <- bridge[bridge$prot == "SB_DP",]
  protocol <- bridge.sl[bridge.sl$`Book Name` == i, "Protocol"][[1]]
  ms.mz.sl[ms.mz.sl$`Book Name` == i, "Protocol"] <- protocol
}

ms.mz.sl$Protocol = ms.mz.sl$`Entry Book Name`
mz.sl_list <-  split(ms.mz.sl, list(ms.mz.sl$`Book Name`, ms.mz.sl$Protocol))
mz.sl_list.f <-  mz.sl_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(mz.sl_list.f)){
  #i=2
  print(i)
  measobs.mzh = mz.sl_list.f[[i]]
  prot <- "22NA_MZH_SL"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb = mzh.sl(prot)
  
  filename <- names(mz.sl_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZH_SL/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

###MZH_SS
ms.mz$prot.t <- sapply(strsplit(as.character(ms.mz$`Entry Book Name`),'_'), "[", 3)
mz.ss = ms.mz %>% filter(!`Entry Book Name` == "Filler") %>%filter(!prot.t %in% c("D1","D1.1","D1.2","D2","DMPop","GT","ObsD1","ObsD2+","PC","SL"))
mz.ss$mg <- substr(mz.ss$`Entry Book Name`,8,8) ## mg: maturity group
mz.ss$Protocol = paste0("22NA_MZH_SS_",mz.ss$mg)

bridge.ss <- readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=6, guess_max = 1000000)
colnames(bridge.ss)[5] <- "prot"
bridge.ss$`Book Name` <- paste0(bridge.ss$CROShort,"_", bridge.ss$Loc, "_", bridge.ss$State)
bridge.ss$Protocol <- paste0("22NA_",bridge.ss$prot, "_", bridge.ss$EntrySet)
bridge.ss.list <- as.vector(bridge.ss$`Book Name`)

for (i in bridge.ss.list){
  protocol <- bridge.ss[bridge.ss$`Book Name` == i, "Protocol"][[1]]
  mz.ss[mz.ss$`Book Name` == i, "Protocol"] <- protocol
}

mz.ss_list <-  split(mz.ss, list(mz.ss$`Book Name`, mz.ss$Protocol))
mz.ss_list.f <-  mz.ss_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(mz.ss_list.f)){
  #i=2
  print(i)
  measobs.mzh = mz.ss_list.f[[i]]
  prot <- "22NA_MZH_SS"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb = mzh.ss(prot)
  
  filename <- names(mz.ss_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZH_SS/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZI_ObdD1/D2+
mzi.obsd1 = ms.mz %>% filter(!`Entry Book Name` == "Filler") %>%filter(prot.t == "ObsD1")
mzi.obsd2 = ms.mz %>% filter(!`Entry Book Name` == "Filler") %>%filter(prot.t == "ObsD2+")

mzi.obsd1_list <-  split(mzi.obsd1, list(mzi.obsd1$`Book Name`))
mzi.obsd1_list.f <-  mzi.obsd1_list %>% discard(function(x) nrow(x) == 0)

mzi.obsd2_list <-  split(mzi.obsd2, list(mzi.obsd2$`Book Name`))
mzi.obsd2_list.f <-  mzi.obsd2_list %>% discard(function(x) nrow(x) == 0)

for (i in 1:length(mzi.obsd1_list.f)){
 # i=1
  print(i)
  measobs.mzh = mzi.obsd1_list.f[[i]]
  prot <- "22NA_MZI_ObsD1"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb = mzi.d1(prot)
  
  filename <- names(mzi.obsd1_list.f)[i]
  filename_ <- paste0(filename,"_22NA_MZI_ObsD1_EML")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZI_ObsD1/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

for (i in 1:length(mzi.obsd2_list.f)){
   #i=1
  print(i)
  measobs.mzh = mzi.obsd2_list.f[[i]]
  prot <- "22NA_MZI_ObsD2"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb = mzi.d2(prot)
  
  filename <- names(mzi.obsd2_list.f)[i]
  filename_ <- paste0(filename,"_22NA_MZI_ObsD2+_EML")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/MZI_ObsD2/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

###SB_SS
sb = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22NA_SB_4.28.22.xlsx")
sb.ed = sb %>% filter(`Entry Book Name` %in% c('22NA_SB_EDSY1-2_2.3','22NA_SB_EDSY1-2_3.3', '22NA_SB_EDSY1-3_3.3' ))
sb.ss = sb %>% filter(!(`Entry Book Name` %in% c('22NA_SB_EDDMLPD', '22NA_SB_EDDMSales', '22NA_SB_EDSY1-2_2.3',
                                                 '22NA_SB_EDSY1-2_3.3', '22NA_SB_EDSY1-3_3.3' ,'22NA_SB_EDTI',
                                                 '22NA_SB_DP_0.5','22NA_SB_DP_1.5','22NA_SB_DP_2.3','22NA_SB_DP_2.7','22NA_SB_DP_3.3',
                                                 '22NA_SB_DP_3.7','22NA_SB_DP_4.5')))
mapping.sbed = readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=8, guess_max = 1000000)
mapping.sbss = readxl::read_excel("~/Documents/22_Analyses/Data_templates/22NA_SiteInfo.xlsx", sheet=7, guess_max = 1000000)
mapping.sbss[c(3:4,14,19),5] = '2.3'
mapping.sbed[c(1:3,6:7,11,14,23,25:26),5] = '2.3'

mapping.sbss$`Book Name` <- paste0(mapping.sbss$CROShort,"_", mapping.sbss$Loc, "_", mapping.sbss$State)
colnames(mapping.sbss)[4] = "prot"
mapping.sbss$Protocol = paste0("22NA_",mapping.sbss$prot, "_", mapping.sbss$EntrySet)

mapping.sbed$`Book Name` <- paste0(mapping.sbed$CROShort,"_", mapping.sbed$Loc, "_", mapping.sbed$State)
colnames(mapping.sbed)[4] = "prot"
mapping.sbed$Protocol = paste0("22NA_",mapping.sbed$prot, "_", mapping.sbed$EntrySet)

ms.sb.ss.f = sb.ss %>% left_join(mapping.sbss[,c(4,6,7)]) 
sb.ss_list <-  split(ms.sb.ss.f, list(ms.sb.ss.f$`Book Name`, ms.sb.ss.f$Protocol))
sb.ss_list.f <-  sb.ss_list %>% discard(function(x) nrow(x) == 0)

ms.sb.ed.f = sb.ed %>% left_join(mapping.sbed[,c(4,6,7)]) 
sb.ed_list <-  split(ms.sb.ed.f, list(ms.sb.ed.f$`Book Name`, ms.sb.ed.f$Protocol))
sb.ed_list.f <-  sb.ed_list %>% discard(function(x) nrow(x) == 0)

for (i in 1:length(sb.ss_list.f)){
 # i=2
  print(i)
  measobs.mzh = sb.ss_list.f[[i]]
  prot <- "22NA_SB_SS"
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  wb = sb.ss(prot)
  
  filename <- names(sb.ss_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/SB_SS/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

for (i in 1:length(sb.ed_list.f)){
  # i=7
  print(i)
  measobs.mzh = sb.ed_list.f[[i]]
  prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% select('Book Name', 'Entry Book Name', 'RecId','Row','Range')
  
  if (prot == "SB_EDa"){
    wb = sb.eda(prot)
  } else{
    wb = sb.edf(prot)
  }
  
  
  filename <- names(sb.ed_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_DataTemplates/SB_Ed/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

#####################################
####RA templates
##MZH_DPa+DPf
ms = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_3.24.22.xlsx")
ms.mz = ms %>% filter(`Book Project` == "IC")
ms.mz$prot.t <- sapply(strsplit(as.character(ms.mz$`Entry Book Name`),'_'), "[", 3)
ms.mz.dp <- ms.mz %>% filter(prot.t %in% c("D1","D1.1","D1.2","D2", "PC"))

ms.mz.dp.f = ms.mz.dp %>% left_join(mapping.mz[,c(5,8:9)]) %>% select(-c("prot.t"))

mz.dp_list <-  split(ms.mz.dp.f, list(ms.mz.dp.f$`Book Name`, ms.mz.dp.f$Protocol))
mz.dp_list.f <-  mz.dp_list %>% discard(function(x) nrow(x) == 0)

for (i in 1:length(mz.dp_list.f)){
  #i=9
  measobs.mzh <- mz.dp_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.dp(prot)
  
  
  filename <- names(mz.dp_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZH_DP"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##SB_DP
for (i in 1:length(sb.dp_list.f)){
  print(i)
  #i=9
  measobs.mzh <- sb.dp_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.sb.dp(prot)
  
  
  filename <- names(sb.dp_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/SB_DP"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_SL
for (i in 1:length(mz.sl_list.f)){
  print(i)
  #i=9
  measobs.mzh <- mz.sl_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.sl(prot)
  
  
  filename <- names(mz.sl_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZH_SL"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##SB_DP
for (i in 1:length(sb.ed_list.f)){
  print(i)
  #i=9
  measobs.mzh <- sb.dp_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.sb.dp(prot)
  
  
  filename <- names(sb.dp_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/SB_DP"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##SB_Ed
for (i in 1:length(sb.ed_list.f)){
  print(i)
  #i=9
  measobs.mzh <- sb.ed_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.sb.dp(prot)
  
  
  filename <- names(sb.ed_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/SB_ED"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_GT
ms.gt = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22_GT.xlsx")
mz.gt_list <-  split(ms.gt, ms.gt$`Book Name`)
mz.gt_list.f <-  mz.gt_list %>% discard(function(x) nrow(x) == 0)
for (i in 1:length(mz.gt_list.f)){
  print(i)
  #i=1
  measobs.mzh <- mz.gt_list.f[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.gt(prot)
  
  bookname <- names(mz.gt_list.f)[i]
  filename <- paste0(bookname, "_22NA_MZH_GT_M+L")
  
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZH_GT/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}


##MZH_DMPop
ms.dm = readxl::read_excel("~/Documents/22_Analyses/Data_Templates/Yield Trial Master Catalog_22_DMPop.xlsx")
mz.dm_list <-  split(ms.dm, ms.dm$`Book Name`)
for (i in 1:length(mz.dm_list)){
  print(i)
  #i=1
  measobs.mzh <- mz.dm_list[[i]]
  #Prot <- measobs.mzh$prot[1]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.dmpop(prot)
  
  bookname <- names(mz.dm_list)[i]
  filename <- paste0(bookname, "_22NA_MZH_DMPop_M+L")
  
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZH_DMPop/",cro)
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZH_SS
for (i in 1:length(mz.ss_list.f)){
  print(i)
  #i=1
  wb = ag.mzh.ss(prot)
  
  filename <- names(mz.ss_list.f)[i]
  filename_ <- str_replace(filename, "\\.", "_")
  
  dirname <- "~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZH_SS/"
  dir.create(dirname)
  path <- paste0(dirname,"/", filename_, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##MZI
for (i in 1:length(mzi.obsd1_list.f)){
  # i=1
  print(i)
  measobs.mzh = mzi.obsd1_list.f[[i]]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.gt(prot)
  
  bookname <- names(mzi.obsd1_list.f)[i]
  filename <- paste0(bookname, "_22NA_MZI_ObsD1_EML")
  
  
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZI/")
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

for (i in 1:length(mzi.obsd2_list.f)){
  #i=1
  print(i)
  measobs.mzh = mzi.obsd2_list.f[[i]]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.mzh.gt(prot)
  
  bookname <- names(mzi.obsd2_list.f)[i]
  filename <- paste0(bookname, "_22NA_MZI_ObsD2+_EML")
  
  cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/MZI/")
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

##SB_SS
for (i in 1:length(sb.ss_list.f)){
  #i=1
  print(i)
  measobs.mzh = sb.ss_list.f[[i]]
  measobs = measobs.mzh %>% dplyr::select('Book Name','Entry Book Name','RecId','Set Code','Range','Row','Local Range',
                                          'Local Row','Exp Stage','GEName','Experimental Name','Commercial Name',
                                          'Maturity','Plot Discarded','FldTrtApp','RepNo')
  wb = ag.sb.ss(prot)
  
  bookname <- names(sb.ss_list.f)[i]
  filename <- str_replace(bookname, "\\.", "_")
  
  #cro <- strsplit(filename,"_")[[1]][1]
  dirname <- paste0("~/Documents/22_Analyses/Data_Templates/22NA_ResAgr_DataTemplates/SB_SS/")
  dir.create(dirname)
  path <- paste0(dirname,"/", filename, ".xlsx")
  saveWorkbook(wb, path, overwrite = TRUE)
}

