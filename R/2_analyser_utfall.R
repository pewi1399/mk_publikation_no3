rm(list=objects())
library(data.table)
#library(XLConnect)
library(openxlsx)
library(ggplot2)
library(dplyr)
library(reshape2)
library(magrittr)
library(tidyr)
library(parallel)

dat <- readRDS("Output/get_data.rds")
#dat <- dat[1:30000,]
source("R/functions.R")

metadata <- openxlsx::read.xlsx(file.path("Dokument", "Outcomes_160208.xlsx"), sheet = 1)
spop1vars <- unique(metadata[metadata$File == "Studpop1.sav", "Variable"])
spop2vars <- unique(metadata[metadata$File == "Studpop2.sav", "Variable"])
febrVars <-  unique(metadata[metadata$File == "Outcome_febr.sav", "Variable"])

new_diags <- readRDS("Data/analysfiler/derive_sos_allaDiagnoser.rds")

setDT(new_diags)

# add new diagnoses added june 2016
dat <- merge(dat, new_diags, by = "lpnr_barn", all.x = TRUE)
rm(new_diags);

#invert MALE sex to FEMALE in order for easier interpretation
dat$FEMALE <- ifelse(dat$MALE == 0,1,0)

#invert diff after new specification
dat$Diff <- dat$Diff *-1

#remove na in new vars 
mergeVars <- grep("^n_", names(dat), value = TRUE)

#table(is.na(dat$n_DM))
#table(dat$n_DM)
setDT(dat)
dat[,(mergeVars):=lapply(.SD, function(x){ifelse(is.na(x),0,x)}), .SDcols = mergeVars]
#table(is.na(dat$n_DM))
#table(dat$n_DM)


dat$DiffKlass <- cut(dat$Diff, breaks = c(-20, -4, -1, 3, 9, 20), include.lowest = TRUE)

# fix data
#dat$DiffKlass <- cut(dat$Diff, quantile(dat$Diff, probs = c(0, 0.1, 0.45, 0.55, 0.9, 1)))
dat$DiffKlass <- factor(dat$DiffKlass, levels = levels(dat$DiffKlass)[c(3,1,2,4,5)])

# updates jan 2017
dat$BMI <- dat$MVIKT/(dat$MLANGD/100)^2
dat$BMI_lower18_5_higher30<- ifelse(dat$BMI<18.5 | dat$BMI>30, 1,0) 
dat$MALDER_younger20_older30 <- ifelse(dat$MALDER < 20 | dat$MALDER >30,1,0)
dat$any_nicotine <- ifelse(as.character(dat$ROK1) == " £No smoking" & as.character(dat$SNUS1) == " £No snuff use", 0, 1)

# impute NA = 0
dat$any_nicotine[is.na(dat$any_nicotine)] <- 0

#-------------------------- make function out of the rest ----------------------
analyzerYeah <- function(condition, sex, type){
  #prints excel sheets with analyses and descriptives for publikation no 3
  # stratified by grdvs length
  
  if(condition == "normal"){
    dat <-
      dat %>% 
      filter(MALE %in% sex) %>% 
      filter(GRDBS/7 >= 37 & GRDBS/7 <= 42)
  }else if(condition == "underburen"){
    dat <-
      dat %>% 
      filter(MALE %in% sex) %>% 
      filter(GRDBS/7 < 37)
  }else if(condition == "overburen"){
    dat <-
      dat %>% 
      filter(MALE %in% sex) %>% 
      filter(GRDBS/7 > 42)
  }else if(condition == "vecka35_38"){
    dat <-
      dat %>% 
      filter(MALE %in% sex) %>% 
      filter(GRDBS >= 35*7 & GRDBS < (38*7)+6)
  }else if(condition == "vecka41_43"){
    dat <-
      dat %>% 
      filter(MALE %in% sex) %>% 
      filter(GRDBS >= 41*7 & GRDBS < (43*7) + 6)
  }else{
    
    cat("all!")
    dat <- dat %>% 
      filter(MALE %in% sex)
  }
  
  
 # Jag skulle också vilja att du gjorde samma analyser i fyra ytterligare grupper - med 
 # enbart födslar i v 37+0-37+6, enbart v 41+0-41+6, v 35+0-37+6 samt 41+0-44+0. 
 # Det är för att renodla de gestationslängder där jag vill titta på underburenhets- resp. 
 # överburenhetsrelaterade utfall, men lägg bara till ytterligare kolumner efter de andra 
 # med alla utfallen för dessa grupper också (på samma sätt, med small/large discrepancies)
  
  if(type == "normal"){
  datPos <- subset(dat, as.character(dat$DiffKlass) %in% c("(-1,3]", "(3,9]", "(9,20]"))
  datPos$DiffKlass <- as.character(datPos$DiffKlass)
  datPos$DiffKlass <- as.factor(datPos$DiffKlass)
  
  datNeg <- subset(dat, as.character(dat$DiffKlass) %in% c("(-1,3]", "[-20,-4]", "(-4,-1]"))
  datNeg$DiffKlass <- as.factor(as.character(datNeg$DiffKlass))
  datNeg$DiffKlass <- factor(datNeg$DiffKlass, levels = c("(-1,3]", "(-4,-1]", "[-20,-4]"))
  }else if(type == "sensitivity"){
    
    
  dat$DiffKlass <- cut(dat$Diff, breaks = c(-20, -6, -1, 3, 6, 20), include.lowest = TRUE)
  
  # fix data
  #dat$DiffKlass <- cut(dat$Diff, quantile(dat$Diff, probs = c(0, 0.1, 0.45, 0.55, 0.9, 1)))
  dat$DiffKlass <- factor(dat$DiffKlass, levels = levels(dat$DiffKlass)[c(3,1,2,4,5)])
    
  datPos <- subset(dat, as.character(dat$DiffKlass) %in% c("(-1,3]", "(3,6]", "(6,20]"))
  datPos$DiffKlass <- as.character(datPos$DiffKlass)
  datPos$DiffKlass <- as.factor(datPos$DiffKlass)
  
  datNeg <- subset(dat, as.character(dat$DiffKlass) %in% c("(-1,3]", "[-20,-6]", "(-6,-1]"))
  datNeg$DiffKlass <- as.factor(as.character(datNeg$DiffKlass))
  datNeg$DiffKlass <- factor(datNeg$DiffKlass, levels = c("(-1,3]", "(-6,-1]", "[-20,-6]"))  
  
  }
  
  nframe <- data.frame(n_pos = nrow(datPos), 
             n_neg = nrow(datNeg),
             condition = condition,
             type = type
             )
  
  write.csv2(nframe, paste0("Data/N/",condition, type,".csv"))
  
  #extraVars <- c("perinatalDeath", "neonatalDeath","MSGA", "MLGA", 
  #             "SGA10_UL", "SGA10_SM", 
  #             "SGA2SD_UL", "SGA2SD_SM", "LGA90_UL", "LGA90_SM", "LGA2SD_UL", 
  #             "LGA2SD_SM", "Pnumoth", "RDS", "Andn_sv", "Gulsot","intrauterin_fdod", "stillborn",
  #             "LowApgar","n_blod_forl", "n_blod_pp", "n_DM", "n_Ind", "n_PE", "n_Perineal", 
  #             "n_prim_varksvaghet", "n_sd", "n_sectio", "n_akut_sectio", "n_sek_varksvaghet", 
  #             "n_utdr_forl", "n_varkrubbning", "n_VE_forceps", "n_blodn_infant", 
  #             "n_EXTREME_SGA", "n_feed_probl", "n_fet_stress", "n_hyperbil", 
  #             "n_hyp_forl_m", "n_hyp_forl_severe", "n_hypoglyc", "n_hypoterm", 
  #             "n_Hypox_forl", "n_Hypox_preg", "n_infek_infant", "n_Kramp", 
  #             "n_LGA", "n_Mec_asp", "n_NEC", "n_Nerv", "n_SGA_IUGR")
  
  
  extraVars <- c("Andn_sv", "Gulsot", "intrauterin_fdod", 
                 "LowApgar", "MLGA", "MSGA", "n_akut_sectio", "n_blod_pp", "n_blodn_infant", 
                 "n_DM", "n_feed_probl", "n_hyperbil", "n_hypoglyc", "n_hypoterm", 
                 "n_Hypox_forl", "n_infek_infant", "n_Kramp", "n_Mec_asp", "n_NEC", 
                 "n_Nerv", "n_PE", "n_Perineal", "n_sd", "n_sectio", "n_sek_varksvaghet", 
                 "n_VE_forceps", "neonatalDeath", "Pnumoth", "RDS")
  
  
  binVars <- c("n_PE"
               ,"n_DM"
               ,"n_sek_varksvaghet"
               ,"n_VE_forceps"
               ,"n_sectio"
               ,"n_akut_sectio"
               ,"n_sd"
               ,"n_blod_pp"
               ,"LowApgar"
               ,"n_Hypox_forl"
               ,"intrauterin_fdod"
               ,"neonatalDeath"
               ,"MSGA"
               ,"MLGA"
  )
  
  binVars <- c(binVars, extraVars)
  binVars <- binVars[!duplicated(binVars)]
  
  
  
  if(type == "sensitivity" & condition == "vecka41_43"){
    
    binVars <- c("n_PE"
                 ,"n_DM"
                 ,"n_sek_varksvaghet"
                 ,"n_VE_forceps"
                 ,"n_sectio"
                 ,"n_akut_sectio"
                 ,"n_sd"
                 ,"n_blod_pp"
                 ,"n_Hypox_forl"
                 ,"intrauterin_fdod"
                 ,"neonatalDeath"
                 ,"MSGA"
                 ,"MLGA"
    )
    
    
  }

  
 # example
  #xx <- simpleYeah2(binVars[1], datPos = datPos, datNeg = datNeg)
   
  #ll <- lapply(binVars, simpleYeah, datPos = datPos, datNeg = datNeg)
  #browser()
  no_cores <- detectCores() - 2
  # Initiate cluster
  cl <- makeCluster(no_cores)
  
  clusterExport(cl, c("simpleYeah2", "glm_model1", "glm_model2", "glm_model3", "datPos", "datNeg", "tabler", "type"), envir = environment())
  
  ll <- parLapply(cl, binVars, simpleYeah2, datPos = datPos, datNeg = datNeg, type = type)
  
  # close cluster
  parallel::stopCluster(cl)
  
  ll <- lapply(1:length(ll), function(i){ll[[i]]$OR = as.numeric(ll[[i]]$OR)
                                  return(ll[[i]])})

  #browser()
  
  out <- bind_rows(ll)
  
  sex <- ifelse(length(sex) == 2, "All",
                ifelse(sex == 1, "MALE", "FEMALE"))
  
  saveRDS(out, paste0("Data/tabell/deltabell_", format(Sys.time(), "%Y%m%d_%H%M"), "_",condition,"_",sex, ".rds"))
  #close workbook
  return(out)
}

#-------------------------------------------------------------------------------

# ---------------------------- run function ------------------------------------
 system.time(analyzerYeah("Alla", sex = c(0,1), type = "normal"))
 gc()

 #system.time(analyzerYeah("underburen", sex = c(0,1)))
 #gc()

 #system.time(analyzerYeah("overburen", sex = c(0,1))) #problems with var after pnumoth or with pnumoth
 #gc()

 #system.time(analyzerYeah("normal", sex = c(0,1)))
 #gc()


 system.time(analyzerYeah("vecka35_38", sex = c(0,1), type = "normal"))
 gc()

 system.time(analyzerYeah("vecka41_43", sex = c(0,1), type = "normal"))
 gc()


# #-----------------------------  FEMALES ----------------------------------------
 system.time(analyzerYeah("Alla", sex = 0, type = "normal"))
 gc()
# 
# #system.time(analyzerYeah("underburen", sex = 0))
# #gc()
# 
# #system.time(analyzerYeah("overburen", sex = 0)) #problems with var after pnumoth or with pnumoth
# #gc()
# 
# #system.time(analyzerYeah("normal", sex = 0))
# #gc()
# 
 system.time(analyzerYeah("vecka35_38", sex = 0, type = "normal"))
 gc()
# 
 system.time(analyzerYeah("vecka41_43", sex = 0, type = "normal"))
 gc()
# 
# 
# #-------------------------------- MALES ----------------------------------------
 system.time(analyzerYeah("Alla", sex = 1, type = "normal"))
 gc()
# #system.time(analyzerYeah("underburen", sex = 1))
# #gc()
# 
# #system.time(analyzerYeah("overburen", sex = 1)) #problems with var after pnumoth or with pnumoth
# #gc()
# 
# #system.time(analyzerYeah("normal", sex = 1))
# #gc()
# 
 system.time(analyzerYeah("vecka35_38", sex = 1, type = "normal"))
 gc()
# 
 system.time(analyzerYeah("vecka41_43", sex = 1, type = "normal"))
 gc()



#neg <- 
#  dat %>% 
#  filter(type == "neg")
#-------------------------------------------------------------------------------

#system.time(analyzerYeah("Alla"))
#--------------------------------- testruns ------------------------------------
#tabs <- readRDS("C:/Users/perwim/Desktop/TMP/MOVE/MERIT/P105_ultraljud_publikation3/Data/tabell/deltabell_20170131_1723_Alla_All.rds")
#tabs <- readRDS("C:/Users/perwim/Desktop/TMP/MOVE/MERIT/P105_ultraljud_publikation3/Data/tabell/deltabell_20170131_1724_underburen_All.rds")

 
# dat$DiffKlass_tmp <- cut(dat$Diff, breaks = c(-20, -6, -1, 3, 6, 20), include.lowest = TRUE)
# 
# # fix data
# #dat$DiffKlass <- cut(dat$Diff, quantile(dat$Diff, probs = c(0, 0.1, 0.45, 0.55, 0.9, 1)))
# dat$DiffKlass_tmp <- factor(dat$DiffKlass_tmp, levels = levels(dat$DiffKlass_tmp)[c(3,1,2,4,5)])
# 
# 
# 
#testModel<-
# dat %>%
#  filter(MALE %in% c(0,1)) %>% # both
#  #filter(GRDBS >= 41*7 & GRDBS < (43*7) + 6) %>% # vecka 41 till 43
#   #filter(GRDBS/7 < 37) %>% #underburen
#   #filter(GRDBS/7 >= 41 & GRDBS/7 <= 44) %>% #41 till 44
#   #filter(as.character(DiffKlass_tmp) %in% c("(-1,3]", "(3,6]", "(6,20]")) %>%  # Positiv diff new
#   #filter(as.character(DiffKlass) %in% c("(-1,3]", "(3,9]", "(9,20]")) %>%  # Positiv diff
#   #filter(as.character(DiffKlass) %in% c("(-1,3]", "[-20,-4]", "(-4,-1]")) %>% # negativ diff
#  filter(as.character(DiffKlass_tmp) %in% c("(-1,3]", "(-6,-1]", "[-20,-6]")) %>%  # negativ new
#   glm(n_PE~DiffKlass_tmp+MSGA+MLGA, data = ., family = "binomial")


#tmp <- glm_model3("LowApgar", testModel)
#logitPos3 <- glm_model3(var, datPos) 



# "n_PE"
# ,"n_DM"
# ,"n_sek_varksvaghet"
# ,"n_VE_forceps"
# ,"n_sectio"
# ,"n_akut_sectio"
# ,"n_sd"
# ,"n_blod_pp"
# ,"LowApgar"
# ,"n_Hypox_forl"
# ,"intrauterin_fdod"
# ,"neonatalDeath"
# ,"MSGA"
# ,"MLGA"



#  
#  exp(coef(testModel))  
#  exp(confint(testModel))
#  
#  realModelData <- readRDS("Data/tabell/deltabell_20160909_1307_underburen.rds")
#  
#  realModelData %>% 
#    filter(var == "MALE" & type == "pos")

#-------------------------------------------------------------------------------

alla <- readRDS("Data/tabell/deltabell_20170530_1819_Alla_All.rds")
vecka35_38 <- readRDS("Data/tabell/deltabell_20170530_1840_vecka35_38_All.rds") 
vecka41_43 <- readRDS("Data/tabell/deltabell_20170530_1909_vecka41_43_All.rds") 

alla_female <- readRDS("Data/tabell/deltabell_20170530_2004_Alla_FEMALE.rds")
vecka35_38_female <- readRDS("Data/tabell/deltabell_20170530_2012_vecka35_38_FEMALE.rds") 
vecka41_43_female <- readRDS("Data/tabell/deltabell_20170530_2023_vecka41_43_FEMALE.rds") 
#
alla_male <- readRDS("Data/tabell/deltabell_20170530_2122_Alla_MALE.rds")
vecka35_38_male <- readRDS("Data/tabell/deltabell_20170530_2132_vecka35_38_MALE.rds") 
vecka41_43_male <- readRDS("Data/tabell/deltabell_20170530_2148_vecka41_43_MALE.rds") 


xlsxWriter(alla, "alla_all", population = "normal")
xlsxWriter(vecka35_38, "vecka35_38_all", population = "normal")
xlsxWriter(vecka41_43, "vecka41_43_all", population = "normal")

xlsxWriter(alla_female, "alla_female", population = "normal")
xlsxWriter(vecka35_38_female, "vecka35_38_female", population = "normal")
xlsxWriter(vecka41_43_female, "vecka41_43_female", population = "normal")

xlsxWriter(alla_male, "alla_male", population = "normal")
xlsxWriter(vecka35_38_male, "vecka35_38_male", population = "normal")
xlsxWriter(vecka41_43_male, "vecka41_43_male", population = "normal")


# write summary file
createFiles <- function(cls, population = "normal"){
  
  files <- grep(paste0("^",cls ,"_diff_[av]"), list.files("Output"), value = TRUE)
  
  if(population == "sensitivity"){
    files <- grep("alla|vecka", files, value = TRUE)
    }
  
  ll = lapply(files, function(x){openxlsx::read.xlsx(paste0("Output/",x))})
  
  list_all <- grep("_all.xlsx$", files)
  list_female <- grep("_female.xlsx$", files)
  list_male <- grep("_male.xlsx$", files)
  
  # LowApgar saknas för vissa lägg till en extra rad
  ll <- lapply(ll, function(x){
    if(!("LowApgar" %in% x$var)){
      
      extra <- x[1,]
      extra[,"var"] <- "LowApgar"
      extra[,!grepl("var", names(extra))] <- "-"
      
      out <- rbind(x, extra)
      
      out<-out[c(1:9,15,10:14),]
      
      return(out)
    }else{
      return(x)
    }
  })
    

  out_all <- do.call("cbind", ll[list_all])
  out_female <- do.call("cbind", ll[list_female])
  out_male <- do.call("cbind", ll[list_male])
  
  out_all <- out_all[,!grepl("^type$", names(out_all))]
  out_female<- out_female[,!grepl("^type$", names(out_female))]
  out_male <- out_male[,!grepl("^type$", names(out_male))]
  
  
  filename <- paste0("Output/","Tabell_", cls, "_" ,format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")
  
  wb <- openxlsx::createWorkbook()
  
  openxlsx::addWorksheet(wb,"all")
  openxlsx::addWorksheet(wb,"female")
  openxlsx::addWorksheet(wb,"male")
  
  writeData(wb, "all", out_all)
  writeData(wb, "female", out_female)
  writeData(wb, "male", out_male)
  
  saveWorkbook(wb, filename)
}

createFiles("positiv", population = "normal")
createFiles("negativ", population = "normal")
#--------------------------- beräkna punktprevalenser --------------------------
calcPrevalence <- function(var,data, sex){
  data <- subset(data,MALE == sex)
  
  n <- sum(data[,get(var),], na.rm = TRUE)
  antal <- nrow(data)
  
  prevalence <- round((n/antal)*100000)
}



binVars <- c("n_PE"
             ,"n_DM"
             ,"n_sek_varksvaghet"
             ,"n_VE_forceps"
             ,"n_sectio"
             ,"n_akut_sectio"
             ,"n_sd"
             ,"n_blod_pp"
             ,"LowApgar"
             ,"n_Hypox_forl"
             ,"intrauterin_fdod"
             ,"neonatalDeath"
             ,"MSGA"
             ,"MLGA"
             ,"n_infek_infant"
             ,"n_blodn_infant"
             ,"n_hypoglyc"
             ,"Pnumoth"
             ,"RDS"
             ,"Andn_sv"
             ,"n_hyperbil"
             ,"n_hypoterm"
             ,"n_feed_probl"
             ,"n_Mec_asp"
             ,"n_Kramp"
             ,"n_Nerv"
             ,"n_Perineal"
             )

ll <- sapply(binVars, calcPrevalence, data = dat, sex = 0)




#----------------------------- antal per grupp ---------------------------------
if(FALSE){
  conditions <- c("All", "normal", "underburen", "overburen", "vecka35_38","vecka41_43")
  sexes <- list(0, 1, c(0,1))
  
  filter_data = function(data, condition, sex){
    
    dat <- data
    
    if(condition == "normal"){
      dat <-
        dat %>% 
        filter(MALE %in% sex) %>% 
        filter(GRDBS/7 >= 37 & GRDBS/7 <= 42)
    }else if(condition == "underburen"){
      dat <-
        dat %>% 
        filter(MALE %in% sex) %>% 
        filter(GRDBS/7 < 37)
    }else if(condition == "overburen"){
      dat <-
        dat %>% 
        filter(MALE %in% sex) %>% 
        filter(GRDBS/7 > 42)
    }else if(condition == "vecka35_38"){
      dat <-
        dat %>% 
        filter(MALE %in% sex) %>% 
        filter(GRDBS >= 35*7 & GRDBS < (38*7)+6)
    }else if(condition == "vecka41_43"){
      dat <-
        dat %>% 
        filter(MALE %in% sex) %>% 
        filter(GRDBS >= 41*7 & GRDBS < (43*7) + 6)
    }else{
      #browser()
      cat("all!")
      dat <- dat %>% 
        filter(MALE %in% sex)
    }
    
    out <- data.frame(table(dat$DiffKlass))
    out$src <- condition
    out$sex <- paste0(sex, collapse = "")
    
    return(out)
  }
  
  
  all <- lapply(conditions, filter_data, sex = c(0,1), data = dat)
  males <- lapply(conditions, filter_data, sex = c(1), data = dat)
  females <- lapply(conditions, filter_data, sex = c(0), data = dat)
  
  
  all <- do.call("rbind", all)
  males <- do.call("rbind", males)
  females <- do.call("rbind", females)
  
  out <- rbind(all, females, males)
  
  out$Var1 <- ifelse(out$Var1 == "(-1,3]", "Reference",
                     ifelse(out$Var1 == "[-20,-4]", "Large negative",
                            ifelse(out$Var1 == "(-4,-1]", "Small negative",
                                   ifelse(out$Var1 == "(3,9]", "Small positive",
                                          ifelse(out$Var1 == "(9,20]", "Large positive",NA)
                                   ))))
  
  out$sex <- ifelse(out$sex == 1, "Male", 
                    ifelse(out$sex == 0, "Female", "All"
                    ))
  
  names(out) <- c("Category", "N", "Population", "Sex")
  
  out %>% 
    select(Population, Category, Sex, N) %>% 
    arrange(Population, Sex) %>% 
    openxlsx::write.xlsx("Output/Antal_per_klass.xlsx")
  
  
  
  
}


