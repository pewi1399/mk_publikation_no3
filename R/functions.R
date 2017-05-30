
printerDiffClass <- function(var, dat, datPos, datNeg, wb){
  #var: binary variable for printing
#   dat$test <- cut(dat$Diff, breaks = seq(-20,20,5))
#   dat$test <- factor(dat$test, levels = levels(dat$test)[c(4,1,2,3,5,6,7,8)])
#   test <- glm(paste0(var, "~test"), data = dat, family = "binomial")
#   exp(coef(test))
  print(var)
  #browser()
  logit <- tryCatch(glm(paste0(var, "~DiffKlass"), data = dat, family = "binomial"), error = function(e) "Unable to fit model")
  logitNeg <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datNeg, family = "binomial"), error = function(e) "Unable to fit model")
  logitPos <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datPos, family = "binomial"), error = function(e) "Unable to fit model")
 
#   logitMulti <- tryCatch(glm(paste0(var, "~DiffKlass+"), data = dat, family = "binomial"), error = function(e) "Unable to fit model")
#   logitNegMulti <- tryCatch(glm(paste0(var, "~DiffKlass+"), data = datNeg, family = "binomial"), error = function(e) "Unable to fit model")
#   logitPosMulti <- tryCatch(glm(paste0(var, "~DiffKlass+"), data = datPos, family = "binomial"), error = function(e) "Unable to fit model")
   
  #if(class(dat[,var]) == "factor"){
  out <- tryCatch(tablePipe(logit) %>% rbind(c(paste0("DiffKlass",levels(dat$DiffKlass)[1]),"1", "-"), .), error = function(e) "Unable to fit model, probable cause; to few cases or to little variation")
  outPos <- tryCatch(tablePipe(logitPos) %>% rbind(c(paste0("DiffKlass",levels(dat$DiffKlass)[1]),"1", "-"), .), error = function(e) "Unable to fit model, probable cause; to few cases or to little variation")
  outNeg <- tryCatch(tablePipe(logitNeg) %>% rbind(c(paste0("DiffKlass",levels(dat$DiffKlass)[1]),"1", "-"), .), error = function(e) "Unable to fit model, probable cause; to few cases or to little variation")
  #}
  
  # Create sheet for variable
  createSheet(wb, name = var)
  
  # Create a CellStyle with yellow solid foreground
  CellColor1 <- createCellStyle(wb)
  setFillPattern(CellColor1, fill = XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(CellColor1, color = XLC$"COLOR.LIGHT_GREEN")
  
  CellColor2 <- createCellStyle(wb)
  setFillPattern(CellColor2, fill = XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(CellColor2, color = XLC$"COLOR.ROSE")
  
  writeWorksheet(wb, out, sheet = var,  startCol = 1, startRow = 42)
  writeWorksheet(wb, outPos, sheet = var,  startCol = 5, startRow = 42)
  writeWorksheet(wb, outNeg, sheet = var,  startCol = 9, startRow = 42)
  
  # apply the CellStyle to a given cell, here: (10,10)
  setCellStyle(wb, sheet=var, row=42, col=5:7, cellstyle=CellColor1)
  setCellStyle(wb, sheet=var, row=42, col=9:11, cellstyle=CellColor2)
  
}

#test_that("output from pipe is consistent with raw object",
#           exp(coef(summary(logit1))[,"Estimate"])
# klassification case
#logit2 <- glm(paste0(var, "~Diff"), data = dat, family = "binomial")

#-------------------------------------------------------------------------------

#-------------------------- Categorical outcomes ------------------------------- 
# for categorical variables
printerCategorical <- function(var, dat, datPos, datNeg, wb){
#browser()
  createSheet(wb, name = var)
  
  out <- data.frame(table(dat$DiffKlass, dat[,var])) %>% 
    dcast(Var1~Var2, value.var = "Freq") %>% 
    rename(DiffLevel = Var1)
  
  outPos <- data.frame(table(datPos$DiffKlass, datPos[,var])) %>% 
    dcast(Var1~Var2, value.var = "Freq") %>% 
    rename(DiffLevel = Var1)
  
  outNeg <- data.frame(table(datNeg$DiffKlass, datNeg[,var])) %>% 
    dcast(Var1~Var2, value.var = "Freq") %>% 
    rename(DiffLevel = Var1)
  
  # get pvalue   
  pval <- chisq.test(table(dat$DiffKlass, dat[,var]))$p.value %>% 
    round(3) %>% 
    format(nsmall=3) %>% 
    ifelse(.=="0.000", "<0.001",.)
  
  pvalPos <- chisq.test(table(datPos$DiffKlass, datPos[,var]))$p.value %>% 
    round(3) %>% 
    format(nsmall=3) %>% 
    ifelse(.=="0.000", "<0.001",.)
  
  pvalNeg <- chisq.test(table(datNeg$DiffKlass, datNeg[,var]))$p.value %>% 
    round(3) %>% 
    format(nsmall=3) %>% 
    ifelse(.=="0.000", "<0.001",.)
  
  
  out$p_value <- pval
  out$p_value[2:nrow(out)] <- ""
  
  outPos$p_value <- pvalPos
  outPos$p_value[2:nrow(outPos)] <- ""
  
  outNeg$p_value <- pvalNeg
  outNeg$p_value[2:nrow(outNeg)] <- ""
  
  # Create a CellStyle with yellow solid foreground
  CellColor1 <- createCellStyle(wb)
  setFillPattern(CellColor1, fill = XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(CellColor1, color = XLC$"COLOR.LIGHT_GREEN")
  
  CellColor2 <- createCellStyle(wb)
  setFillPattern(CellColor2, fill = XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(CellColor2, color = XLC$"COLOR.ROSE")
  
  writeWorksheet(wb, out, sheet = var,  startCol = 1, startRow = 48)
  writeWorksheet(wb, outPos, sheet = var,  startCol = ncol(out)+2, startRow = 48)
  writeWorksheet(wb, outNeg, sheet = var,  startCol = ncol(out)*2+3, startRow = 48)
  
  # apply the CellStyle to a given cell, here: (10,10)
  setCellStyle(wb, sheet=var, row=48, col=(ncol(out)+2):(ncol(out)+2+4), cellstyle=CellColor1)
  setCellStyle(wb, sheet=var, row=48, col=(ncol(out)*2+3):(ncol(out)*2+3+4), cellstyle=CellColor2)
}


#-------------------------------------------------------------------------------

tablePipe <- function(model){
  #Model: an object of class GLM with a binary link
  
out <- model %>% 
  summary %>% 
  coef %>% 
  data.frame %>% 
  mutate(Variable = row.names(coef(summary(model))),
         OR = format(round(exp(Estimate),3), nsmall = 3), 
         "p_value" = .$"Pr...z..") %>% 
  select(Variable, OR, p_value) %>% 
  mutate(p_value = format(round(p_value,3), nsmall = 3)) %>% 
  mutate(p_value = ifelse(p_value == "0.000", "<0.001", p_value)) %>% 
  filter(Variable != "(Intercept)")

return(out)
}


# -------------------------- new functions fall 2016 ---------------------------
tabler <- function(model, type, var, modelname = NULL){
  
  if(is.character(model)){
    
  dlog <- data.frame(var, name = "-", type = type , OR = model , CI= model)
    
  }else{
  OR <- exp(coef(model))
  namelist <- names(model$coefficients)
  
  #browser()
  CI <- tryCatch(exp(confint(model)), error = function(e) rep("NA", length(OR)))
  
  dlog <- data.frame(var, name = namelist, type = type , OR = OR, CI)
  }
  
  if(!is.null(modelname)){
    dlog$modelnumber <- modelname  
  }
  
  return(dlog)
}

simpleYeah <- function(var =NULL, datPos = datPos, datNeg = datNeg){
  #library(reshape2)
  library(data.table)
  #library(XLConnect)
  library(openxlsx)
  #library(ggplot2)
  library(dplyr)
  #library(reshape2)
 #library(magrittr)
  #library(tidyr)
  
  print(var)
  #browser()
  #logit <- tryCatch(glm(paste0(var, "~DiffKlass"), data = dat, family = "binomial"), error = function(e) "Unable to fit model")
  logitNeg <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datNeg, family = "binomial"), error = function(e) "Unable to fit model")
  logitPos <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datPos, family = "binomial"), error = function(e) "Unable to fit model")
  
  #dlog <- data.frame(var, name = names(logit$coefficients), type = "all", OR = exp(coef(logit)), exp(confint(logit)))
  #dNeg <- data.frame(var, name = names(logitNeg$coefficients), type = "Negative", OR = exp(coef(logitNeg)), exp(confint(logitNeg)))
  #dPos <- data.frame(var, name = names(logitPos$coefficients), type = "Positive", OR = exp(coef(logitPos)), exp(confint(logitPos)))
  
  #dlog <- tabler(logit, type = "all", var = var)
  dNeg <- tabler(logitNeg, type = "neg", var = var)
  dPos <- tabler(logitPos, type = "pos", var = var)
  
  
  out <- bind_rows(list(dNeg, dPos))
  rm(list = c( "dNeg", "dPos", "logitNeg", "logitPos"))
  gc()
  return(out)
}


xlsxWriter <- function(dat, name, population){
  #levelList <- c(
  #  "n_PE",
  #  "n_DM",
  #  "n_Ind", 
  #  "n_varkrubbning",
  #  "n_utdr_forl",
  #  "n_prim_varksvaghet",
  #  "n_sek_varksvaghet",
  #  "n_sd", 
  #  "n_blod_forl",
  #  "n_blod_pp", 
  #  "n_Perineal",
  #  "n_VE_forceps", 
  #  "n_sectio",
  #  "n_akut_sectio",
  #  "LowApgar",
  #  "n_fet_stress", 
  #  "n_Hypox_preg",
  #  "n_Hypox_forl",
  #  "n_hyp_forl_severe",
  #  "n_hyp_forl_m",
  #  "stillborn",
  #  "intrauterin_fdod",
  #  "perinatalDeath",
  #  "neonatalDeath", 
  #  "MSGA",
  #  "MLGA",
  #  "SGA10_UL",
  #  "SGA10_SM",
  #  "SGA2SD_UL",
  #  "SGA2SD_SM",
  #  "LGA90_UL",
  #  "LGA90_SM",
  #  "LGA2SD_UL",
  #  "LGA2SD_SM",
  #  "n_SGA_IUGR",
  #  "n_EXTREME_SGA",
  #  "n_LGA",
  #  "n_infek_infant", 
  #  "n_blodn_infant",
  #  "n_hypoglyc", 
  #  "Pnumoth",
  #  "RDS",
  #  "Andn_sv", 
  #  "Gulsot",
  #  "n_hyperbil",
  #  "n_NEC", 
  #  "n_hypoterm",
  #  "n_feed_probl",
  #  "n_Mec_asp",
  #  "n_Kramp",
  #  "n_Nerv", 
  #  "MALE",
  #  "FEMALE"
  #)
  
  
 # levelList <- c("n_PE"
 # ,"n_DM"
 # ,"n_sek_varksvaghet"
 # ,"n_VE_forceps"
 # ,"n_sectio"
 # ,"n_akut_sectio"
 # ,"n_sd"
 # ,"n_blod_pp"
 # ,"n_Perineal"
 # ,"LowApgar"
 # ,"n_Hypox_forl"
 # ,"intrauterin_fdod"
 # ,"neonatalDeath"
 # ,"MSGA"
 # ,"MLGA")
  
  levelList <- c("n_PE", "n_DM", "n_sek_varksvaghet", "n_VE_forceps", "n_sectio", 
    "n_akut_sectio", "n_sd", "n_blod_pp", "LowApgar", "n_Hypox_forl", 
    "intrauterin_fdod", "neonatalDeath", "MSGA", "MLGA", "Andn_sv", 
    "Gulsot", "n_blodn_infant", "n_feed_probl", "n_hyperbil", "n_hypoglyc", 
    "n_hypoterm", "n_infek_infant", "n_Kramp", "n_Mec_asp", "n_NEC", 
    "n_Nerv", "n_Perineal", "Pnumoth", "RDS")
  
  
  #----------------------------- formatera tabeller ------------------------------
  if(length(names(dat)) == 7){
      names(dat) <- c("var", "name", "type", "OR", "lo", "hi", "model")
      #dat$CI <- NULL
  } else if(length(names(dat)) == 8){
      names(dat) <- c("var", "name", "type", "OR", "lo", "hi", "model", "CI")
  }
  
  
  dat[,c("OR", "lo", "hi")] <- lapply(dat[,c("OR", "lo", "hi")], function(x) ifelse(x>1000,NA,x))
  
  dat$CI <- paste0("(", format(round(dat$lo,2), nsmall = 2), 
                   " - ", format(round(dat$hi,2), nsmall = 2),
                   ")"
  )
  
  dat$OR <- format(round(dat$OR,2), nsmall = 2)
  
  dat$lo <- NULL
  dat$hi <- NULL
  
  dat <- 
    dat %>% 
    filter(name != "(Intercept)")
  
  pos <- 
    dat %>% 
    filter(type == "pos") %>% 
    gather(key, value, -var, -name, -type, -model)
  
  if(population == "normal"){
  pos$name <- ifelse(pos$name == "DiffKlass(3,9]", "Small_discr", 
                     ifelse(pos$name == "DiffKlass(9,20]", "Large_discr", pos$name)
                     )
  } else if(population == "sensitivity"){
    pos$name <- ifelse(pos$name == "DiffKlass(3,6]", "Small_discr", 
                       ifelse(pos$name == "DiffKlass(6,20]", "Large_discr", pos$name)
    )
    
  }
  
  pos <- reshape2::dcast(pos, var+type+model ~name+key, fill = "value")
  
  pos$type <- "positiv/underburen"
  
  pos <-
    pos %>% 
    mutate(gg = paste(Small_discr_OR,
                      Small_discr_CI,
                      Large_discr_OR,
                      Large_discr_CI
    )
    ) %>% 
    mutate(gg = grepl("NA",gg)) %>% 
    mutate(
      Small_discr_OR = ifelse(gg == TRUE, "-", Small_discr_OR),
      Small_discr_CI = ifelse(gg == TRUE, "-", Small_discr_CI),
      Large_discr_OR = ifelse(gg == TRUE, "-", Large_discr_OR),
      Large_discr_CI = ifelse(gg == TRUE, "-", Large_discr_CI)
    ) %>% 
    select(
      var,
      type,
      model,
      Small_discr_OR,
      Small_discr_CI,
      Large_discr_OR,
      Large_discr_CI
    )  
  
  pos$var <- as.factor(pos$var)
  pos$var <- factor(pos$var, levels = levelList)

  
  
  pos <- pos[order(pos$var),]
  extraRow <- pos[1,]
  extraRow[,] <- name
  
  pos <- bind_rows(extraRow, pos)
  
  # remove stupid value label fill
  pos[,names(pos)] <- lapply(names(pos), function(x){gsub("value", "-",pos[,x])})
  
  openxlsx::write.xlsx(pos, paste0("output/positiv_diff_", name, ".xlsx"))
  
  #-------------------------------------------------------------------------------
  
  
  neg <- 
    dat %>% 
    filter(type == "neg") %>% 
    gather(key, value, -var, -name, -type, -model)
  
  
  if(population == "normal"){
  neg$name <- ifelse(neg$name == "DiffKlass(-4,-1]","Small_discr",
                     ifelse(neg$name == "DiffKlass[-20,-4]","Large_discr", neg$name)
  )
  } else if(population == "sensitivity"){
    neg$name <- ifelse(neg$name == "DiffKlass(-6,-1]","Small_discr",
                       ifelse(neg$name == "DiffKlass[-20,-6]","Large_discr", neg$name)
    )
    
  }
  
  neg <- reshape2::dcast(neg, var+type+model ~name+key, fill = "value")
  
  neg$type <- "negativ/underburen"
  
  neg <-
    neg %>% 
    mutate(gg = paste(Small_discr_OR,
                      Small_discr_CI,
                      Large_discr_OR,
                      Large_discr_CI
    )
    ) %>% 
    mutate(gg = grepl("NA|Unable",gg)) %>% 
    mutate(
      Small_discr_OR = ifelse(gg == TRUE, "-", Small_discr_OR),
      Small_discr_CI = ifelse(gg == TRUE, "-", Small_discr_CI),
      Large_discr_OR = ifelse(gg == TRUE, "-", Large_discr_OR),
      Large_discr_CI = ifelse(gg == TRUE, "-", Large_discr_CI)
    ) %>% 
    select(
      var,
      type,
      model,
      Small_discr_OR,
      Small_discr_CI,
      Large_discr_OR,
      Large_discr_CI
    )


  
  neg$var <- as.factor(neg$var)
  neg$var <- factor(neg$var, levels = levelList)

  neg <- neg[order(neg$var),]
  
  extraRow <- neg[1,]
  extraRow[,] <- name
  
  neg <- bind_rows(extraRow, neg)
  
 # replace stupid value fill... 
  neg[,names(neg)] <- lapply(names(neg), function(x){gsub("value", "-",neg[,x])})

  openxlsx::write.xlsx(neg, paste0("output/negativ_diff_", name, ".xlsx"))
}


#------------------------------- jan 2017 --------------------------------------
simpleYeah2 <- function(var =NULL, datPos = datPos, datNeg = datNeg, type = "normal"){
  #library(reshape2)
  library(data.table)
  #library(XLConnect)
  library(openxlsx)
  #library(ggplot2)
  library(dplyr)
  #library(reshape2)
  #library(magrittr)
  #library(tidyr)
  
  #print(var)
  write.csv(var, paste0("Slask/", var, ".csv"))
  #browser()
  #logit <- tryCatch(glm(paste0(var, "~DiffKlass"), data = dat, family = "binomial"), error = function(e) "Unable to fit model")
  #logitNeg <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datNeg, family = "binomial"), error = function(e) "Unable to fit model")
  #logitPos <- tryCatch(glm(paste0(var, "~DiffKlass"), data = datPos, family = "binomial"), error = function(e) "Unable to fit model")
  
  if(type == "normal"){ #normal
  logitNeg1 <- glm_model1(var, datNeg)
  logitPos1 <- glm_model1(var, datPos)
  
  logitNeg2 <- glm_model2(var, datNeg)
  logitPos2 <- glm_model2(var, datPos)
  }
  
  logitNeg3 <- glm_model3(var, datNeg, type)
  logitPos3 <- glm_model3(var, datPos, type)  
  
  write.csv(var, paste0("Slask/",var, "model_ok.csv")) 
  #dlog <- data.frame(var, name = names(logit$coefficients), type = "all", OR = exp(coef(logit)), exp(confint(logit)))
  #dNeg <- data.frame(var, name = names(logitNeg$coefficients), type = "Negative", OR = exp(coef(logitNeg)), exp(confint(logitNeg)))
  #dPos <- data.frame(var, name = names(logitPos$coefficients), type = "Positive", OR = exp(coef(logitPos)), exp(confint(logitPos)))
  
  #dlog <- tabler(logit, type = "all", var = var)
  if(type == "normal"){ #normal
  dNeg1 <- tabler(logitNeg1, type = "neg", var = var, modelname = "model1")
  dPos1 <- tabler(logitPos1, type = "pos", var = var, modelname = "model1")
  
  dNeg2 <- tabler(logitNeg2, type = "neg", var = var, modelname = "model2")
  dPos2 <- tabler(logitPos2, type = "pos", var = var, modelname = "model2")
  }
  
  dNeg3 <- tabler(logitNeg3, type = "neg", var = var, modelname = "model3")
  dPos3 <- tabler(logitPos3, type = "pos", var = var, modelname = "model3")
  
  write.csv(var, paste0("Slask/",var, "table_ok.csv")) 
  
  if(type == "normal"){ #normal
  out <- bind_rows(list(dNeg1, dPos1, dNeg2, dPos2, dNeg3, dPos3))
  rm(list = c( "dNeg1", "dPos1", "logitNeg1", "logitPos1", 
               "dNeg2", "dPos2", "logitNeg2", "logitPos2",
               "dNeg3", "dPos3", "logitNeg3", "logitPos3"))
  }else{ #sensitivity
   
  out <- bind_rows(list(dNeg3, dPos3))
  rm(list = c( "dNeg1", "dPos1", "logitNeg1", "logitPos1", 
               "dNeg2", "dPos2", "logitNeg2", "logitPos2",
               "dNeg3", "dPos3", "logitNeg3", "logitPos3")) 
  }
  
  write.csv(var, paste0("Slask/Done",var, ".csv"))  
  
  
  gc()
  return(out)
}


#----------------------------- first model ------------------------------------- 
glm_model1 <- function(var, dataset){
  
  logit1 <- tryCatch(glm(paste0(var, "~DiffKlass + BMI_lower18_5_higher30 + MALDER_younger20_older30 + any_nicotine + FAMSIT_EJ_SAMBO + ARBETE_EJ"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
  
  return(logit1)
}

#----------------------------- second model ----------------------------------- 
glm_model2 <- function(var, dataset){
  
  if(var == "n_DM"){
    logit2 <- tryCatch(glm(paste0(var, "~DiffKlass + BMI_lower18_5_higher30 + MALDER_younger20_older30 + any_nicotine + n_PE + FAMSIT_EJ_SAMBO + ARBETE_EJ"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
  }else if(var == "n_PE"){
    logit2 <- tryCatch(glm(paste0(var, "~DiffKlass + BMI_lower18_5_higher30 + MALDER_younger20_older30 + any_nicotine + n_DM + FAMSIT_EJ_SAMBO + ARBETE_EJ"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
  }else{
    logit2 <- tryCatch(glm(paste0(var, "~DiffKlass + BMI_lower18_5_higher30 + MALDER_younger20_older30 + any_nicotine + n_DM + n_PE + FAMSIT_EJ_SAMBO + ARBETE_EJ"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
  }
  
  return(logit2)
}

#----------------------------- third model ------------------------------------- 
glm_model3 <- function(var, dataset, type = "normal"){
  if(type == "normal"){
    if(var == "MSGA"){
      
      logit3 <- tryCatch(glm(paste0(var, "~DiffKlass + MLGA"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
      
    }else if (var == "MLGA"){
      
      logit3 <- tryCatch(glm(paste0(var, "~DiffKlass + MSGA"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
      
    } else {
      
      logit3 <- tryCatch(glm(paste0(var, "~DiffKlass + MSGA + MLGA"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
    }
  }else if(type == "sensitivity"){
    logit3 <- tryCatch(glm(paste0(var, "~DiffKlass"), data = dataset, family = "binomial"), error = function(e) "Unable to fit model")
  }
    
  return(logit3)
}

