rm(list=objects())
library(data.table)
#library(XLConnect)
library(openxlsx)
library(ggplot2)
library(reshape2)
library(dplyr)

#load data
load(file=file.path("Data","Publikation2","dmp01.Rdata"))
outcome_febr <- data.table(data.frame(readRDS("Data/outcome_febr.rds")))
setnames(outcome_febr, "Lpnr_barn", "lpnr_barn")

metadata <- openxlsx::read.xlsx(file.path("Dokument", "Outcomes_160208.xlsx"), sheet = 1)
spop1vars <- unique(metadata[metadata$File == "Studpop1.sav", "Variable"])
spop2vars <- unique(metadata[metadata$File == "Studpop2.sav", "Variable"])
febrVars <-  unique(metadata[metadata$File == "Outcome_febr.sav", "Variable"])

spop1 <- data.table(data.frame(readRDS("Data/Studpop_1.rds")))
spop1 <- spop1[,c("lpnr_barn", spop1vars), with = FALSE] 

spop2 <- data.table(data.frame(readRDS("Data/Studpop2.rds")))
spop2 <- spop2[,c("lpnr_barn", "lpnr_mor", "BFODDAT", spop2vars), with = FALSE]

#save copy
dat0 <- dmp01; rm(dmp01)
names(dat0) <- gsub("\xe4", "", names(dat0))
dat0 <- data.table(dat0)
#-------------------------------- merge files ----------------------------------
setkey(outcome_febr, "lpnr_barn")
setkey(dat0, "lpnr_barn")
setkey(spop2,"lpnr_barn")
setkey(spop1,"lpnr_barn")

# dupes exist in studpop 2
spop2 <- spop2[!duplicated(spop2)]

# merge for one studypopulation
dat1 <- merge(dat0, spop2) ;rm(dat0) ;rm(spop2)
dat1 <- merge(dat1, spop1, all.x = TRUE)
dat1 <- merge(dat1, outcome_febr, all.x = TRUE)

dat1$KON <- dat1$KON.y; dat1$KON.y <- NULL
dat1$lpnr_mor <- dat1$lpnr_mor.x 
dat1$lpnr_mor.x <- NULL
dat1$lpnr_mor.y <- NULL

dat2 <- dat1[,c("lpnr_barn","lpnr_mor","Diff", "MVIKT", "MLANGD", "MALDER", 
                "ROK0", "ROK1", "SNUS0", "SNUS1", "BFODDAT","FAMSIT_EJ_SAMBO", "ARBETE_EJ" , febrVars, spop2vars, spop1vars), with= FALSE]
#-------------------------------------------------------------------------------
dat2$DKLASS <- as.numeric(as.character(dat2$DKLASS))
dat2$DKLASS[is.na(dat2$DKLASS)] <- 0

dat2$DKLASS_3kl <- ifelse(dat2$DKLASS == 1, "1. Stillborn",
                          ifelse(dat2$DKLASS == 2, "2. Early neonatal death",
                                 ifelse(dat2$DKLASS == 3, "3. Late neonatal death","0. Ej dodfodd")))

dat2$perinatalDeath <- ifelse(dat2$DKLASS == 1 | dat2$DKLASS == 2, 1, 0)
dat2$perinatalDeath[is.na(dat2$perinatalDeath)] <- 0

dat2$neonatalDeath <- ifelse(dat2$DKLASS >= 2, 1, 0)
dat2$neonatalDeath[is.na(dat2$neonatalDeath)] <- 0

dat2$DODFOD <- as.numeric(as.character(dat2$DODFOD))
dat2$DODFOD[is.na(dat2$DODFOD)] <- 0

dat2$DODFOD <- ifelse(dat2$DODFOD == 1, "1. Fore forlossning", 
                      ifelse(dat2$DODFOD == 2, "2. Under forlossning", "0. Ej dodfodd")
                      )


dat2$stillborn <- ifelse(dat2$DKLASS == 1, 1, 0)

dat2$intrauterin_fdod <- ifelse(dat2$DODFOD  == "1. Fore forlossning",1,0)

dat2$UL_Forvantad_vikt <- ifelse(dat2$KON==1,
                          -(1.907345*10**(-6))*dat2$ULGravl**4 +
                            (1.140644*10**(-3))*dat2$ULGravl**3-
                            1.336265*10**(-1)*dat2$ULGravl**2+
                            1.976961*10**(0)*dat2$ULGravl+
                            2.410053*10**(2),
                          -(2.761948*10**(-6))*dat2$ULGravl**4+
                            (1.744841*10**(-3))*dat2$ULGravl**3-
                            2.893626*10**(-1)*dat2$ULGravl**2+
                            1.891197*10**(1)*dat2$ULGravl-
                            4.135122*10**(2)
  )

dat2$SM_Forvantad_vikt <- ifelse(dat2$KON==1,
                                 -(1.907345*10**(-6))*dat2$SMGravl**4 +
                                   (1.140644*10**(-3))*dat2$SMGravl**3-
                                   1.336265*10**(-1)*dat2$SMGravl**2+
                                   1.976961*10**(0)*dat2$SMGravl+
                                   2.410053*10**(2),
                                 -(2.761948*10**(-6))*dat2$SMGravl**4+
                                   (1.744841*10**(-3))*dat2$SMGravl**3-
                                   2.893626*10**(-1)*dat2$SMGravl**2+
                                   1.891197*10**(1)*dat2$SMGravl-
                                   4.135122*10**(2)
)

dat2$UL_Viktavvikelse <- ((dat2$BVIKTBS- dat2$UL_Forvantad_vikt)/dat2$UL_Forvantad_vikt)*100
dat2$SM_Viktavvikelse <- ((dat2$BVIKTBS- dat2$SM_Forvantad_vikt)/dat2$SM_Forvantad_vikt)*100

dat2p <- subset(dat2, KON == 1)
dat2f <- subset(dat2, KON == 2)


# KOn = 1
dat2p$SGA10_UL <- ifelse(dat2p$BVIKTBS < quantile(dat2p$UL_Forvantad_vikt, probs = 0.10),1,0)
dat2p$SGA10_SM <- ifelse(dat2p$BVIKTBS < quantile(dat2p$SM_Forvantad_vikt, probs = 0.10),1,0)
dat2p$SGA2SD_UL <- ifelse(dat2p$BVIKTBS < dat2p$UL_Forvantad_vikt - 0.24*mean(dat2p$UL_Forvantad_vikt),1,0)
dat2p$SGA2SD_SM <- ifelse(dat2p$BVIKTBS < dat2p$SM_Forvantad_vikt - 0.24*mean(dat2p$SM_Forvantad_vikt),1,0)
dat2p$LGA90_UL <- ifelse(dat2p$BVIKTBS > quantile(dat2p$UL_Forvantad_vikt, probs = 0.90),1,0)
dat2p$LGA90_SM <- ifelse(dat2p$BVIKTBS > quantile(dat2p$SM_Forvantad_vikt, probs = 0.90),1,0)
dat2p$LGA2SD_UL <- ifelse(dat2p$BVIKTBS > dat2p$UL_Forvantad_vikt + 0.24*mean(dat2p$UL_Forvantad_vikt),1,0)
dat2p$LGA2SD_SM <- ifelse(dat2p$BVIKTBS > dat2p$SM_Forvantad_vikt + 0.24*mean(dat2p$SM_Forvantad_vikt),1,0)

# dat2f$SD <- 

dat2f$SGA10_UL <- ifelse(dat2f$BVIKTBS < quantile(dat2f$UL_Forvantad_vikt, probs = 0.10),1,0)
dat2f$SGA10_SM <- ifelse(dat2f$BVIKTBS < quantile(dat2f$SM_Forvantad_vikt, probs = 0.10),1,0)
dat2f$SGA2SD_UL <- ifelse(dat2f$BVIKTBS < dat2f$UL_Forvantad_vikt - 0.24*mean(dat2f$UL_Forvantad_vikt),1,0)
dat2f$SGA2SD_SM <- ifelse(dat2f$BVIKTBS < dat2f$SM_Forvantad_vikt - 0.24*mean(dat2f$SM_Forvantad_vikt),1,0)
dat2f$LGA90_UL <- ifelse(dat2f$BVIKTBS > quantile(dat2f$UL_Forvantad_vikt, probs = 0.90),1,0)
dat2f$LGA90_SM <- ifelse(dat2f$BVIKTBS > quantile(dat2f$SM_Forvantad_vikt, probs = 0.90),1,0)
dat2f$LGA2SD_UL <- ifelse(dat2f$BVIKTBS > dat2f$UL_Forvantad_vikt + 0.24*mean(dat2f$UL_Forvantad_vikt),1,0)
dat2f$LGA2SD_SM <- ifelse(dat2f$BVIKTBS > dat2f$SM_Forvantad_vikt + 0.24*mean(dat2f$SM_Forvantad_vikt),1,0)

dat2 <- bind_rows(dat2p,dat2f)

dat2 <- data.frame(dat2)

#recode factors
dat2$MALE <- ifelse(as.numeric(as.character(dat2$KON))==1, 1, 0)
dat2$MSGA <- as.numeric(as.character(dat2$MSGA))
dat2$MLGA <- as.numeric(as.character(dat2$MLGA))

saveRDS(dat2, "Output/get_data.rds")

