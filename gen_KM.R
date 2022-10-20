# Generate data based on parameters from a distribution
# args      <-commandArgs(trailingOnly=T)
# file_path <- args[1]
# 

file_path <- dirname(rstudioapi::getActiveDocumentContext()$path) # Path to excel file + R scripts 
setwd(file_path)   # CHANGE TO FILE PATH  

# Load packages
if (!require('pacman')) install.packages('pacman'); library(pacman) # use this package to conveniently install other packages
# load (install if required) packages from CRAN
p_load( "dplyr", "flexsurv", "ggplot2", "openxlsx", "dampack", "matrixStats", "survminer")    

# load (install if required) packages from GitHub
#install_github("DARTH-git/darthtools", force = TRUE) #Uncomment if there is a newer version
p_load_gh("DARTH-git/darthtools")

# Load Excel workbook
wb_path <- paste0(file_path, "/3_state.xlsm")
wb      <- loadWorkbook(wb_path)

# Load functions
source("fun_3_state.R")

############################
# Input parameters: Survival
############################
# Load parameters and covariance matrix of survival distributions for different transitions 
df_trans <- read.xlsx(wb, namedRegion = "HS", colNames = F) 

# Assign appropriate column names
surv_colnames <- as.character(read.xlsx(wb, namedRegion = "surv_colnames", colNames = F))
names(df_trans)  <- surv_colnames

# Selected distribution from Excel
dist <- as.character(read.xlsx(wb, namedRegion = "dist_HS", colNames = F)) # Healthy to Sick

# Process Excel Data (survival distributions)
l_models <- fun_model_list(df_trans)

l_models[[dist]]

#####################################
# Input parameters: Generate Dataset
#####################################
df_gen <- data.frame(time = rweibull(n = 1000, l_models[[dist]]$df$Coefficients),
                     status = 0)
df_gen <- df_gen %>% mutate(status = case_when(time > median(df_gen$time) ~ 0,
                                               TRUE ~ 1),
                            time = case_when(time > median(df_gen$time) ~ median(df_gen$time),
                                             TRUE ~ time))

surv_gen <- survfit(Surv(time, status)~1, data = df_gen)

pdf(file = "KMplot.pdf", onefile = FALSE)
ggsurvplot(surv_gen, conf.int = FALSE, risk.table = TRUE, legend = "none", ggtheme = theme_bw(), palette = "black") 
dev.off()
