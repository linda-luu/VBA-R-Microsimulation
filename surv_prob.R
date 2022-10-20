# Get survival probabilities given parameter estimates from Excel 

args      <-commandArgs(trailingOnly=T)
file_path <- args[1]

setwd(file_path)   # CHANGE TO FILE PATH  

# Load packages
if (!require('pacman')) install.packages('pacman'); library(pacman) # use this package to conveniently install other packages
# load (install if required) packages from CRAN
p_load( "dplyr", "flexsurv", "ggplot2", "openxlsx", "dampack", "matrixStats")    

# load (install if required) packages from GitHub
#install_github("DARTH-git/darthtools", force = TRUE) #Uncomment if there is a newer version
p_load_gh("DARTH-git/darthtools")

# Load Excel workbook
wb_path <- paste0(file_path, "/3_state.xlsm")
wb      <- loadWorkbook(wb_path)

# Load functions
source("fun_3_state.R")

####################
# Input parameters
####################
n_t    <- as.numeric(read.xlsx(wb, namedRegion = "n_years",    colNames = F)) # Number of time cycles (in years)
c_l    <- as.numeric(read.xlsx(wb, namedRegion = "cy_len",     colNames = F)) # Cycle length
times  <- seq(from = 0, to = n_t, by = c_l)                                   # Vector of time points

############################
# Input parameters: Survival
############################
# Load parameters and covariance matrix of survival distributions for different transitions 
df_HS <- read.xlsx(wb, namedRegion = "HS", colNames = F) # Healthy to Sick
df_HD <- read.xlsx(wb, namedRegion = "HD", colNames = F) # Healthy to Dead
df_SD <- read.xlsx(wb, namedRegion = "SD", colNames = F) # Sick to Dead
# Assign appropriate column names
surv_colnames <- as.character(read.xlsx(wb, namedRegion = "surv_colnames", colNames = F))
names(df_HS)  <- names(df_HD) <- names(df_SD) <- surv_colnames

# Selected distribution from Excel
dist_HS <- as.character(read.xlsx(wb, namedRegion = "dist_HS", colNames = F)) # Healthy to Sick
dist_HD <- as.character(read.xlsx(wb, namedRegion = "dist_HD", colNames = F)) # Healthy to Dead
dist_SD <- as.character(read.xlsx(wb, namedRegion = "dist_SD", colNames = F)) # Sick to Dead

# Process Excel Data (survival distributions)
l_HS <- fun_model_list(df_HS) # Healthy to Sick
l_HD <- fun_model_list(df_HD) # Healthy to Dead
l_SD <- fun_model_list(df_SD) # Sick to Dead

# Initialize matrices to store survival probabilities
m_surv_HS <- m_surv_HD <- m_surv_SD <- matrix(nrow = length(times), ncol = 0)

# Calculate survival probabilities and store them in list
for (i in 1:length(l_HS)) {
  # Healthy to Sick
  dist_name_HS     <- names(l_HS)[i]        # Obtain name of distribution
  dist_df_HS       <- l_HS[[i]]$df          # Obtain distribution info (params, coeffs, SE, covariance matrix)
  n_dist_params_HS <- nrow(dist_df_HS)      # Obtain the number of parameters
  S_t              <- St(t      = times,    # Calculate survival probabilities
                         dist   = dist_name_HS, 
                         params = dist_df_HS$Coefficients)
  m_surv_HS        <- cbind(m_surv_HS, S_t) # store in matrix of survival probabilities
  
  # Healthy to Dead
  dist_name_HD     <- names(l_HD)[i]        # Obtain name of distribution
  dist_df_HD       <- l_HD[[i]]$df          # Obtain distribution info (params, coeffs, SE, covariance matrix)
  n_dist_params_HD <- nrow(dist_df_HD)      # Obtain the number of parameters
  S_t              <- St(t      = times,    # Calculate survival probabilities
                         dist   = dist_name_HD, 
                         params = dist_df_HD$Coefficients) 
  m_surv_HD        <- cbind(m_surv_HD, S_t) # store in matrix of  survival probabilities
  
  # Sick to Dead
  dist_name_SD     <- names(l_SD)[i]        # Obtain name of distribution
  dist_df_SD       <- l_SD[[i]]$df          # Obtain distribution info (params, coeffs, SE, covariance matrix)
  n_dist_params_SD <- nrow(dist_df_SD)      # Obtain the number of parameters
  S_t              <- St(t      = times,    # Calculate survival probabilities
                         dist   = dist_name_SD, 
                         params = dist_df_SD$Coefficients)
  m_surv_SD        <- cbind(m_surv_SD, S_t) # store in matrix of  survival probabilities
}

# Assign appropriate column names
colnames(m_surv_HS) <- colnames(m_surv_HD) <- colnames(m_surv_SD) <- names(l_HS)
m_surv <- cbind(times, m_surv_HS, m_surv_HD, m_surv_SD)    # Combine survival matrices for all health state transitions
write.table(m_surv, file = "surv_prob.txt", row.names = F) # Export combined matrix as .txt
