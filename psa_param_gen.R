# This R script reads in data from 3_state.xlsm for survival distributions and generates parameter estimates

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


####################
# Input parameters
####################
n_sim  <- as.numeric  (read.xlsx(wb, namedRegion = "n_sims",   colNames = F)) # Number of simulations
psa    <- as.character(read.xlsx(wb, namedRegion = "psa",      colNames = F)) # Probabilistic Analysis? (Y/N)


############################
# Input parameters: Survival
############################
# load parameters and covariance matrix of survival distributions for different transitions 
df_HS   <- read.xlsx(wb, namedRegion = "HS", colNames = F) # Healthy to Sick
df_HD   <- read.xlsx(wb, namedRegion = "HD", colNames = F) # Healthy to Dead
df_SD   <- read.xlsx(wb, namedRegion = "SD", colNames = F) # Sick to Dead
# Assign appropriate column names
surv_colnames <- as.character(read.xlsx(wb, namedRegion = "surv_colnames", colNames = F))
names(df_HS)  <- names(df_HD) <- names(df_SD) <- surv_colnames

dist_HS <- as.character(read.xlsx(wb, namedRegion = "dist_HS", colNames = F)) # Healthy to Sick
dist_HD <- as.character(read.xlsx(wb, namedRegion = "dist_HD", colNames = F)) # Healthy to Dead
dist_SD <- as.character(read.xlsx(wb, namedRegion = "dist_SD", colNames = F)) # Sick to Dead


##############################
# Load Functions & run
##############################
source("fun_3_state.R") # load functions

# Process Excel Data (survival distributions)
l_HS <- fun_model_list(df_HS) # Healthy to Sick
l_HD <- fun_model_list(df_HD) # Healthy to Dead
l_SD <- fun_model_list(df_SD) # Sick to Dead


#########################################
# Generate parameters for PSA runs: Surv 
#########################################
if(psa == "N"){  # If not conducting PSA
  # Use deterministic values for parameters of survival distributions
  m_coef_HS  <- matrix(rep(l_HS[[dist_HS]]$df$Coefficients, n_sim), nrow = n_sim, byrow = T) # Healthy to Sick
  m_coef_HD  <- matrix(rep(l_HD[[dist_HD]]$df$Coefficients, n_sim), nrow = n_sim, byrow = T) # Healthy to Dead
  m_coef_SD  <- matrix(rep(l_SD[[dist_SD]]$df$Coefficients, n_sim), nrow = n_sim, byrow = T) # Sick to Dead
}else if(psa == "Y"){  # If conducting PSA
  # Sample parameters of survival distributions from multivariate normal distributions
  m_coef_HS <- mvtnorm::rmvnorm(n_sim, l_HS[[dist_HS]]$df$Coefficients, l_HS[[dist_HS]]$vcov, checkSymmetry = F) # Healthy to Sick
  m_coef_HD <- mvtnorm::rmvnorm(n_sim, l_HD[[dist_HD]]$df$Coefficients, l_HD[[dist_HD]]$vcov, checkSymmetry = F) # Healthy to Dead
  m_coef_SD <- mvtnorm::rmvnorm(n_sim, l_SD[[dist_SD]]$df$Coefficients, l_SD[[dist_SD]]$vcov, checkSymmetry = F) # Sick to Dead
}
# Assign appropriate column names
colnames(m_coef_HS) <- l_HS[[dist_HS]]$df$Parameters 
colnames(m_coef_HD) <- l_HS[[dist_HD]]$df$Parameters
colnames(m_coef_SD) <- l_HS[[dist_SD]]$df$Parameters

## Save Outputs to txt files
write.table(m_coef_HS, file="probabilistic_params_HS.txt", row.names = F) # Healthy to Sick
write.table(m_coef_HD, file="probabilistic_params_HD.txt", row.names = F) # Health to Dead
write.table(m_coef_SD, file="probabilistic_params_SD.txt", row.names = F) # Sick to Dead


############################
# Input parameters: Cost
############################
cost.info <- read.xlsx(wb, namedRegion = "costs", colNames = F)
colnames(cost.info) <- c("name", "mean", "sd", "dist")

#########################################
# Generate costs for PSA  
#########################################
m_costs <- matrix(nrow = n_sim, ncol = 0)

if (psa == "Y"){ # If conducting PSA
  
  for (i in 1:nrow(cost.info)){ # Loop through each cost parameter
    row <- cost.info[i,]
    
    if(is.na(row$dist)){                                           # If the cost parameter does not have a distribution
      temp_costs  <- rep(row$mean, n_sim)                          # Use the static mean value for all samples
    } else if(row$dist == "gamma"){                                # If the cost distribution is Gamma
       gam        <- gamma_params(row$mean, row$sd, scale = FALSE) # Obtain corresponding Gamma distribution parameters 
       temp_costs <- rgamma(n_sim, gam$shape, gam$rate)            # Sample from corresponding Gamma distribution
    } else if(row$dist == "beta"){                                 # If the cost distribution is Beta
       beta       <- beta_params(row$mean, row$sd)                 # Obtain corresponding Beta distribution parameters 
       temp_costs <- rbeta(n_sim, beta$alpha, beta$beta)           # Sample from corresponding Beta distribution
    } else if(row$dist == "normal"){                               # If the cost distribution is Normal
       temp_costs <- rnorm(n_sim, mean = row$mean, sd = row$sd)    # Sample from corresponding Normal distribution
    }
    
    m_costs <- cbind(m_costs, temp_costs)
  }
  colnames(m_costs) <- cost.info$name
  
}else if(psa == "N"){ # If not conducting PSA
  
  for (i in 1:nrow(cost.info)){
    m_costs <- cbind(m_costs, rep(cost.info[i,]$mean, n_sim)) # Use deterministic cost parameter values
  }
  colnames(m_costs) <- cost.info$name
}

write.table(m_costs, file = "costs_input.txt", row.names = F)
