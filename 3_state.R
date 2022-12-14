# Simple 3 state R/Excel Model
args <- commandArgs(trailingOnly=T)
file_path <- args[1]

setwd(file_path)  # CHANGE TO FILE PATH  

####################
# Load Packages
####################
# use this package to conveniently install other packages
if (!require('pacman')) install.packages('pacman'); library(pacman)

# load (install if required) packages from CRAN
p_load( "dplyr", "flexsurv", "ggplot2", "openxlsx", "dampack", "matrixStats")    

# load (install if required) packages from GitHub
p_load_gh("DARTH-git/darthtools")

############################
# Load Excel Workbook to wb
############################
wb_path <- paste0(file_path, "/3_state.xlsm") # path to workbook
wb      <- loadWorkbook(wb_path)              # Open workbook and save to wb variable

####################
# Input Functions
####################
source("fun_3_state.R") # Run fun_3_state.R Script to load in functions

####################
# Input parameters
####################
n_sim  <- as.numeric(read.xlsx(wb, namedRegion = "n_sims",     colNames = F))    # Number of simulations
n_ind  <- as.numeric(read.xlsx(wb, namedRegion = "n_i",        colNames = F))    # Number of individuals
n_t    <- as.numeric(read.xlsx(wb, namedRegion = "n_years",    colNames = F))    # Number of time cycles (in years)
n_s    <- as.numeric(read.xlsx(wb, namedRegion = "n_s",        colNames = F))    # Number of states
init_s <- as.numeric(read.xlsx(wb, namedRegion = "state_init", colNames = F))    # Initial states
c_l    <- as.numeric(read.xlsx(wb, namedRegion = "cy_len",     colNames = F))    # Cycle length
times  <- seq(from = 0, to = n_t, by = c_l)                                      # Vector of cycle times
n_cy   <- length(times)                                                          # Number of cycles

trial_t<- as.numeric(read.xlsx(wb, namedRegion = "n_yr_t",     colNames = F))    # Number of years of trial
surv_t <- as.numeric(read.xlsx(wb, namedRegion = "surv_t",     colNames = F))    # Present survival at time:   

p_adv  <- as.numeric(read.xlsx(wb, namedRegion = "p_adverse",  colNames = F))    # Probability of AE

v_s    <- as.character(read.xlsx(wb, namedRegion = "name_states", colNames = F)) # Vector of states
v_grp  <- as.character(read.xlsx(wb, namedRegion = "name_groups", colNames = F)) # Vector of group names

psa    <- as.character(read.xlsx(wb, namedRegion = "psa", colNames = F))         # Probabilistic Analysis? (Y/N)          

dwc    <- as.numeric(read.xlsx(wb, namedRegion = "dwc", colNames = F))           # Discounting rate for costs
dwe    <- as.numeric(read.xlsx(wb, namedRegion = "dwe", colNames = F))           # Discounting rate for effects
v_dc   <- 1 / ((1 + dwc) ^ (times))                                              # Vector of discount weights for costs
v_de   <- 1 / ((1 + dwe) ^ (times))                                              # Vector of discount weights for utilities


# Ceiling Ratio (willingness to pay)
ceiling_ratio <- read.xlsx(wb, namedRegion = "ceiling_ratio", colNames = F)

###########################################
# Input parameters: Surv
## psa params generated by psa_param_gen.R
###########################################

# Import sampled survival distribution coefficients from Excel
m_coef_HS <- as.matrix(   read.xlsx(wb, namedRegion = "psa_HS",  colNames = F)) # Healthy to Sick
m_coef_HD <- as.matrix(   read.xlsx(wb, namedRegion = "psa_HD",  colNames = F)) # Healthy to Dead
m_coef_SD <- as.matrix(   read.xlsx(wb, namedRegion = "psa_SD",  colNames = F)) # Sick to Dead

# Import name of selected distribution from Excel
dist_HS   <- as.character(read.xlsx(wb, namedRegion = "dist_HS", colNames = F)) # Healthy to Sick
dist_HD   <- as.character(read.xlsx(wb, namedRegion = "dist_HD", colNames = F)) # Health to Dead
dist_SD   <- as.character(read.xlsx(wb, namedRegion = "dist_SD", colNames = F)) # Sick to Dead

#########################################
# Input parameters: Costs
#########################################
m_costs    <- read.xlsx(wb, namedRegion = "m_cost", colNames = F)
cost_names <- m_costs[1,]
m_costs    <- m_costs[-1,]
m_costs    <- matrix(as.numeric(as.matrix(m_costs)), nrow = n_sim)
colnames(m_costs) <- cost_names

#########################################
# Organize Costs in matrices for SoC and Trt  
## Rows   : Simulation    
## Columns: State
#########################################
m_costs_soc <- matrix(c(rowSums(m_costs[,c("cost_PF_acq_soc", "cost_PF_admin_soc", "cost_PF_soc")]), # Healthy  
                        rowSums(m_costs[,c("cost_P_acq_soc", "cost_P_soc")]),                        # Sick
                        rep(0, n_sim)),                                                              # Dead
                      nrow = n_sim)

m_costs_trt <- matrix(c(rowSums(m_costs[,c("cost_PF_acq_trt", "cost_PF_admin_trt", "cost_PF_trt")]), # Healthy
                        rowSums(m_costs[,c("cost_P_acq_trt", "cost_P_trt")]),                        # Sick
                        rep(0, n_sim)),                                                              # Dead
                      nrow = n_sim)

# Costs for adverse events
costs_adv <- m_costs[,"cost_adv"]

#########################################
# Import Utilities
#########################################
u_H   <- as.numeric(read.xlsx(wb, namedRegion = "util_H", colNames = F))
u_S   <- as.numeric(read.xlsx(wb, namedRegion = "util_S", colNames = F))
u_D   <- as.numeric(read.xlsx(wb, namedRegion = "util_D", colNames = F))
u_trt <- as.numeric(read.xlsx(wb, namedRegion = "util_T", colNames = F))

v_utils     <- c(u_H, u_S, u_D) 
v_utils_trt <- c(u_H, u_trt, u_D) 

####################
# Initialize vectors 
####################
v_M_init  <- rep(v_s[init_s], times = n_ind)    # Vector state individuals start in 
v_Ts_init <- ifelse(v_M_init == v_s[2],c_l, 0)  # Vector with the time of being Sick at the start of the model 

######################################
# Run microsim and write to txt files 
######################################

msm_out <- MicroSim() # Run microsimulation

write.table(msm_out$msm_trace,     file = "Microsim_trace_R.txt",    row.names = F)  # Export state distributions (trace)
write.table(msm_out$outputs,       file = "Microsim_outputs_R.txt",  row.names = F)  # Export outputs (costs and effectiveness estimates for all simulations)
write.table(msm_out$surv,          file = "Microsim_survival_R.txt", row.names = F)  # Export survival probabilities
write.table(msm_out$extrapolation, file = "Microsim_ext_R.txt",      row.names = F)  # Export extrapolated outputs (costs and effectiveness estimates for all simulations)
write.table(msm_out$trial,         file = "Microsim_trial_R.txt",    row.names = F)  # Export trial outputs (costs and effectiveness estimates for all simulations)

