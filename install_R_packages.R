#Install R packages required for 3_state.xlsm 


print("loading github functions")

if (!require('pacman')) install.packages('pacman'); library(pacman) # use this package to conveniently install other packages
# load (install if required) packages from CRAN
p_load("devtools", "dplyr", "flexsurv", "ggplot2", "openxlsx", "dampack", "matrixStats")    

# load (install if required) packages from GitHub
#install_github("DARTH-git/darthtools", force = TRUE) #Uncomment if there is a newer version
p_load_gh("DARTH-git/darthtools")



args      <-commandArgs(trailingOnly=T)
file_path <- args[1]

setwd("C:/Users/LuuLi/Documents/CADTH ITP/analysis/excel")

write.table(file_path, file = "test.txt")

