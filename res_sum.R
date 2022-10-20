# Simple 3 state R/Excel Model - Results
args <- commandArgs(trailingOnly = T) # variables loaded through shell from VBA  
file_path <- args[1] # Path to directory holding Rscripts, workbook and outputs 

model_name <- args[2] # Which outputs requested: R or Excel 

setwd(file_path)  # set working directory to file path

# Load workbook 
wb_path <- paste0(file_path, "/3_state.xlsm")
wb <- openxlsx:::loadWorkbook(wb_path)
names(wb) # list worksheets

# Load packages
library(openxlsx)
library(dampack)
library(dplyr)
library(matrixStats)

## Function not included in dampack for some reason


#' Calculate incremental cost-effectiveness ratios from a \code{psa} object.
#'
#' @description The mean costs and QALYs for each strategy in a PSA are used
#' to conduct an incremental cost-effectiveness analysis. \code{\link{calculate_icers}} should be used
#' if costs and QALYs for each strategy need to be specified manually, whereas \code{calculate_icers_psa}
#' can be used if mean costs and mean QALYs from the PSA are assumed to represent a base case scenario for
#' calculation of ICERS.
#'
#' Optionally, the \code{uncertainty} argument can be used to provide the 2.5th and 97.5th
#' quantiles for each strategy's cost and QALY outcomes based on the variation present in the PSA.
#' Because the dominated vs. non-dominated status and the ordering of strategies in the ICER table are
#' liable to change across different samples of the PSA, confidence intervals are not provided for the
#' incremental costs and QALYs along the cost-effectiveness acceptability frontier.
#' \code{link{plot.psa}} does not show the confidence intervals in the resulting plot
#' even if present in the ICER table.
#'
#' @param psa \code{psa} object from \code{link{make_psa_object}}
#' @param uncertainty whether or not 95% quantiles for the cost and QALY outcomes should be included
#' in the resulting ICER table. Defaults to \code{FALSE}.
#'
#' @return A data frame and \code{icers} object of strategies and their associated
#' status, cost, effect, incremental cost, incremental effect, and ICER. If \code{uncertainty} is
#' set to \code{TRUE}, four additional columns are provided for the 2.5th and 97.5th quantiles for
#' each strategy's cost and effect.
#' @seealso \code{\link{plot.icers}}
#' @seealso \code{\link{calculate_icers}}
#' @importFrom tidyr pivot_longer
#' @export
calculate_icers_psa <- function(psa, uncertainty = FALSE) {
  
  # check that psa has class 'psa'
  dampack:::check_psa_object(psa)
  
  # Calculate mean outcome values
  psa_sum <- summary(psa)
  
  # Supply mean outcome values to calculate_icers
  icers <- calculate_icers(cost = psa_sum$meanCost,
                           effect = psa_sum$meanEffect,
                           strategies = psa_sum$Strategy)
  
  if (uncertainty == TRUE) {
    
    # extract cost and effect data.frames from psa object
    cost <- psa$cost
    effect <- psa$effectiveness
    
    # Calculate quantiles across costs and effects
    cost_bounds <- cost %>%
      pivot_longer(cols = everything(), names_to = "Strategy") %>%
      group_by(.data$Strategy) %>%
      summarize(Lower_95_Cost = quantile(.data$value, probs = 0.025, names = FALSE),
                Upper_95_Cost = quantile(.data$value, probs = 0.975, names = FALSE))
    
    effect_bounds <- effect %>%
      pivot_longer(cols = everything(), names_to = "Strategy") %>%
      group_by(.data$Strategy) %>%
      summarize(Lower_95_Effect = quantile(.data$value, probs = 0.025, names = FALSE),
                Upper_95_Effect = quantile(.data$value, probs = 0.975, names = FALSE))
    
    # merge bound data.frames into icers data.frame
    icers <- icers %>%
      left_join(cost_bounds, by = "Strategy") %>%
      left_join(effect_bounds, by = "Strategy") %>%
      select(.data$Strategy, .data$Cost, .data$Lower_95_Cost, .data$Upper_95_Cost,
             .data$Effect, .data$Lower_95_Effect, .data$Upper_95_Effect,
             .data$Inc_Cost, .data$Inc_Effect, .data$ICER, .data$Status)
  }
  
  return(icers)
}

###############################
# Read microsim outputs from workbook
###############################
results       <- (openxlsx:::read.xlsx(xlsxFile =  wb, namedRegion = paste0("results_", model_name), colNames = T,))
results_trial <- (openxlsx:::read.xlsx(xlsxFile =  wb, namedRegion = paste0("results_", model_name, "_trial"), colNames = T,)) 
results_ext   <- (openxlsx:::read.xlsx(xlsxFile =  wb, namedRegion = paste0("results_", model_name, "_ext"),   colNames = T,)) 

name_groups   <- (openxlsx:::read.xlsx(xlsxFile =  wb, namedRegion = "name_groups", colNames = F,)) # load vector of groups 
max_wtp       <- (openxlsx:::read.xlsx(xlsxFile =  wb, namedRegion = "ceiling_ratio", colNames = F,)) # Maximum willingness to pay 
max_wtp       <- as.numeric(max_wtp) # change from data frame to numeric vector 

df_eff        <- results[,c("effects", "trt.effects")] # dataframe with effects for different strategies
df_cost       <- results[,c("costs"  , "trt.costs")]   # dataframe with costs for different strategies

df_psa        <- make_psa_obj(effectiveness = df_eff, cost = df_cost, strategies = name_groups)

###############################
# Cost effectiveness Analysis 
###############################
# Calculate average costs, effects, incremental costs and increamental effects, and ICER for each strategy 
res_cea  <- calculate_icers_psa(df_psa)
colnames(res_cea) <- c("Strategy", "Cost", "QALYs", "Inc Costs", "Inc QALYs", "ICER", "Status")
write.table(res_cea, "res_cea.txt", row.names = F) # Save output to text 


# Calculate cost effectiveness acceptability curve based on maximum willingness to pay  
res_ceac <- ceac(psa = df_psa, wtp = seq(0, max_wtp, length.out =100))
res_ceac <- dplyr::select(res_ceac,!On_Frontier)

res_ceac_wide <- reshape(res_ceac, idvar = "WTP", timevar = "Strategy", direction = "wide")
write.table(res_ceac_wide,"res_ceac_wide.txt", row.names = F) # Save output to text

# Calculate cost effectiveness plane 
CEP <- data.frame(IC = df_cost$trt.costs - df_cost$costs,
                  IE = df_eff$trt.effects - df_eff$effects)
write.table(CEP, "res_cep.txt", row.names = F) # save output to text 


###############################
# Calculate mean outputs 
###############################
final_outputs <- as.data.frame(rbind(colMeans(results),
                                    colQuantiles(as.matrix(results), probs = 0.025),
                                    colQuantiles(as.matrix(results), probs = 0.975),
                                    colMeans(results_trial),
                                    colQuantiles(as.matrix(results_trial), probs = 0.025),
                                    colQuantiles(as.matrix(results_trial), probs = 0.975),
                                    colMeans(results_ext),
                                    colQuantiles(as.matrix(results_ext), probs = 0.025),
                                    colQuantiles(as.matrix(results_ext), probs = 0.975)))
rownames(final_outputs) <- c("Overall",  "Overall lower 95%", "Overall upper 95%",
                            "In Trial", "Trial lower 95%", "Trial upper 95%", 
                            "Extrapolated", "Extrapolated lower 95%", "Extrapolated upper 95%")
final_outputs <- round(final_outputs,3)
write.table(final_outputs, file = "final_outputs.txt", row.names = T)
