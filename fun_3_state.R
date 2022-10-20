# R functions used in 3_state.xlsm



#' fun_model_list
#'
#' @param df_model dataframe of distributions, parameters and covariance matrices from Excel
#'
#' @return List of list for distributions, storing data frame of parameters and covariance matrix from each 

fun_model_list <- function(df_model){
  l_outputs <- list()
  l_split   <- split(df_model, df_model$Function)
  
  cov_cols  <- names(df_model)[grepl("Cov", names(df_model), fixed = T)]
  
  for (i in 1:length(l_split)){
    element <- l_split[[i]]
    
    covs <- element[, cov_cols]  # Extract covariance matrix
    list_model <- list(df   = element,
                       vcov = matrix(covs[!is.na(covs)], ncol = nrow(covs)))
    
    l_outputs[[i]] <- list_model
  }
  
  names(l_outputs) <- names(l_split)
  
  return(l_outputs)
}



# Function to calculate survival probabilities with coefficients
#' St
#'
#' @param t Vector of time points to calculate survival probabilities at
#' @param dist Name of survival distribution 
#' @param params Vector of survival distribution parameters
#'
#' @return A vector of survival probabilities
#' 
St <- function(t, dist, params) {
  param1 <- params[1]; param2 <- params[2]; param3 <- params[3]
  if (dist == "Exponential") {
    S_t <- 1 - pexp(t, rate = exp(param1))
  } else if (dist == "Gamma") {
    S_t <- 1 - pgamma(t, shape = exp(param1), rate = exp(param2))
  } else if (dist == "Generalised Gamma") {
    S_t <- 1 - pgengamma(t, mu = param1, sigma = exp(param2), q = param3)
  } else if (dist == "Gompertz") {
    S_t <- 1 - pgompertz(t, shape = param1, rate = exp(param2))
  } else if (dist == "Log-logistic") {
    S_t <- 1 - pllogis(t, shape = exp(param1), scale = exp(param2))
  } else if (dist == "Lognormal") {
    S_t <- 1 - plnorm(t, meanlog = param1, sdlog = exp(param2))
  } else if (dist == "Weibull") {
    S_t <- 1 - pweibull(t, shape = exp(param1), scale = exp(param2))
  }
  return(S_t)
}



#' transition probabilities
#'
#' @param dist.v character: model distribution
#' @param d.data vector: model estimates (flexsurvreg(Surv(t, event) ~ x1 + x2 + x3, dist = "...")$res[,1])
#' @param dat.x  matrix: individuals covariates
#' @param t      vector: time spent in state 
#' @param model  character: name of model
#' @param step   numeric: cycle length
#'
#' @return transition probability for this cycle, for each individual 
#' 
model.dist.f <- function(dist.v = NA, d.data = NA, dat.x = NA, t = NA, trans = NA, step = c_l){
  
  # Get cumulative survival probability for each individual
  ############################ Gamma ############################
  if(dist.v == "Gamma"){
    
    pars <- d.data[1:2]
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {                             
      if(i == 2){ 
        
        # beta affects the rate
        if(length(beta) == 0){beta.raw = 0 # if no covariates
        }else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- (pars[i]) + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- exp(pred.raw)           # fit$dlist$inv.transforms
      } 
      else{print("gamma")}
    }
    p.res   <- pgamma(t,        shape = pars[1], rate = pred, lower.tail = F)
    p.res.1 <- pgamma(t - step, shape = pars[1], rate = pred, lower.tail = F)
    
    ############################ Exponential ############################
  } else if (dist.v == "Exponential") {
    
    pars <- d.data[1]
    beta <- if(length(d.data) > 1){d.data[2:length(d.data)]} else{NA} 
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {
      
      # beta affects the rate
      pars <- d.data[1]
      beta <- if(length(d.data) > 1){d.data[2:length(d.data)]} else{NA} 
      beta <- beta[!is.na(beta)]
      
      if(length(beta) ==0){beta.raw = 0}
      else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
      
      pred.raw <- (pars[i]) + beta.raw    # fit$dlist$transforms + beta.raw
      pred     <- exp(pred.raw)           # fit$dlist$inv.transforms
    }
    p.res   <- pexp(t,        rate = pred, lower.tail = F)
    p.res.1 <- pexp(t - step, rate = pred, lower.tail = F)
    
    ############################ Weibull ############################
  } else if (dist.v == "Weibull") {
    
    pars <- exp(d.data[1:2])
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {
      if(i == 2) { 
        # Beta affects the scale
        
        if(length(beta) ==0){beta.raw = 0}
        else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- log(pars[i]) + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- exp(pred.raw)              # fit$dlist$inv.transforms
        
      } 
      else{}
    }
    p.res   <- pweibull(t,        shape = pars[1], scale = pred, lower.tail = F)
    p.res.1 <- pweibull(t - step, shape = pars[1], scale = pred, lower.tail = F)
    
    ############################ LogNormal ############################
  } else if (dist.v == "Lognormal") {
    
    pars <- d.data[1:2]
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {
      if(i == 1) {   
        #meanlog
        
        if(length(beta) ==0){beta.raw = 0}
        else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- pars[i] + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- pred.raw              # fit$dlist$inv.transforms
      }
      else{}
    }
    p.res   <- plnorm(t,        meanlog = pred, sdlog = exp(pars[2]), lower.tail = F)
    p.res.1 <- plnorm(t - step, meanlog = pred, sdlog = exp(pars[2]), lower.tail = F)
    
    ############################ Gompertz ############################
  } else if (dist.v == "Gompertz") {   
    
    pars <- d.data[1:2]
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {
      if(i == 2) {
        
        #rate
        if(length(beta) ==0){beta.raw = 0}
        else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- pars[i] + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- exp(pred.raw)         # fit$dlist$inv.transforms
        
      } 
      else{}
    }
    p.res   <- pgompertz(t,        shape = pars[1], rate = pred, lower.tail = F)
    p.res.1 <- pgompertz(t - step, shape = pars[1], rate = pred, lower.tail = F)
    
    ############################ LogLogistic ############################
  } else if (dist.v == "Log-logistic") {   
    
    pars <- exp(d.data[1:2])
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    
    for (i in 1:length(pars)) {
      if(i == 2){ 
        
        #Scale
        if(length(beta) ==0){beta.raw = 0}
        else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- log(pars[i]) + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- exp(pred.raw)              # fit$dlist$inv.transforms
      } 
      else{}
    }
    p.res   <- pllogis(t,        shape = pars[1], scale = pred, lower.tail = F)
    p.res.1 <- pllogis(t - step, shape = pars[1], scale = pred, lower.tail = F)
    
    ############################ Generalised Gamma ############################
  } else if (dist.v == "Generalised Gamma") {
    
    pars <- d.data[1:3]
    beta <- if(length(d.data) > 2){d.data[3:length(d.data)]} else{NA}  
    beta <- beta[!is.na(beta)]
    
    for (i in 1:length(pars)) {
      if(i == 1) {   
        
        if(length(beta) == 0){beta.raw = 0}
        else{beta.raw <- t(as.matrix(beta)) %*% t(dat.x)}
        
        pred.raw <- pars[i] + beta.raw    # fit$dlist$transforms + beta.raw
        pred     <- pred.raw              # fit$dlist$inv.transforms
      }
      else{}
    }
    
    p.res   <- pgengamma(t,        mu = pred, sigma = exp(pars[2]), q = pars[3], lower.tail = F)
    p.res1  <- pgengamma(t - step, mu = pred, sigma = exp(pars[2]), q = pars[3], lower.tail = F)
    
    ############################ No Distribution ############################
  } else { 
    print("no distribution found") 
    p.res <- NA
    p.res.1 <- NA
  }
  
  return(1 - p.res/p.res.1)
}



#' @param M_t   : health state occupied by at cycle t (character variable)
#' @param v_Ts  : vector with the duration of being Sick
#' @param t     : current cycle
#' @param model  character: name of model
#' @param step   numeric: cycle length
#'
#' @return transition probability for this cycle, for each individual 
#' 
Probs <- function(M_t, v_Ts, t, p_HS, p_HD, p_SD) { 
  
  # create matrix of state transition probabilities
  m_p_t           <- matrix(0, nrow = n_s, ncol = n_ind) 
  # give the state names to the rows
  rownames(m_p_t) <-  v_s
  
  
  # update m_p_t with the appropriate probabilities   
  # (all transition probabilities are conditional on survival)
  # transition probabilities when Healthy 
  m_p_t["Healthy", M_t == "Healthy"] <- (1 - p_HD[t]) * (1 - p_HS[t])
  m_p_t["Sick",    M_t == "Healthy"] <- (1 - p_HD[t]) *      p_HS[t]
  m_p_t["Dead",    M_t == "Healthy"] <-      p_HD[t]  
  
  # transition probabilities when Sick 
  m_p_t["Healthy", M_t == "Sick"]    <- 0
  m_p_t["Sick",    M_t == "Sick"]    <- 1 - p_SD[v_Ts]
  m_p_t["Dead",    M_t == "Sick"]    <-     p_SD[v_Ts]
  
  # transition probabilities when Dead     
  m_p_t["Healthy", M_t == "Dead"]    <- 0                         
  m_p_t["Sick",    M_t == "Dead"]    <- 0
  m_p_t["Dead",    M_t == "Dead"]    <- 1
  
  return(t(m_p_t))
}  




#' Costs: calculates a vector of costs for a cycle
#'
#' @param v_M       vector of individuals states for cycle 
#' @param adverse   add adverse event costs? (T/F)
#' @param v_adverse vector of individuals adverse event status
#' @param trt       treatment group? (T/F)
#' @param k         simulation number 
#' @param dc        discounting factor this cycle 
#'
#' @return          vector of discounted costs
#' 
Costs <- function(v_M, adverse, v_adverse = v_adv, trt, k, dc = 1){
  
  # Individuals in each state
  pos.1 <- which(v_M == v_s[1])
  pos.2 <- which(v_M == v_s[2])
  pos.3 <- which(v_M == v_s[3])
  
  # Empty vector to store costs for this cycle
  v_c <- vector(mode = "numeric", length = length(v_M))
  
  options(warn=-1)
  
  # Costs based on intervention
  if (trt == F){
    v_c[pos.1] <- m_costs_soc[k,1]
    v_c[pos.2] <- m_costs_soc[k,2]
    v_c[pos.3] <- m_costs_soc[k,3]
  }else if(trt == T){
    v_c[pos.1] <- m_costs_trt[k,1]
    v_c[pos.2] <- m_costs_trt[k,2]
    v_c[pos.3] <- m_costs_trt[k,3]
  }
  
  # Add adverse costs if adverse events present
  if(adverse == T){
    v_c[v_adverse == 1] = v_c[v_adverse == 1] + costs_adv[k]
  }
  
  options(warn = 0)
  
  return(v_c * dc) # multiply by discount factor
}



#' Utils: Caclulate utilities for each individual for that cycle
#'
#' @param v_M  vector of individuals states for cycle
#' @param trt  treatment group? (T/F)
#' @param de   discounting factor for this cycle
#' 
#' @return     vector discounted utilities 
#'
Utils <- function(v_M, trt, de = 1){
  
  pos.1 <- which(v_M == v_s[1])
  pos.2 <- which(v_M == v_s[2])
  pos.3 <- which(v_M == v_s[3])
  
  v_u <- vector(mode = "numeric", length = length(v_M))
  
  if(trt == F){
    v_u[pos.1] <- v_utils[1]
    v_u[pos.2] <- v_utils[2]
    v_u[pos.3] <- v_utils[3]
  }else if(trt == T){
    v_u[pos.1] <- v_utils_trt[1]
    v_u[pos.2] <- v_utils_trt[2]
    v_u[pos.3] <- v_utils_trt[3]
  }
  return(v_u * de)
}


#' Microsim: Runs the microsimulation
#'
#' @param n_ind   number of individuals to simulate
#' @param n_sim   number of simulations to run  
#' @param seed    seed number for reproducible results 
#'
#' @return list of results containing: 
#'            msm_trace - matrix of cycle times with proportion of individuals in each health state
#'            outputs - matrix of life expectancy, costs, effects, treatment costs, treatment effects, for each simulation
#'            df_surv - data frame of survival probabilities and 95% confidence intervals at each cycle 
#'
MicroSim <- function(seed = 1) {
  set.seed(seed) # set the seed for reproducable results
  
  # Empty dataframes to store outputs
  outputs <- df_trial  <- data.frame(le = as.numeric(),               # life expectancy
                                     costs = as.numeric(),            # Costs for first intervention
                                     effects = as.numeric(),          # Effects for first intervention
                                     `trt costs` = as.numeric(),      # Costs for second intervention (trt)
                                     `trt effects` = as.numeric())    # Effects for second intervention (trt)
  
  # Matrix to store survival probabilities 
  m_surv <- matrix(NA, 
                   nrow = n_sim,
                   ncol = n_cy)
  
  
  
  # Start simulation 
  for(k in 1:n_sim){
    
    # m_M is used to store the health state information over time for every individual
    m_M  <-  matrix(nrow = n_ind, ncol = n_cy , 
                    dimnames = list(paste("ind" , 1:n_ind, sep = " "), 
                                    paste("year", times, sep = " ")))  
    
    # Store cost, utilities, treatment costs and treatment utilities for every individual
    m_c <- m_u <- m_c_trt <- m_u_trt <- m_M
    
    
    m_M[, 1] <- v_M_init                                  # initial health state for individual i
    v_Ts     <- v_Ts_init                                 # initialize time since illness onset for individual i
    v_adv    <- rbinom(n = n_ind, size = 1, prob = p_adv) # sample individuals who will experience adverse events
    v_dead   <- as.numeric(m_M[,1] == "Dead")             # vector storing whether individuals are dead 
    
    
    # costs and effects
    m_c[,1]     <- Costs(v_M = m_M[,1], adverse = T, v_adverse = v_adv, trt = F, k = k, dc = v_dc[1]) # costs for cycle 0 
    m_c_trt[,1] <- Costs(v_M = m_M[,1], adverse = T, v_adverse = v_adv, trt = T, k = k, dc = v_dc[1]) # costs for treatment in cycle 0 
    
    m_u[, 1]     <- Utils(v_M = m_M[,1], trt = F, de = v_de[1]) # effects for cycle 0 
    m_u_trt[, 1] <- Utils(v_M = m_M[,1], trt = T, de = v_de[1]) # effects for treatment in cycle 0 
    
    # Vector of transition probabilities 
    p_HS <- model.dist.f(dist.v = dist_HS, d.data = m_coef_HS[k,], t = times, step = c_l)[-1] # probability of transitioning from healthy to sick 
    p_HD <- model.dist.f(dist.v = dist_HD, d.data = m_coef_HD[k,], t = times, step = c_l)[-1] # probability of transitioning from healthy to dead 
    p_SD <- model.dist.f(dist.v = dist_SD, d.data = m_coef_SD[k,], t = times, step = c_l)[-1] # probability of transitioning from sick to dead 
    
    
    
    # open a loop for time, running cycles 1 to n_t 
    for (t in 1:(n_cy - 1)) {
      
      # calculate the transition probabilities for the cycle based on health state t
      m_p        <- Probs(m_M[, t], v_Ts, t, p_HS, p_HD, p_SD) 
      
      # sample next state based on transition probabilities in m_p
      m_M[,t+1]  <- samplev(m_p) 
      
      # costs and effects
      m_c[,t+1]     <- Costs(v_M = m_M[,t+1], adverse = F, v_adverse = v_adv, trt = F, k = k, dc = v_dc[t+1]) # costs for cycle
      m_c_trt[,t+1] <- Costs(v_M = m_M[,t+1], adverse = F, v_adverse = v_adv, trt = T, k = k, dc = v_dc[t+1])
      
      m_c    [(v_dead != 1 & m_M[,t+1] == "Dead"), t+1] <- m_c    [(v_dead != 1 & m_M[,t+1] == "Dead"), t+1] + m_costs[k,"cost_D"] * v_dc[t+1] # End of life costs for individuals that just died this cycle
      m_c_trt[(v_dead != 1 & m_M[,t+1] == "Dead"), t+1] <- m_c_trt[(v_dead != 1 & m_M[,t+1] == "Dead"), t+1] + m_costs[k,"cost_D"] * v_dc[t+1]

      v_dead <- as.numeric(m_M[, t+1] == "Dead") # update death vector
      
      m_u[,t+1]     <- Utils(v_M = m_M[,t+1], trt = F, de = v_de[t+1]) # effects for cycle
      m_u_trt[,t+1] <- Utils(v_M = m_M[,t+1], trt = T, de = v_de[t+1])
      
      # update time since illness onset for t + 1 
      v_Ts <- ifelse(m_M[, t + 1] == "Sick", v_Ts + 1, 0) 
      
    } # close the loop for the time points 
    
    
    # Process outputs 
    # 1. Microsimulation Trace 
      m_M_Micro <- t(apply(m_M, 2, function(x) table(factor(x, levels = v_s, ordered = TRUE)))) 
      m_M_Micro <- m_M_Micro / n_ind    # calculate the proportion of individuals 
      m_M_Micro <- cbind(times, m_M_Micro) # add column of cycle times 
    
      colnames(m_M_Micro) <- c("Cycle", v_s)    # name columns  
      rownames(m_M_Micro) <- paste("Cycle", times, sep = " ") # name rows  
    
    # 2. Calculate life expectancy, average costs and effects for each simulation
      # Outputs for overall simulation period
      le_Micro    <- sum(m_M_Micro[,2:3]) * c_l
      c_Micro     <- sum(m_c)/n_ind
      u_Micro     <- sum(m_u)/n_ind * c_l
      c_trt_Micro <- sum(m_c_trt)/n_ind 
      u_trt_Micro <- sum(m_u_trt)/n_ind * c_l
    
      # Outputs for trial period 
      le_trial    <- sum(m_M_Micro[1:(trial_t/c_l),][,2:3]) * c_l
      c_trial     <- sum(m_c    [,1:(trial_t/c_l)])/n_ind
      u_trial     <- sum(m_u    [,1:(trial_t/c_l)])/n_ind * c_l
      c_trt_trial <- sum(m_c_trt[,1:(trial_t/c_l)])/n_ind
      u_trt_trial <- sum(m_u_trt[,1:(trial_t/c_l)])/n_ind * c_l
    
    #2b. Organize outputs 
      outputs[nrow(outputs)+1,]   <- c(le_Micro, c_Micro, u_Micro, 
                                       c_trt_Micro, u_trt_Micro)
      
      
      df_trial[nrow(df_trial)+1,] <- c(le_trial, c_trial, u_trial, 
                                       c_trt_trial, u_trt_trial)
    
    
    # 3. Survival for each cycle in this simulation run: Sum columns of states that aren't death.
      m_surv[k,] <- rowSums(m_M_Micro[, 2:n_s])       # Assumes the last column is death state 
    
    # Display simulation progress
    if(k %in% seq(0,n_sim,5)) { # display progress every 10%
      cat('\r', paste(round(k/n_sim*100,0), "% done", sep = " "))
    } else if (k == n_sim) {cat('\r', paste("100% done"))}
    
  } # close loop for simulations 
  
  # Calculate survival and 95% CI from all simulations
  df_surv <- data.frame(surv  = colMeans(m_surv),
                        lower = colQuantiles(m_surv, probs = seq(from = 0, to = 1, by = 0.025))[, "2.5%"],
                        upper = colQuantiles(m_surv, probs = seq(from = 0, to = 1, by = 0.025))[, "97.5%"])
  
  # Calculate outputs for extrapolated period
  df_ext <- outputs - df_trial 
  
  # store the results from the simulation in a list
  results <- list(msm_trace = m_M_Micro, outputs = outputs, surv = df_surv, extrapolation = df_ext, trial = df_trial)   
  
  return(results)  # return the results
  
} # end of the MicroSim function 
