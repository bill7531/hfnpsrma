library(dmetar)
library(meta)
library(esc)
library(tidyverse)
library(openxlsx)
library(metafor)
library(dbplyr)
library(RTSA)

df.cont <- read.xlsx("cont-data.xlsx")
df.bin <- read.xlsx("bin-data.xlsx")
df.dem <- read.xlsx("dem-data.xlsx")

df.cont <- left_join(df.cont, df.dem, by = "author")
df.bin <- left_join(df.bin, df.dem, by = "author")

df.reintubation <- df.bin %>% filter(outcome == "reintubation")%>% filter(control1 != "NIV")
df.pneumonia <- df.bin %>% filter(outcome == "pneumonia")%>% filter(control1 != "NIV")
df.atelectasis <- df.bin %>% filter(outcome == "atelectasis")%>% filter(control1 != "NIV")
df.hypoxemia <- df.bin %>% filter(outcome == "hypoxemia")%>% filter(control1 != "NIV")
df.mortality <- df.bin %>% filter(outcome == "mortality")%>% filter(control1 != "NIV")
df.toniv <- df.bin %>% filter(outcome == "toniv")%>% filter(control1 != "NIV")
df.ppc <- df.bin %>% filter(outcome == "ppc")%>% filter(control1 != "NIV")

df.hlos <- df.cont %>% filter(outcome =="hlos")%>% filter(control1 != "NIV")
df.iculos <- df.cont %>% filter(outcome =="iculos")%>% filter(control1 != "NIV")
df.discomfort <- df.cont %>% filter(outcome =="discomfort")%>% filter(control1 != "NIV")
df.pfr2h <- df.cont %>% filter(outcome =="pfr2h")%>% filter(control1 != "NIV")
df.paco2early <- df.cont %>% filter(outcome =="paco2early")%>% filter(control1 != "NIV")
df.pao2early <- df.cont %>% filter(outcome =="pao2early")%>% filter(control1 != "NIV")
df.phearly <- df.cont %>% filter(outcome =="phearly")%>% filter(control1 != "NIV")
df.rrearly <- df.cont %>% filter(outcome =="rrearly")%>% filter(control1 != "NIV")
df.spoearly <- df.cont %>% filter(outcome =="spoearly")%>% filter(control1 != "NIV")
df.hrearly <- df.cont %>% filter(outcome =="hrearly")%>% filter(control1 != "NIV")
df.map <- df.cont %>% filter(outcome =="map")%>% filter(control1 != "NIV")

df.reintubation <- df.reintubation %>% arrange(author)

df.composite <- inner_join(df.reintubation, df.toniv, 
                           by = c("author", "year"), 
                           suffix = c(".reintub", ".escal"))
df.composite <- df.composite %>%
  mutate(
    n.c1.outcome = n.c1.outcome.reintub + n.c1.outcome.escal,
    n.e.outcome = n.e.outcome.reintub + n.e.outcome.escal,
    n.c1 = n.c1.reintub,
    n.e = n.e.reintub
  ) %>%
  select(author, year, surgery = surgery.reintub, n.c1.outcome, n.c1, n.e.outcome, n.e)

#forest 

m.reintu <- metabin(event.e = n.e.outcome, 
                 n.e = n.e,
                 event.c = n.c1.outcome,
                 n.c = n.c1,
                 studlab = author,
                 data = df.reintubation,
                 sm = "OR",
                 method = "Peto",
                 MH.exact = TRUE,
                 common = TRUE,
                 random = TRUE,
                 method.tau = "PM",
                 method.random.ci = "HK",
                 title = "Odds ratio for re-intubation")
summary(m.reintu)

m.pneumo <- metabin(event.e = n.e.outcome, 
                    n.e = n.e,
                    event.c = n.c1.outcome,
                    n.c = n.c1,
                    studlab = author,
                    data = df.pneumonia,
                    sm = "OR",
                    method = "Peto",
                    MH.exact = TRUE,
                    common = TRUE,
                    random = TRUE,
                    method.tau = "PM",
                    method.random.ci = "HK",
                    title = "Odds ratio for pneumonia")
summary(m.pneumo)

m.ate <- metabin(event.e = n.e.outcome, 
                    n.e = n.e,
                    event.c = n.c1.outcome,
                    n.c = n.c1,
                    studlab = author,
                    data = df.atelectasis,
                    sm = "OR",
                    method = "Peto",
                    MH.exact = TRUE,
                    common = TRUE,
                    random = TRUE,
                    method.tau = "PM",
                    method.random.ci = "HK",
                    title = "Odds ratio for atelectasis")
summary(m.ate)

m.hypox <- metabin(event.e = n.e.outcome, 
                   n.e = n.e,
                   event.c = n.c1.outcome,
                   n.c = n.c1,
                   studlab = author,
                   data = df.hypoxemia,
                   sm = "OR",
                   method = "Peto",
                   MH.exact = TRUE,
                   common = TRUE,
                   random = TRUE,
                   method.tau = "PM",
                   method.random.ci = "HK",
                   title = "Odds ratio for hypoxemia")
summary(m.hypox)

m.mortal <- metabin(event.e = n.e.outcome, 
                   n.e = n.e,
                   event.c = n.c1.outcome,
                   n.c = n.c1,
                   studlab = author,
                   data = df.mortality,
                   sm = "OR",
                   method = "Peto",
                   MH.exact = TRUE,
                   common = TRUE,
                   random = TRUE,
                   method.tau = "PM",
                   method.random.ci = "HK",
                   title = "Odds ratio for mortality")
summary(m.mortal)

m.toniv <- metabin(event.e = n.e.outcome, 
                    n.e = n.e,
                    event.c = n.c1.outcome,
                    n.c = n.c1,
                    studlab = author,
                    data = df.toniv,
                    sm = "OR",
                    method = "Peto",
                    MH.exact = TRUE,
                    common = TRUE,
                    random = TRUE,
                    method.tau = "PM",
                    method.random.ci = "HK",
                    title = "Odds ratio for escalate to NIV")
summary(m.toniv)
m.ppc <- metabin(event.e = n.e.outcome, 
                   n.e = n.e,
                   event.c = n.c1.outcome,
                   n.c = n.c1,
                   studlab = author,
                   data = df.ppc,
                   sm = "OR",
                   method = "Peto",
                   MH.exact = TRUE,
                   common = TRUE,
                   random = TRUE,
                   method.tau = "PM",
                   method.random.ci = "HK",
                   title = "Odds ratio for PPC")
summary(m.ppc)

m.hlos <- metacont(n.e = n.e,
                          mean.e = e.outcome,
                          sd.e = e.sd,
                          n.c = n.c1,
                          mean.c = c1.outcome,
                          sd.c = c1.sd,
                          studlab = author,
                          data = df.hlos,
                          sm = "MD",
                          common = TRUE,
                          random = TRUE,
                          method.tau = "REML",
                          method.random.ci = "HK",
                          title = "Hospital length of stay")
summary(m.hlos)
m.iculos <- metacont(n.e = n.e,
                   mean.e = e.outcome,
                   sd.e = e.sd,
                   n.c = n.c1,
                   mean.c = c1.outcome,
                   sd.c = c1.sd,
                   studlab = author,
                   data = df.iculos,
                   sm = "MD",
                   common = TRUE,
                   random = TRUE,
                   method.tau = "REML",
                   method.random.ci = "HK",
                   title = "ICU length of stay")
summary(m.iculos)

m.discom <- metacont(n.e = n.e,
                     mean.e = e.outcome,
                     sd.e = e.sd,
                     n.c = n.c1,
                     mean.c = c1.outcome,
                     sd.c = c1.sd,
                     studlab = author,
                     data = df.discomfort,
                     sm = "MD",
                     common = TRUE,
                     random = TRUE,
                     method.tau = "REML",
                     method.random.ci = "HK",
                     title = "Discomfort")
summary(m.discom)

m.pfr <- metacont(n.e = n.e,
                     mean.e = e.outcome,
                     sd.e = e.sd,
                     n.c = n.c1,
                     mean.c = c1.outcome,
                     sd.c = c1.sd,
                     studlab = author,
                     data = df.pfr2h,
                     sm = "MD",
                     common = TRUE,
                     random = TRUE,
                     method.tau = "REML",
                     method.random.ci = "HK",
                     title = "Early PF Ratio")
summary(m.pfr)

m.pao2 <- metacont(n.e = n.e,
                  mean.e = e.outcome,
                  sd.e = e.sd,
                  n.c = n.c1,
                  mean.c = c1.outcome,
                  sd.c = c1.sd,
                  studlab = author,
                  data = df.pao2early,
                  sm = "MD",
                  common = TRUE,
                  random = TRUE,
                  method.tau = "REML",
                  method.random.ci = "HK",
                  title = "Early PaO2")
summary(m.pao2)
m.paco2 <- metacont(n.e = n.e,
                   mean.e = e.outcome,
                   sd.e = e.sd,
                   n.c = n.c1,
                   mean.c = c1.outcome,
                   sd.c = c1.sd,
                   studlab = author,
                   data = df.paco2early,
                   sm = "MD",
                   common = TRUE,
                   random = TRUE,
                   method.tau = "REML",
                   method.random.ci = "HK",
                   title = "Early PaCO2")
summary(m.paco2)
m.ph <- metacont(n.e = n.e,
                    mean.e = e.outcome,
                    sd.e = e.sd,
                    n.c = n.c1,
                    mean.c = c1.outcome,
                    sd.c = c1.sd,
                    studlab = author,
                    data = df.phearly,
                    sm = "MD",
                    common = TRUE,
                    random = TRUE,
                    method.tau = "REML",
                    method.random.ci = "HK",
                    title = "Early pH")
summary(m.ph)
summary(m.paco2)
m.rr <- metacont(n.e = n.e,
                 mean.e = e.outcome,
                 sd.e = e.sd,
                 n.c = n.c1,
                 mean.c = c1.outcome,
                 sd.c = c1.sd,
                 studlab = author,
                 data = df.rrearly,
                 sm = "MD",
                 common = TRUE,
                 random = TRUE,
                 method.tau = "REML",
                 method.random.ci = "HK",
                 title = "Early RR")
summary(m.rr)
m.sp <- metacont(n.e = n.e,
                 mean.e = e.outcome,
                 sd.e = e.sd,
                 n.c = n.c1,
                 mean.c = c1.outcome,
                 sd.c = c1.sd,
                 studlab = author,
                 data = df.spoearly,
                 sm = "MD",
                 common = TRUE,
                 random = TRUE,
                 method.tau = "REML",
                 method.random.ci = "HK",
                 title = "Early SaO2")
summary(m.sp)
m.hr <- metacont(n.e = n.e,
                 mean.e = e.outcome,
                 sd.e = e.sd,
                 n.c = n.c1,
                 mean.c = c1.outcome,
                 sd.c = c1.sd,
                 studlab = author,
                 data = df.hrearly,
                 sm = "MD",
                 common = TRUE,
                 random = TRUE,
                 method.tau = "REML",
                 method.random.ci = "HK",
                 title = "Early Heart Rate")
summary(m.hr)
m.map <- metacont(n.e = n.e,
                 mean.e = e.outcome,
                 sd.e = e.sd,
                 n.c = n.c1,
                 mean.c = c1.outcome,
                 sd.c = c1.sd,
                 studlab = author,
                 data = df.map,
                 sm = "MD",
                 common = TRUE,
                 random = TRUE,
                 method.tau = "REML",
                 method.random.ci = "HK",
                 title = "Mean arterial pressure")
summary(m.map)

#sensitivity analysis
reintu.inf <- metainf(m.reintu, pooled = "random")
plot(reintu.inf)
baujat(m.reintu,
       xlim = c(0,6),
       symbol = "slab",
       cex = 0.8,
       col = "blue",
       pos = 4)
m.reintu_sens <- update(m.reintu, 
                         subset = !(author %in% c("Yu", "Sahin")))



# 1. Create the dataframe with the exact required names
tsa_data <- data.frame(
  eI = c(2, 7, 1, 0, 0, 1, 0, 3, 7, 5, 2), # Intervention Events
  nI = c(169, 108, 47, 30, 40, 90, 32, 47, 66, 51, 140), # Intervention Total
  eC = c(0, 4, 1, 0, 2, 2, 0, 1, 2, 1, 1), # Control Events
  nC = c(171, 112, 48, 30, 40, 90, 32, 43, 33, 49, 148) # Control Total
)

# 2. Run the RTSA
tsa_reintu <- RTSA(
  type    = "analysis",
  data    = tsa_data,
  outcome    = "OR",           # Outcome metric: Relative Risk
  mc         = 0.8,            # Minimal clinical relevant outcome (20% RRR)
  side       = 2,              # Two-sided hypothesis test
  alpha      = 0.05,           # 5% Type I error
  beta       = 0.1,            # 10% Type II error (90% power)
  fixed      = FALSE,          # Disable fixed-effect (enables random-effects)
  re_method  = "DL_HKSJ",      # Apply Hartung-Knapp-Sidik-Jonkman adjustment
  random_adj = "D2",           # Adjust sample size based on Diversity (D2)
  es_alpha   = "esOF" 
)

# 3. View the Result
plot(tsa_reintu)

#meta::funnel(m.reintu, studlab = TRUE)

bias_result <- metabias(m.reintu, method.bias = "harbord")
print(bias_result)
funnel(m.reintu, 
       main = "Funnel Plot: Re-intubation (k=11)",
       xlab = "Odds Ratio (Peto Method)",
       studlab = TRUE,    # Adds study names like Yu, Sahin, Futier
       cex.stud = 0.8)

m.reintu_highrisk <- update(m.reintu, 
                             subgroup = highrisk, 
                             print.subgroup.name = FALSE)
m.reintu_rob <- update(m.reintu, 
                            subgroup = rob, 
                            print.subgroup.name = FALSE)

m.reintu_surg <- update(m.reintu, 
                       subgroup = surgery, 
                       print.subgroup.name = FALSE)
meta::forest(m.reintu_highrisk,
             common = FALSE,                # Plot only random effects
             random = TRUE,
             smlab = "Reintubation Rates by Risk",
             layout = "RevMan5",  # Plot random effects estimate
             col.random = "black",          # Set random effects diamond color
             digits = 2,                    # Significant digits for treatment effects
             digits.pval = 2,               # Significant digits for p-values
             digits.I2 = 1,                 # Significant digits for I-squared
             digits.tau2 = 2,               # Significant digits for between-study variance
             test.subgroup = TRUE,          # Print test for subgroup differences
             subgroup.name = "high risk",    # Label for the grouping variable
             sep.subgroup = "=",            # Separator between label and levels
             print.subgroup.name = TRUE,     # Print variable name in front of labels
             addrows.below.overall = 2,
             )
