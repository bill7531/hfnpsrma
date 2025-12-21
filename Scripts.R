library(dmetar)
library(meta)
library(esc)
library(tidyverse)
library(openxlsx)
library(metafor)
library(dbplyr)

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
                   title = "Early PaO2")
summary(m.paco2)
