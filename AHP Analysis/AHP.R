rm(list = ls())
cat("\014")
graphics.off()

# Import library
library(ahpsurvey)
library(table1)
library(dplyr)
library(knitr)
library(kableExtra)
library(ggplot2)

library(tidyr)


# Step 1 Import Data
ImportedData <- read_excel("AHP Analysis/SurveyData07.xlsx")
SurveyData <- ImportedData[, 1:6]

# Step 2:- Define attributes
atts <- c("Occupancy", "ServiceAvailability", "TravelTimeReliability", "EnvironmentFactor")


# Step 03:- Creating pairwise comparison matrices
surveyahp <- SurveyData %>%
  ahp.mat(atts, negconvert = T) 

eigentrue <- ahp.indpref(surveyahp, atts, method = "eigen")


# Step 04: Measuring Consistency Ratio
cr <- SurveyData %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

table(cr <= 0.1)


# Step 5: Visualising individual priorities and consistency ratios
thres <- 0.1

dict <- c("Occupancy" = "Occupancy", 
          "ServiceAvailability" = "ServiceAvailability", 
          "TravelTimeReliability" = "TravelTimeReliability", 
          "EnvironmentFactor" = "EnvironmentFactor")

cr.df <- SurveyData %>%
  ahp.mat(atts, negconvert = TRUE) %>% 
  ahp.cr(atts) %>% 
  data.frame() %>%
  mutate(rowid = 1:length(cr), cr.dum = as.factor(ifelse(cr <= thres, 1, 0))) %>%
  dplyr::select(cr.dum, rowid)


SurveyData %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.indpref(atts, method = "eigen") %>% 
  mutate(rowid = 1:nrow(eigentrue)) %>%
  left_join(cr.df, by = 'rowid') %>%
  gather(Occupancy, ServiceAvailability, TravelTimeReliability, EnvironmentFactor, key = "var", value = "pref") %>%
  ggplot(aes(x = var, y = pref)) + 
  geom_violin(alpha = 0.6, width = 0.8, color = "transparent", fill = "gray") +
  geom_jitter(alpha = 0.6, height = 0, width = 0.1, aes(color = cr.dum)) +
  geom_boxplot(alpha = 0, width = 0.3, color = "#808080") +
  scale_x_discrete(" ", label = dict) +
  scale_y_continuous("Weight", 
                     labels = scales::percent, 
                     breaks = c(seq(0,0.7,0.1))) +
  guides(color=guide_legend(title=NULL))+
  scale_color_discrete(breaks = c(0,1), 
                       labels = c(paste("CR >", thres), 
                                  paste("CR <", thres))) +
  #labs(NULL, caption = paste("n =", nrow(SurveyData), ",", "Mean CR =",round(mean(cr),3)))+
  theme_minimal()


# Step 06: Collect only consistent matrix 
ImportedDataFinal <- ImportedData %>% 
                        mutate(rowid = 1:nrow(eigentrue)) %>%
                        left_join(cr.df, by = 'rowid')


ConsistentData <- ImportedDataFinal[ImportedDataFinal$'cr.dum' == '1', ]
InconsistentData <- ImportedDataFinal[ImportedDataFinal$'cr.dum' == '0', ]


# Step 09: Aggregated Weight Estimation for attributes
ConsistentDataAHP <- ConsistentData[,1:6] %>% ahp.mat(atts, negconvert = T)

ConAggAHPMatrix <- ConsistentData[,1:6] %>%
                      ahp.mat(atts = atts, negconvert = TRUE) %>% 
                      ahp.aggjudge(atts, aggmethod = "geometric")


amean <- ahp.aggpref(ConsistentDataAHP, atts, method = "arithmetic")
amean

mean <- ConsistentData[,1:6] %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.aggpref(atts, method = "arithmetic")
print(mean)

sd <- ConsistentData[,1:6] %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.aggpref(atts, method = "arithmetic", aggmethod = "sd")
print(sd)


#================================================--#================================================--
#================================================--#================================================--

# Step 10: Finding inconsistent pairwise comparisons by maximum
atts <- c("OC", "SA", "TTR", "EF")

InconsistentDataAHP <- InconsistentData[,1:6] %>% ahp.mat(atts, negconvert = T)

pwInconsistentAttributes <- InconsistentDataAHP %>%
  ahp.pwerror(atts)

print(pwInconsistentAttributes)

InconsistentDataAHP %>%
  ahp.pwerror(atts) %>% 
  gather(top1, top2, top3, key = "max", value = "pair") %>%
  table() %>%
  as.data.frame() %>%
  ggplot(aes(x = pair, y = Freq, fill = max)) + 
  geom_bar(stat = 'identity') +
  scale_y_continuous("Frequency", breaks = c(seq(0,180,20))) +
  scale_fill_discrete(breaks = c("top1", "top2", "top3"), labels = c("1", "2", "3")) +
  scale_x_discrete("Pair") +
  guides(fill = guide_legend(title="Rank")) +
  theme(axis.text.x = element_text(angle = 20, hjust = 1),
        panel.background = element_rect(fill = NA),
        panel.grid.major.y = element_line(colour = "grey80"),
        panel.grid.major.x = element_blank(),
        panel.ontop = FALSE)


# Draw plot in python. It will look good--> take data from here and use matplotlib.pyplot


# Step 11: Example of Transforming --> Inconsistent Matrix to --> Consistent Matrix
print(InconsistentDataAHP[[1]])
family <- c(1, 1/6, 1/7, 5, 6, 1, 1/6, 5, 7, 6, 1,  7, 1/5, 1/5, 1/7,  1)
fam.mat <- list(matrix(family, nrow = 4 , ncol = 4))
atts <- c("OC", "SA", "TTR", "EF")
rownames(fam.mat[[1]]) <- colnames(fam.mat[[1]]) <- atts
fam.mat[[1]] %>% kable()
ahp.cr(fam.mat, atts)

edited <- ahp.harker(fam.mat, atts, iterations = 2, stopcr = 0.1, limit = T)
edited[[1]]

ahp.cr(edited, atts)


# Step 12: Transform the inconsistent pairwise comparison matrix into consistent pairwise comparison matrix
crmat <- matrix(NA, nrow = 63, ncol = 6)
colnames(crmat) <- 0:5

atts <- c("OC", "SA", "TTR", "EF")

crmat[,1] <- SurveyData %>%
  ahp.mat(atts, negconvert = TRUE) %>%
  ahp.cr(atts)

for (it in 1:5){
  crmat[,it+1] <- SurveyData %>%
    ahp.mat(atts, negconvert = TRUE) %>%
    ahp.harker(atts, iterations = it, stopcr = 0.1, 
               limit = T, round = T, printiter = T) %>%
    ahp.cr(atts)
}

data.frame(table(crmat[,1] <= 0.1),
           table(crmat[,2] <= 0.1),
           table(crmat[,3] <= 0.1),
           table(crmat[,4] <= 0.1),
           table(crmat[,5] <= 0.1)) %>% 
  select(Var1, Freq, Freq.1, Freq.2, Freq.3, Freq.4) %>%
  rename("Consistent?" = "Var1", "No Iteration" = "Freq",
         "1 Iterations" = "Freq.1", "2 Iterations" = "Freq.2",
         "3 Iterations" = "Freq.3", "4 Iterations" = "Freq.4")


crmat %>% 
  as.data.frame() %>%
  gather(key = "iter", value = "cr", `0`, 1,2,3,4,5,6) %>%
  mutate(iter = as.integer(iter)) %>%
  ggplot(aes(x = iter, y = cr, group = iter)) +
  geom_hline(yintercept = 0.1, color = "red", linetype = "dashed")+
  geom_jitter(alpha = 0.2, width = 0.3, height = 0, color = "turquoise4") +
  geom_boxplot(fill = "transparent", color = "#808080", outlier.shape = NA) + 
  scale_x_continuous("Iterations", breaks = 0:10) +
  scale_y_continuous("Consistency Ratio") +
  theme_minimal()


# Step 13: Individual preference weights with respect to goal (1 iteration)
it <- 1
thres <- 0.1
cr.df1 <- data.frame(cr = SurveyData %>%
                       ahp.mat(atts, negconvert = TRUE) %>%
                       ahp.harker(atts, iterations = it, stopcr = 0.1, limit = T, round = T, printiter = F) %>%
                       ahp.cr(atts))

cr.df2 <- cr.df1 %>%
  mutate(rowid = 1:nrow(SurveyData), cr.dum = as.factor(ifelse(. <= thres, 1, 0))) %>%
  select(cr.dum, rowid)


SurveyData %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.harker(atts, iterations = it, stopcr = 0.1, limit = T, round = T, printiter = F) %>%
  ahp.indpref(atts, method = "eigen") %>% 
  mutate(rowid = 1:nrow(SurveyData)) %>%
  left_join(cr.df2, by = 'rowid') %>%
  gather(OC, SA, TTR, EF, key = "var", value = "pref") %>%
  ggplot(aes(x = var, y = pref)) + 
  geom_violin(alpha = 0.6, width = 0.8, color = "transparent", fill = "gray") +
  geom_jitter(alpha = 0.3, height = 0, width = 0.1, aes(color = cr.dum)) +
  geom_boxplot(alpha = 0, width = 0.3, color = "#808080") +
  scale_x_discrete("Attribute", label = dict) +
  scale_y_continuous("Weight (dominant eigenvalue)", 
                     labels = scales::percent, breaks = c(seq(0,0.7,0.1))) +
  guides(color=guide_legend(title=NULL))+
  scale_color_discrete(breaks = c(0,1), 
                       labels = c(paste("CR >", thres), 
                                  paste("CR <", thres))) +
  labs(NULL, caption =paste("n =",nrow(SurveyData), ",", "Mean CR =",round(mean(cr),3)))+
  theme_minimal()



#================================================--#================================================--
#================================================--#================================================--


options(scipen = 99)
inconsistent <- SurveyData %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.aggpref(atts, method = "geometric")

consistent <- SurveyData %>%
  ahp.mat(atts = atts, negconvert = TRUE) %>% 
  ahp.harker(atts, iterations = 5, stopcr = 0.1, limit = T, round = T, printiter = F) %>%
  ahp.aggpref(atts, method = "geometric")


weight <- c(5,-3,2,-5,
            -7,-1,-7,
            4,-3,
            -7)
sample_mat <- ahp.mat(t(weight), atts, negconvert = TRUE)


true <- t(ahp.indpref(sample_mat, atts, method = "eigen"))

aggpref.df <- data.frame(Attribute = atts, true,inconsistent,consistent) %>%
  mutate(error.incon = abs(true - inconsistent),
         error.con = abs(true - consistent))

aggpref.df


#================================================--#================================================--
#================================================--#================================================--

# Final consistent matrix after Harker inconsistent incorporation

cr.df3 <- crmat[,3]  #here second column save as it is the 2nd iteration
cr.df3 <- data.frame(rowid = seq_along(cr.df3), values = cr.df3)
names(cr.df3)[names(cr.df3) == "values"] <- "cr.dum"

FinalDataSet <- ImportedData %>%
                    mutate(rowid = 1:nrow(eigentrue)) %>%
                    left_join(cr.df3, by = 'rowid')

FinalConsistentData <- FinalDataSet[FinalDataSet$'cr.dum' <= '0.1', ]

FinalData <- FinalConsistentData[, 1:6]

FinalDataAHP <- FinalData %>% ahp.mat(atts, negconvert = T)

FinalDataAHPMatrix <- FinalData %>% 
                          ahp.mat(atts, negconvert = T) %>%
                          ahp.aggjudge(atts, aggmethod = "geometric")    

amean <- ahp.aggpref(FinalDataAHP, atts, aggmethod = "arithmetic")
amean


FinalDataAHPMatrix


#================================================--#================================================--
#================================================--#================================================--

# Weights and consistency for each categories for experts

ResearcherAcademician <- FinalConsistentData[FinalConsistentData$"Your role" == "Researcher or Academician", ]
IndustryProfessional  <- FinalConsistentData[FinalConsistentData$`Your role` == 'Industry Professional', ]
ServiceOperatorServiceProvider <- FinalConsistentData[FinalConsistentData$`Your role` == 'Service operator or service provider', ]
PolicyMakerAdviser <- FinalConsistentData[FinalConsistentData$`Your role` == 'Policy maker or Adviser', ]
Other <- FinalConsistentData[FinalConsistentData$`Your role` == 'Other', ]


ResearcherAcademicianMAT <- ResearcherAcademician[,1:6] %>% ahp.mat(atts, negconvert = T) %>% ahp.aggjudge(atts, aggmethod = "geometric")
IndustryProfessionalMAT  <- IndustryProfessional[, 1:6] %>% ahp.mat(atts, negconvert = T) %>% ahp.aggjudge(atts, aggmethod = "geometric")
ServiceOperatorServiceProviderMAT <- ServiceOperatorServiceProvider[,1:6] %>% ahp.mat(atts, negconvert = T) %>% ahp.aggjudge(atts, aggmethod = "geometric")
PolicyMakerAdviserMAT <- PolicyMakerAdviser[,1:6] %>% ahp.mat(atts, negconvert = T) %>% ahp.aggjudge(atts, aggmethod = "geometric")
OtherMAT <- Other[,1:6] %>% ahp.mat(atts, negconvert = T) %>% ahp.aggjudge(atts, aggmethod = "geometric")
  

ResearcherAcademicianAHP <- ResearcherAcademician[,1:6] %>% ahp.mat(atts, negconvert = T) 
IndustryProfessionalAHP  <- IndustryProfessional[, 1:6] %>% ahp.mat(atts, negconvert = T) 
ServiceOperatorServiceProviderAHP <- ServiceOperatorServiceProvider[,1:6] %>% ahp.mat(atts, negconvert = T) 
PolicyMakerAdviserAHP <- PolicyMakerAdviser[,1:6] %>% ahp.mat(atts, negconvert = T) 
OtherAHP <- Other[,1:6] %>% ahp.mat(atts, negconvert = T) 


# Researcher and academician
cr.academician <- ResearcherAcademician[,1:6] %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

ResearcherAcademicianMean <- ahp.aggpref(ResearcherAcademicianAHP, atts, method = "arithmetic")
ResearcherAcademicianMean


# Industry and Professional
cr.IndustryProfessional <- IndustryProfessional[, 1:6] %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

IndustryProfessionalMean <- ahp.aggpref(IndustryProfessionalAHP, atts, method = "arithmetic")
IndustryProfessionalMean


# ServiceOperat or ServiceProvider
cr.ServiceOperator <- ServiceOperatorServiceProvider[,1:6] %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

ServiceOperatorServiceProviderMean <- ahp.aggpref(ServiceOperatorServiceProviderAHP, atts, method = "arithmetic")
ServiceOperatorServiceProviderMean


# PolicyMaker or Adviser
cr.PolicyMaker <- PolicyMakerAdviser[,1:6] %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

PolicyMakerAdviserMean <- ahp.aggpref(PolicyMakerAdviserAHP, atts, method = "arithmetic")
PolicyMakerAdviserMean


# Other 
cr.Other <- Other[,1:6] %>%
  ahp.mat(atts, negconvert = T) %>% 
  ahp.cr(atts)

OtherMean <- ahp.aggpref(OtherAHP, atts, method = "arithmetic")
OtherMean


#calculate the indivisual CR and CI using prepared excel sheet

# Library
library(openxlsx)


# Define the file path where you want to save the Excel file
excel_file <- "AHP Analysis/AHP_DATA"

# Export the dataset to Excel
#write.xlsx(FinalConsistentData, excel_file, rowNames = FALSE)

# Provide a message once the export is complete
#cat("Data has been exported to:", excel_file, "\n")



