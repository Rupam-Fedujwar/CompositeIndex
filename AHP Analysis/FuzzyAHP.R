rm(list = ls())
cat("\014")
graphics.off()

library("readxl")
library("dplyr")
library("tidyverse")
library("FuzzyAHP")
library("table1")
library("psych")

#1. reading data from excel file ====

#data <- read_excel("AHP Analysis/Fuzzy AHP/Updated_FuzzyAHPData.xlsx", sheet = "Raw Dataset", range = "A1:M63", col_names = TRUE)
data <- read_excel("AHP Analysis/Fuzzy AHP/Updated_ConsistentData.xlsx", sheet = "Raw Dataset", range = "A1:M52", col_names = TRUE)
data <- tibble(data)
colname<- c("Name", "Gender", "Designation", "Affliation", "YourRole", "Role_Others", "TotalExperience",
            "OC_SA",	"OC_TTR",	"OC_EF", "SA_TTR", "SA_EF", "TTR_EF")

colnames(data) <- colname

#2. reorienting data as per criteria sequence to create pairwise comparison matrix ====

data <- data.frame(data[, 1:7], data[, 8:10], data[, 11:12], data[, 13])
data <- data[2:nrow(data), ]

data[, 8:13] <- lapply(data[, 8:13], as.numeric)

#Making Consistent Entries for Experience
unique <- sort(unique(data$TotalExperience))

data$TotalExperience[data$TotalExperience %in% unique[1]] <- "0-5"
data$TotalExperience[data$TotalExperience %in% unique[2]] <- "5-10"
data$TotalExperience[data$TotalExperience %in% unique[3]] <- "10-15"
data$TotalExperience[data$TotalExperience %in% unique[4]] <- "15-20"
data$TotalExperience[data$TotalExperience %in% unique[5]] <- "More than 20"


#3. Preparing Tables for Sample Characteristics ====

data$Gender <- factor(data$Gender, levels = c("Male", "Female", "Did not revealed"))
data$TotalExperience <- factor(data$TotalExperience, levels = c("0-5", "5-10", "10-15", "15-20", "More than 20"))
data$YourRole <- factor(data$YourRole, levels = c("Researcher or Academician", "Policy maker or Adviser", "Service operator or service provider", "Industry Professional", "Other"))

table1(~ Gender + TotalExperience | YourRole, data =  data)


#4. Defining expert_array function  ====

#Note:- "expert_array" function will take rating info based on expert's role and returns Paired Comparison Matrices array

expert_array <- function(rating_info) {
  
  result <- array(0, dim=c(4,4, nrow(rating_info)))
  
  
  #Assigning expert's rating to the attributes
  
  for(i in 1:nrow(rating_info)) {
    
    #Creating singular matrix for all the attributes
    OC <- matrix(1, 1, 4, byrow = TRUE)
    SA <- matrix(1, 1, 4, byrow = TRUE)
    TTR <- matrix(1, 1, 4, byrow = TRUE)
    EF <- matrix(1, 1, 4, byrow = TRUE)

    expert <- matrix(NA, 4, 4, byrow = TRUE)
    expert2 <- matrix(NA, 4, 4, byrow = TRUE)
    finsam <- matrix(NA, 4, 4, byrow = TRUE)
    
    OC[1, 2:4] <- matrix(rating_info[i, 8:10], 1, 3); unlist(OC)
    SA[1, 3:4] <- matrix(rating_info[i, 11:12], 1, 2); unlist(SA)
    TTR[1, 4:4] <- matrix(rating_info[i, 13]); unlist(TTR)
    
    
    #Creating Paired Comparison Matrix from Expert's Sample    
    resp <- rbind(OC, SA, TTR, EF)
    resp[lower.tri(resp)] <- 0
    
    for(j in 1:4) {
      for(k in 1:4){
        expert [k,j] <- as.numeric(resp[k,j])
        expert2[k,j] <- 1/as.numeric(resp[j,k])
      }
    }
    
    expert2[upper.tri(expert2)] <- 0
    finsam <- expert + expert2 - diag(1, 4, 4)
    round(finsam, 2)
    tibble(finsam)
    
    result[, ,i] <- finsam
    
  }
  
  rownames(result) <- c("OC", "SA", "TTR", "EF")
  colnames(result) <- c("OC", "SA", "TTR", "EF")
  
  return(result)
}


#5 Creating PCM from expert_array function ====

#For Individual PCM from Total Dataset
total_PCM <- expert_array(data)

#6. Obtaining All Experts' weight ====

#6.1 Creation of Fuzzy-PCM ====
#Note: This step return PCM object as list format having one element as individual Fuzzy PCM

Mat <- data %>%
  expert_array()

a <- dim(Mat)

PCM <- list()

for(i in 1:a[3]) {
  
  comparisonMatrixValues <- Mat[1:4, 1:4, i]
  comparisonMatrix <- matrix(comparisonMatrixValues, nrow = 4, ncol = 4, byrow = FALSE)
  comparisonMatrix <- pairwiseComparisonMatrix(comparisonMatrix)
  comparisonMatrix <- fuzzyPairwiseComparisonMatrix(comparisonMatrix)
  PCM[[i]] <- comparisonMatrix
}

#6.2 Aggregation of Stakeholders Fuzzy PCM and Criteria Weights derivation ====

#Note: This step aggregate all individual Fuzzy-PCM from PCM List object that is created in previous step 6.1

Fmin <- NA
Fmodal <- NA
Fmax <- NA

Emin <- matrix(NA, 4,4)
Emodal <- matrix(NA, 4,4)
Emax <- matrix(NA, 4,4)

Comb <- matrix(NA, 4,4)

GMmin <- matrix(1, 4,1)
GMmodal <- matrix(1, 4,1)
GMmax <- matrix(1, 4,1)

for (j in 1:4) {
  
  for (k in 1:4) {
    
    for(i in 1:a[3]) {
      
      #Each fuzzy cell from individual Fuzzy-PCM is taken and stored in F(Fmin, Fmodal, Fmax)
      Fmin [i] <- PCM[[i]]@fnMin[j,k]
      Fmodal[i] <- PCM[[i]]@fnModal[j,k]
      Fmax[i] <- PCM[[i]]@fnMax[j,k]
    }
    
    #Geometric Mean function is used to aggregate all fuzzy rating from fuzzy cell i.e. F(Fmin, Fmodal, Fmax) to get aggregated PCM i.e. E(Emin, Emodal, Emax)
    Emin[j, k] <- round(geometric.mean(Fmin), digits = 2)
    Emodal[j, k] <- round(geometric.mean(Fmodal), digits = 2)
    Emax[j, k] <- round(geometric.mean(Fmax), digits = 2)
    
    #This combines the E(Emin, Emodal, Emax) as string format to get a single aggregated PCM from all experts.
    Comb[j,k] <- paste0("(", Emin[j,k], ",", Emodal[j,k], ",", Emax[j,k], ")")
    
    #This step perform cell-wise product of all columns of E(Emin, Emodal, Emax) to get a single column G(Gmin, Gmodal, Gmax) which will be further used to get geometric mean of the column.
    GMmin[j] <- GMmin[j]*round(Emin[j,k], digits = 2)
    GMmodal[j] <- GMmodal[j]*round(Emodal[j,k], digits = 2)
    GMmax[j] <- GMmax[j]*round(Emax[j,k], digits = 2)
  }
}

#The G(Gmin, Gmodal, Gmax)^(1/4) will give geometric mean of columns from E(Emin, Emodal, Emax)
for (j in 1:4) {
  
  GMmin[j] <- GMmin[j]^(1/4)
  GMmodal[j] <- GMmodal[j]^(1/4)
  GMmax[j] <- GMmax[j]^(1/4)
}

Comb <- as.data.frame(Comb) #Converts aggregated PCM to a dataframe 
Comb

#Initiation of Criteria Weights derivation
Wt_Min <- NA
Wt_Modal <- NA
Wt_Max <- matrix(NA, 4,1)

for (j in 1:4) {
  Wt_Min[j] <- GMmin[j]/(sum(GMmax))
  Wt_Modal[j] <- GMmodal[j]/(sum(GMmodal))
  Wt_Max[j] <- GMmax[j]/(sum(GMmin))
}

Wt <- NA

for (j in 1:4) {
  Wt[j] <- (Wt_Min[j]+Wt_Modal[j]+Wt_Max[j])/3 #Centroid Area MEthod used to defuzzify the weights
}


for (j in 1:4) {
  # Calculate Wt[j]
  Wt[j] <- (Wt_Min[j] + Wt_Modal[j] + Wt_Max[j]) / 3
  
  # Print the values of Wt_Min, Wt_Modal, Wt_Max, and Wt for each iteration
  cat("Iteration:", j, "\n")
  cat("Wt_Min:", Wt_Min[j], "\n")
  cat("Wt_Modal:", Wt_Modal[j], "\n")
  cat("Wt_Max:", Wt_Max[j], "\n")
  cat("Wt:", Wt[j], "\n")
  
  cat("\n")  # Empty line for better readability
}


#Normalisation of Crisp weights 
Comb_Wt <- NA

for (j in 1:4) {
  Comb_Wt[j] <- Wt[j]/sum(Wt)
}


#7. Similarly obtaining Wt_Acad, Wt_Indu, Wt_Poli, Wt_Serv, Wt_Other These are reads as E.g. Weights of Industry Prof. ====
for( role in unique(data$YourRole)) {
  
  Mat <- data %>%
    filter(YourRole == role) %>% 
    expert_array()
  
  a <- dim(Mat)
  
  PCM <- list()
  
  for(i in 1:a[3]) {
    
    comparisonMatrixValues <- Mat[1:4, 1:4, i]
    comparisonMatrix <- matrix(comparisonMatrixValues, nrow = 4, ncol = 4, byrow = FALSE)
    comparisonMatrix <- pairwiseComparisonMatrix(comparisonMatrix)
    comparisonMatrix <- fuzzyPairwiseComparisonMatrix(comparisonMatrix)
    PCM[[i]] <- comparisonMatrix
  }
  
  Fmin <- NA
  Fmodal <- NA
  Fmax <- NA
  
  Emin <- matrix(NA, 4,4)
  Emodal <- matrix(NA, 4,4)
  Emax <- matrix(NA, 4,4)
  
  E <- matrix(NA, 4,4)
  
  GMmin <- matrix(1, 4,1)
  GMmodal <- matrix(1, 4,1)
  GMmax <- matrix(1, 4,1)
  
  for (j in 1:4) {
    
    for (k in 1:4) {
      
      for(i in 1:a[3]) {
        
        Fmin [i] <- PCM[[i]]@fnMin[j,k]
        Fmodal[i] <- PCM[[i]]@fnModal[j,k]
        Fmax[i] <- PCM[[i]]@fnMax[j,k]
      }
      
      Emin[j, k] <- round(geometric.mean(Fmin), digits = 2)
      Emodal[j, k] <- round(geometric.mean(Fmodal), digits = 2)
      Emax[j, k] <- round(geometric.mean(Fmax), digits = 2)
      
      E[j,k] <- paste0("(", Emin[j,k], ",", Emodal[j,k], ",", Emax[j,k], ")")
      
      GMmin[j] <- GMmin[j]*round(Emin[j,k], digits = 2)
      GMmodal[j] <- GMmodal[j]*round(Emodal[j,k], digits = 2)
      GMmax[j] <- GMmax[j]*round(Emax[j,k], digits = 2)
    }
  }
  
  for (j in 1:4) {
    
    GMmin[j] <- GMmin[j]^(1/4)
    GMmodal[j] <- GMmodal[j]^(1/4)
    GMmax[j] <- GMmax[j]^(1/4)
  }
  
  E <- as.data.frame(E)
  
  name <- paste0("PCM_",substr(role,1, 4), sep="")
  assign(name, E)
  
  Wt_Min <- NA
  Wt_Modal <- NA
  Wt_Max <- matrix(NA, 4,1)
  
  for (j in 1:4) {
    Wt_Min[j] <- GMmin[j]/(sum(GMmax))
    Wt_Modal[j] <- GMmodal[j]/(sum(GMmodal))
    Wt_Max[j] <- GMmax[j]/(sum(GMmin))
  }
  
  Wt <- NA
  
  for (j in 1:4) {
    Wt[j] <- (Wt_Min[j]+Wt_Modal[j]+Wt_Max[j])/3
    
  }
  
  Wt2 <- NA
  
  for (j in 1:4) {
    
    Wt2[j] <- Wt[j]/sum(Wt)    
  }
  
  name <- paste0("Wt_",substr(role,1, 4), sep="")
  assign(name, Wt2)
  
} 



#8. Similarly, deriving weights For NonAcad expert matrix ====
Mat <- data %>%
  filter(YourRole != "Researcher or Academician") %>% 
  expert_array()

a <- dim(Mat)
PCM <- list()

for(i in 1:a[3]) {
  
  comparisonMatrixValues <- Mat[1:4, 1:4, i]
  comparisonMatrix <- matrix(comparisonMatrixValues, nrow = 4, ncol = 4, byrow = FALSE)
  comparisonMatrix <- pairwiseComparisonMatrix(comparisonMatrix)
  comparisonMatrix <- fuzzyPairwiseComparisonMatrix(comparisonMatrix, getFuzzyScale("full"))
  PCM[[i]] <- comparisonMatrix
  
}


Fmin <- NA
Fmodal <- NA
Fmax <- NA

Emin <- matrix(NA, 4,4)
Emodal <- matrix(NA, 4,4)
Emax <- matrix(NA, 4,4)

PCM_NonAcad <- matrix(NA, 4,4)

GMmin <- matrix(1, 4,1)
GMmodal <- matrix(1, 4,1)
GMmax <- matrix(1, 4,1)

for (j in 1:4) {
  
  for (k in 1:4) {
    
    for(i in 1:a[3]) {
      
      Fmin [i] <- PCM[[i]]@fnMin[j,k]
      Fmodal[i] <- PCM[[i]]@fnModal[j,k]
      Fmax[i] <- PCM[[i]]@fnMax[j,k]
    }
    
    Emin[j, k] <- round(geometric.mean(Fmin), digits = 2)
    Emodal[j, k] <- round(geometric.mean(Fmodal), digits = 2)
    Emax[j, k] <- round(geometric.mean(Fmax), digits = 2)
    
    PCM_NonAcad[j,k] <- paste0("(", Emin[j,k], ",", Emodal[j,k], ",", Emax[j,k], ")")
    
    GMmin[j] <- GMmin[j]*round(Emin[j,k], digits = 2)
    GMmodal[j] <- GMmodal[j]*round(Emodal[j,k], digits = 2)
    GMmax[j] <- GMmax[j]*round(Emax[j,k], digits = 2)
  }
}

PCM_NonAcad <- as.data.frame(PCM_NonAcad)
PCM_NonAcad

for (j in 1:4) {
  
  GMmin[j] <- GMmin[j]^(1/4)
  GMmodal[j] <- GMmodal[j]^(1/4)
  GMmax[j] <- GMmax[j]^(1/4)
}

Wt_Min <- NA
Wt_Modal <- NA
Wt_Max <- NA

for (j in 1:4) {
  Wt_Min[j] <- GMmin[j]/(sum(GMmax))
  Wt_Modal[j] <- GMmodal[j]/(sum(GMmodal))
  Wt_Max[j] <- GMmax[j]/(sum(GMmin))
}

Wt_NonAcad <- NA

for (j in 1:4) {
  
  Wt_NonAcad[j] <- (Wt_Min[j]+Wt_Modal[j]+Wt_Max[j])/3
}

Wt_NonAcad2 <- NA
for (j in 1:4) {
  
  Wt_NonAcad2[j] <- Wt_NonAcad[j]/sum(Wt_NonAcad)    
}


#9. Combining all weight ====

Wt <- cbind(Comb_Wt, Wt_Rese, Wt_Indu, Wt_Poli, Wt_Serv, Wt_Othe)
colnames(Wt) <- c("Total", "Wt_Acad", "Wt_Indu", "Wt_Poli", "Wt_Serv", "Wt_other")
row.names(Wt) <- c("OC", "SA", "TTR", "SA")

Wt <- data.frame(Wt)
Wt <- rownames_to_column(Wt, "Attributes") 

#10. Assigning Rank to all weights====
Wt <- Wt %>% 
  mutate(R_Total = rank(-Total), 
         R_Acad = rank(-Wt_Acad), 
         R_Indu = rank(-Wt_Indu),
         R_Poli = rank(-Wt_Poli), 
         R_Serv = rank(-Wt_Serv), 
         R_NonAcad = rank(-Wt_NonAcad))


Wt

#11. Plotting the weights ====
png('../results/fig1.png')
Stat_Pers <- Wt %>%
  data.frame() %>% 
  dplyr::select(Attributes:Wt_Serv) %>% 
  pivot_longer(3:6, names_to = "Stakeholders", values_to = "Weights") %>% 
  ggplot(aes(x=reorder(Attributes, Weights), y= Weights, group=Stakeholders, 
             shape=Stakeholders, color=Stakeholders)) +
  geom_point(stat = "identity", size = 3) +
  geom_line(linewidth= 1) + 
  theme_classic() +
  labs(x = "Attributes", title = "Attributes Weights Variation of Stakeholders")

Stat_Pers
dev.off()

png('../results/fig2.png')
AcadVsNonAcad <- Wt %>%
  data.frame() %>% 
  dplyr::select(Attributes:Wt_NonAcad) %>% 
  pivot_longer(c(3,7), names_to = "Stakeholders", values_to = "Weights") %>% 
  ggplot(aes(x=reorder(Attributes,Weights), y= Weights, fill= Stakeholders)) +
  geom_col(position = "dodge") + 
  theme_classic() +
  labs(x = "Attributes", title = "Attributes Weights Variation between Academician and Non-Academician")

AcadVsNonAcad
dev.off()

#12. Calculation of Unified Weights ====

Unified_Wt <- Wt %>% 
  dplyr::select(Attributes, Wt_Acad:Wt_Serv) %>%
  rowwise() %>% 
  mutate(Sum = sum(Wt_Acad + Wt_Indu + Wt_Poli + Wt_Serv)) %>% 
  mutate(Norm_Acad = Wt_Acad/Sum,
         Norm_Indu = Wt_Indu/Sum,
         Norm_Poli = Wt_Poli/Sum,
         Norm_Serv = Wt_Serv/Sum) %>% 
  mutate(W_Acad = Wt_Acad*Norm_Acad,
         W_Indu = Wt_Indu*Norm_Indu,
         W_Poli = Wt_Poli*Norm_Poli,
         W_Serv = Wt_Serv*Norm_Serv) %>% 
  mutate(W_Sum = W_Acad + W_Indu + W_Poli + W_Serv) %>% 
  dplyr::select(Attributes, W_Sum)

ColSum_Wt <- Unified_Wt %>%
  dplyr::select(W_Sum) %>% 
  colSums()

Norm_Unified_Wt <- round(Unified_Wt$W_Sum/ColSum_Wt , digits = 2) %>% 
  data.frame()

Wt_Unified <- cbind(round(Comb_Wt, digits = 2),Norm_Unified_Wt)

colnames(Wt_Unified) <- c("Total", "Unified")
rownames(Wt_Unified) <- c("OC", "SA", "TTR", "SA")

Wt_Unified

