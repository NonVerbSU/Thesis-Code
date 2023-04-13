# FACE READER

###############################################################################

# Reading automatization for Face Reader files. The script takes Face Reader 
# text files for each participant and returns two comma-delimited files, 
# containing: (1)the merged data of the dyad and (2) a correlation matrix. 


## -----------------------------------------------------------------------------
## Clear RStudio
rm(list = ls())  # Clear the global environment
graphics.off()  # Clear all plots
cat("\014")  # Clear the console

## -----------------------------------------------------------------------------
# (0) Configuration settings, IMPORTANT :

# Please, hard code configuration or load a configuration file:

npairs <- 6 ; # set how many pair you have!
nvariables <- 26; # # select how many variables you need from the data
final_variables <- c("Pair","Valence", "Arousal")

# Gets libraries and packages ready (?):

#install.packages("ggplot")
library(ggplot2)
#install.packages("xlsx")
library("xlsx")

## -----------------------------------------------------------------------------
## (1) Set path where data is:

getwd()
setwd("C:/Users/Usuario/OneDrive - UAM/PhD/2022_Estancia STOCKHOLM UNIVERSITY/DATA/simulation_data")

final_dataset<-  matrix(nrow = npairs, ncol = length(final_variables));

## (2) Import data: initiate loop to import data.

for (pair in 1:npairs){
  
    # import data from participant A
    d1 <- read.table(paste("Pair",as.character (pair), "A" ,".txt", sep=""), 
                     sep = '\t', header = TRUE,  skip = 12, na.strings = c("."))
    d1 <- d1 [, 1:nvariables]; # select how many variables you need from the data
    
    # import data from participant B
    d2 <- read.table(paste("Pair",as.character (pair), "B" ,".txt", sep=""),
                     sep = '\t', header = TRUE,  skip = 12, na.strings = c("."))
    d2 <- d2 [, 1:nvariables]; 
    
    # merge two data frames by time points
    d3 <- merge(d1,d2,by="Video.Time");

    # (3) generate a correlation matrix
    cor_mat <- cor(d3[,unlist(lapply(d3, is.numeric))], use = "complete.obs");
    
    # save dataframe and correlation matrix for this pair :
    write.csv(d3, file = paste("dataframe_pair_" , as.character (pair), 
                               ".csv", sep=""), row.names = FALSE)
    write.csv(cor_mat, file = paste("matrixco_pair_" , as.character (pair), 
                                    ".csv", sep=""), row.names = FALSE)
    
    
    # save correlation in the final dataset
    final_dataset[pair, 1] <- pair;
    final_dataset[pair, 2] <- cor_mat[23, 9]; # valence
    final_dataset[pair, 3] <- cor_mat[24, 10]; #arousal
    
    
    # (4) plot Arousal (as an example):
    test_plot <- ggscatter(d3, x = "Arousal.x", y = "Arousal.y", 
                           cor.coef = TRUE, cor.method = "pearson",
                           xlab = "Arousal - A", ylab = "Arousal - B")
    # save the plot
    ggsave(file = paste("graph_pair_", as.character (pair), ".png", sep=""), plot = test_plot, width = 6, height = 4, dpi = 300)
    
}

# save final_dataset
colnames(final_dataset)<-final_variables
write.csv(final_dataset, file = "final_dataset.csv", row.names = FALSE)

