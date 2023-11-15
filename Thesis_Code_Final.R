## -----------------------------------------------------------------------------
## Clear RStudio
rm(list = ls())  # Clear the global environment
graphics.off()  # Clear all plots
cat("\014")  # Clear the console
## -----------------------------------------------------------------------------

library(ggplot2)
library(readxl)
library(ggpubr)
library(tidyverse)
library(readxl)
library(broom)
library(openxlsx)
library(SUSY)
library(lubridate)
library(ggplot2)
library(reshape2)
library(psych)
library(papaja)
library(BayesFactor)

## -----------------------------------------------------------------------------
## -----------------------------------------------------------------------------
## Setting R to fetch the data from my saved location
getwd()

setwd("D:/T-Boy/SU/Year 2/Thesis/Facereader/15FPS/Lab")

##-----------------------------------------------------------------------------
###-----------------------------------------------------------------------------
  
# Importing the data:

# As multiple files reflecting multiple data frames

# Reading in multiple .txt files to create a list of data frames

my_files <- list.files(pattern = "\\.txt$")

my_files

my_data <- lapply(my_files, read.table, sep = '\t', header = TRUE,  skip = 12,
                  na.strings = c(".")) 
##-----------------------------------------------------------------------------
##-----------------------------------------------------------------------------

dyad_list <- list()

for (i in seq(1, length(my_data), by = 2)) {
  dyad <- merge.data.frame(x = my_data[[i]], y = my_data[[i+1]], 
                           by = "Video.Time", all = TRUE)
  dyad_list[[length(dyad_list) + 1]] <- dyad
  
}

## Concerting the Video.Time columns from a chr string to a time value in 
# format 00:00:00:000

for (i in seq_along(dyad_list)) {
  if ("Video.Time" %in% colnames(dyad_list[[i]])) {
    dyad_list[[i]]$Video.Time <- strptime(dyad_list[[i]]$Video.Time, "%H:%M:%OS")
    print(paste("Video.Time column of data frame", i, "converted successfully to time with milliseconds"))
    print(head(format(dyad_list[[i]]$Video.Time, "%H:%M:%OS3")))
  }
}

# Sanity check to see if the data has been transformed to time from chr

str(dyad_list[[1]]$Video.Time)

## Here is where our variables are selected in both instances of the time series

dyadlist_new <- lapply(dyad_list, `[`, c("Valence.x", 
                                         "Arousal.x", "Pitch.x", "Yaw.x", 
                                         "Roll.x", "Valence.y", "Arousal.y", 
                                         "Pitch.y", "Yaw.y", "Roll.y"))

dyadlist_new[1]

##-----------------------------------------------------------------------------
# Define a function that calculates SUSY for a specific column combination 
# in a data frame
# Define a function that calculates SUSY for a specific column combination 
# in a data frame
calculate_susy <- function(df, col1, col2) {
  susy(df[, c(col1, col2)], segment=30, Hz=15, maxlag = 5,
       permutation=FALSE, restrict.surrogates=FALSE,
       surrogates.total=1000)
}

# Extract p_value for each calculation

result_val <- lapply(dyadlist_new, calculate_susy, col1 = 1, col2 = 6)
result_aro <- lapply(dyadlist_new, calculate_susy, col1 = 2, col2 = 7)
result_pit <- lapply(dyadlist_new, calculate_susy, col1 = 3, col2 = 8)
result_yaw <- lapply(dyadlist_new, calculate_susy, col1 = 4, col2 = 9)
result_rol <- lapply(dyadlist_new, calculate_susy, col1 = 5, col2 = 10)
SUSY_results <- c(result_val, result_aro, result_pit, result_yaw, result_rol)

##-----------------------------------------------------------------------------
##------------------------------------------------------------------------------
# Convert each element of result_val to a data frame and combine them into a 
# single data frame

result_val_df <- do.call(rbind, lapply(result_val, data.frame))

# Add an index column to keep track of which element of result_val each 
# row corresponds to
result_val_df$index <- rep(seq_along(result_val), lengths(result_val))

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet for result_val
addWorksheet(wb, "result_val")

# Write result_val_df to the worksheet
writeData(wb, sheet = "result_val", x = result_val_df)

# Add a worksheet for each of the other result lists
for (i in c("result_aro", "result_pit", "result_yaw", "result_rol")) {
  addWorksheet(wb, i)
  writeData(wb, sheet = i, x = do.call(rbind, lapply(get(i), data.frame)))
}

# Save the workbook as an Excel file
saveWorkbook(wb, "result_allvar_15final.xlsx")

##-----------------------------------------------------------------------------
# Collapse all results for all dyads into a data frame based upon variable type

result_val_df <- do.call(rbind, lapply(result_val, data.frame))

print(result_val_df)

result_aro_df <- do.call(rbind, lapply(result_aro, data.frame))

result_pit_df <- do.call(rbind, lapply(result_pit, data.frame))

result_yaw_df <- do.call(rbind, lapply(result_yaw, data.frame))

result_rol_df <- do.call(rbind, lapply(result_rol, data.frame))

# Find the 'true' synchrony on the group level by taking the mean of Z abs
# and Z abs-pseudo and standardizes their difference by the standard deviation 
# of the Zabs−pseudo. As denoted in Tschacher and Meier (2020) - 
# DOI - 10.1080/10503307.2019.1612114

# Beginning with valence

mean_zabs_val <- mean(result_val_df$Z)
mean_zabs_pseudo_val <- mean(result_val_df$Z.Pseudo)

v_sd_zabs <- sd(result_val_df$Z)
v_sd_zabs_pseudo <- sd(result_val_df$Z.Pseudo)
sync_d_val <- (mean_zabs_val - mean_zabs_pseudo_val) / v_sd_zabs_pseudo

mean_znoabs_val <- mean(result_val_df$Z.noAbs.)
mean_znoabs_pseudo_val <- mean(result_val_df$Z.Pseudo.noAbs.)

v_sd_znoabs <- sd(result_val_df$Z.noAbs.)
v_sd_znoabs_pseudo <- sd(result_val_df$Z.Pseudo.noAbs.)
sync_d_val_no <- (mean_znoabs_val - mean_znoabs_pseudo_val) / v_sd_znoabs_pseudo

# Arousal

mean_zabs_aro <- mean(result_aro_df$Z)
mean_zabs_pseudo_aro <- mean(result_aro_df$Z.Pseudo)

a_sd_zabs <- sd(result_aro_df$Z)
a_sd_zabs_pseudo <- sd(result_aro_df$Z.Pseudo)
sync_d_aro <- (mean_zabs_aro - mean_zabs_pseudo_aro) / a_sd_zabs_pseudo

mean_znoabs_aro <- mean(result_aro_df$Z.noAbs.)
mean_znoabs_pseudo_aro <- mean(result_aro_df$Z.Pseudo.noAbs.)

a_sd_znoabs <- sd(result_aro_df$Z.noAbs.)
a_sd_znoabs_pseudo <- sd(result_aro_df$Z.Pseudo.noAbs.)
sync_d_aro_no <- (mean_znoabs_aro - mean_znoabs_pseudo_aro) / a_sd_znoabs_pseudo

# Pitch 

mean_zabs_pit <- mean(result_pit_df$Z)
mean_zabs_pseudo_pit <- mean(result_pit_df$Z.Pseudo)

p_sd_zabs <- sd(result_pit_df$Z)
p_sd_zabs_pseudo <- sd(result_pit_df$Z.Pseudo)
sync_d_pit <- (mean_zabs_pit - mean_zabs_pseudo_pit) / p_sd_zabs_pseudo

mean_znoabs_pit <- mean(result_pit_df$Z.noAbs.)
mean_znoabs_pseudo_pit <- mean(result_pit_df$Z.Pseudo.noAbs.)

p_sd_znoabs <- sd(result_pit_df$Z.noAbs.)
p_sd_znoabs_pseudo <- sd(result_pit_df$Z.Pseudo.noAbs.)
sync_d_pit_no <- (mean_znoabs_pit - mean_znoabs_pseudo_pit) / p_sd_znoabs_pseudo

# Yaw

mean_zabs_yaw <- mean(result_yaw_df$Z)
mean_zabs_pseudo_yaw <- mean(result_yaw_df$Z.Pseudo)

y_sd_zabs <- sd(result_yaw_df$Z)
y_sd_zabs_pseudo <- sd(result_yaw_df$Z.Pseudo)
sync_d_yaw <- (mean_zabs_yaw - mean_zabs_pseudo_yaw) / y_sd_zabs_pseudo

mean_znoabs_yaw <- mean(result_yaw_df$Z.noAbs.)
mean_znoabs_pseudo_yaw <- mean(result_yaw_df$Z.Pseudo.noAbs.)

y_sd_znoabs <- sd(result_yaw_df$Z.noAbs.)
y_sd_znoabs_pseudo <- sd(result_yaw_df$Z.Pseudo.noAbs.)
sync_d_yaw_no <- (mean_znoabs_yaw - mean_znoabs_pseudo_yaw) / y_sd_znoabs_pseudo

# Roll

mean_zabs_rol <- mean(result_rol_df$Z)
mean_zabs_pseudo_rol <- mean(result_rol_df$Z.Pseudo)

r_sd_zabs <- sd(result_rol_df$Z)
r_sd_zabs_pseudo <- sd(result_rol_df$Z.Pseudo)
sync_d_rol <- (mean_zabs_rol - mean_zabs_pseudo_rol) / r_sd_zabs_pseudo

mean_znoabs_rol <- mean(result_rol_df$Z.noAbs.)
mean_znoabs_pseudo_rol <- mean(result_rol_df$Z.Pseudo.noAbs.)

r_sd_znoabs <- sd(result_rol_df$Z.noAbs.)
r_sd_znoabs_pseudo <- sd(result_rol_df$Z.Pseudo.noAbs.)
sync_d_rol_no <- (mean_znoabs_rol - mean_znoabs_pseudo_rol) / r_sd_znoabs_pseudo

##------------------------------------------------------------------------------
##Calculating the mean, SD and p-values for ES of all variables

# Calculate mean for valence
mean_ES_val <- mean(result_val_df$ES)
mean_ESnoAbs_val <- mean(result_val_df$ES.noAbs.)

# Calculate standard deviation for valence
sd_ES_val <- sd(result_val_df$ES)
sd_ESnoAbs_val <- sd(result_val_df$ES.noAbs.)

# Perform t-test
t_test_ESvalab <- t.test(result_val_df$ES, mu = 0)
t_test_ESvalnoab <- t.test(result_val_df$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESvalab <- t_test_ESvalab$p.value
p_value_ESvalnoab <- t_test_ESvalnoab$p.value

##----------------------------

# Calculate mean for arousal
mean_ES_aro <- mean(result_aro_df$ES)
mean_ESnoAbs_aro <- mean(result_aro_df$ES.noAbs.)

# Calculate standard deviation for arousal
sd_ES_aro <- sd(result_aro_df$ES)
sd_ESnoAbs_aro <- sd(result_aro_df$ES.noAbs.)

## Calculate the p-value using the cumulative distribution function (CDF)

# Perform t-test
t_test_ESaroab <- t.test(result_aro_df$ES, mu = 0)
t_test_ESaronoab <- t.test(result_aro_df$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESaroab <- t_test_ESaroab$p.value
p_value_ESaronoab <- t_test_ESaronoab$p.value

##---------------------------------------

# Calculate mean for yaw
mean_ES_yaw <- mean(result_yaw_df$ES)
mean_ESnoAbs_yaw <- mean(result_yaw_df$ES.noAbs.)

# Calculate standard deviation for yaw
sd_ES_yaw <- sd(result_yaw_df$ES)
sd_ESnoAbs_yaw <- sd(result_yaw_df$ES.noAbs.)

# Perform t-test
t_test_ESyawab <- t.test(result_yaw_df$ES, mu = 0)
t_test_ESyawnoab <- t.test(result_yaw_df$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESyawab <- t_test_ESyawab$p.value
p_value_ESyawnoab <- t_test_ESyawnoab$p.value

##----------------------------------------------

# Calculate mean for pit
mean_ES_pit <- mean(result_pit_df$ES)
mean_ESnoAbs_pit <- mean(result_pit_df$ES.noAbs.)

# Calculate standard deviation for pit
sd_ES_pit <- sd(result_pit_df$ES)
sd_ESnoAbs_pit <- sd(result_pit_df$ES.noAbs.)

# Perform t-test
t_test_ESpitab <- t.test(result_pit_df$ES, mu = 0)
t_test_ESpitnoab <- t.test(result_pit_df$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESpitab <- t_test_ESpitab$p.value
p_value_ESpitnoab <- t_test_ESpitnoab$p.value

##---------------------------------------------------

# Calculate mean for roll
mean_ES_rol <- mean(result_rol_df$ES)
mean_ESnoAbs_rol <- mean(result_rol_df$ES.noAbs.)

# Calculate standard deviation for roll
sd_ES_rol <- sd(result_rol_df$ES)
sd_ESnoAbs_rol <- sd(result_rol_df$ES.noAbs.)

# Perform t-test
t_test_ESrolab <- t.test(result_rol_df$ES, mu = 0)
t_test_ESrolnoab <- t.test(result_rol_df$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESrolab <- t_test_ESrolab$p.value
p_value_ESrolnoab <- t_test_ESrolnoab$p.value

##------------------------------------------------------------------------------
# Calculate Hodges-Lehmann estimator for all variables as this does not rely on 
# assumptions of normality in the distribution of data
# Starting with valence

hl_result_val_abs <- wilcox.test(result_val_df$Z, result_aro_df$Z.Pseudo, 
                             conf.int = TRUE, conf.level = 0.95)

hl_result_val_noabs <- wilcox.test(result_val_df$Z.noAbs., 
                       result_val_df$Z.Pseudo.noAbs.,
                       conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_vabs <- hl_result_val_abs$estimate
hl_estimator_vnoabs <- hl_result_val_noabs$estimate

# Print Hodges-Lehmann estimator
print(hl_result_val_noabs)

# Arousal

hl_result_aro_abs <- wilcox.test(result_aro_df$Z, result_aro_df$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_aro_noabs <- wilcox.test(result_aro_df$Z.noAbs., 
                                   result_aro_df$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_aabs <- hl_result_aro_abs$estimate
hl_estimator_anoabs <- hl_result_aro_noabs$estimate

# Print Hodges-Lehmann estimator
print(hl_result_aro_noabs)

# Pitch

hl_result_pit_abs <- wilcox.test(result_pit_df$Z, result_pit_df$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_pit_noabs <- wilcox.test(result_pit_df$Z.noAbs., 
                                   result_pit_df$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_pabs <- hl_result_pit_abs$estimate
hl_estimator_pnoabs <- hl_result_pit_noabs$estimate

# Yaw

hl_result_yaw_abs <- wilcox.test(result_yaw_df$Z, result_yaw_df$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_yaw_noabs <- wilcox.test(result_yaw_df$Z.noAbs., 
                                   result_yaw_df$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_yabs <- hl_result_yaw_abs$estimate
hl_estimator_ynoabs <- hl_result_yaw_noabs$estimate

# Roll

hl_result_rol_abs <- wilcox.test(result_rol_df$Z, result_rol_df$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_rol_noabs <- wilcox.test(result_rol_df$Z.noAbs., 
                                   result_rol_df$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_rabs <- hl_result_rol_abs$estimate
hl_estimator_rnoabs <- hl_result_rol_noabs$estimate

##------------------------------------------------------------------------------
# Writing a table to represent the group level values of synchrony 

# Create a data frame with the means, standard deviations and p-values
sync_data <- data.frame(
  Variable = c("Valence", "Arousal", "Pitch", "Yaw", "Roll"),
  "Mean Z abs" = c(mean_zabs_val, mean_zabs_aro, mean_zabs_pit, 
                mean_zabs_yaw, mean_zabs_rol),
  "SD Z abs" = c(v_sd_zabs, a_sd_zabs, p_sd_zabs, y_sd_zabs, r_sd_zabs),
  "Mean Z abs Pseudo" = c(mean_zabs_pseudo_val, mean_zabs_pseudo_aro, 
                       mean_zabs_pseudo_pit, mean_zabs_pseudo_yaw, 
                       mean_zabs_pseudo_rol),
  "SD Z abs Pseudo" = c(v_sd_zabs_pseudo, a_sd_zabs_pseudo, p_sd_zabs_pseudo, 
                     y_sd_zabs_pseudo, r_sd_zabs_pseudo),
  "Cohen's d - Z abs" = c(sync_d_val, sync_d_aro, sync_d_pit, 
                          sync_d_yaw, sync_d_rol),
  "Hodeges-Lehmann (HL) - Z abs" = c(hl_estimator_vabs, hl_estimator_aabs, 
                      hl_estimator_pabs, hl_estimator_yabs, hl_estimator_rabs),
  "Mean ES" = c(mean_ES_val, mean_ES_aro, mean_ES_pit, mean_ES_yaw, mean_ES_rol),
  "SD ES" = c(sd_ES_val, sd_ES_aro, sd_ES_pit, sd_ES_yaw, sd_ES_rol),
  "P Values ES" = c(p_value_ESvalab, p_value_ESaroab, p_value_ESpitab, 
                    p_value_ESyawab, p_value_ESrolab),
  "Mean Z noabs" = c(mean_znoabs_val, mean_znoabs_aro, mean_znoabs_pit, 
                     mean_znoabs_yaw, mean_znoabs_rol),
  "SD Z noabs" = c(v_sd_znoabs, a_sd_znoabs, p_sd_znoabs, y_sd_znoabs, 
                   r_sd_znoabs),
  "Mean Z noabs Pseudo" = c(mean_znoabs_pseudo_val, mean_znoabs_pseudo_aro, 
                            mean_znoabs_pseudo_pit, mean_znoabs_pseudo_yaw, 
                            mean_znoabs_pseudo_rol),
  "SD Z noabs Pseudo" = c(v_sd_znoabs_pseudo, a_sd_znoabs_pseudo, 
                          p_sd_znoabs_pseudo, y_sd_znoabs_pseudo, 
                          r_sd_znoabs_pseudo),
  "Cohen's d - Z noabs" = c(sync_d_val_no, sync_d_aro_no, sync_d_pit_no, 
                            sync_d_yaw_no, sync_d_rol_no),
  "Hodeges-Lehmann (HL) - Z noabs" = c(hl_estimator_vnoabs, hl_estimator_anoabs, 
                hl_estimator_pnoabs, hl_estimator_ynoabs, hl_estimator_rnoabs),
  "Mean ES noabs" = c(mean_ESnoAbs_val, mean_ESnoAbs_aro, mean_ESnoAbs_pit, 
                      mean_ESnoAbs_yaw, mean_ESnoAbs_rol),
  "SD ES noabs" = c(sd_ESnoAbs_val, sd_ESnoAbs_aro, sd_ESnoAbs_pit, 
                    sd_ESnoAbs_yaw, sd_ESnoAbs_rol),
  "P Values ES noabs" = c(p_value_ESvalnoab, p_value_ESaronoab, p_value_ESpitnoab, 
                    p_value_ESyawnoab, p_value_ESrolnoab))
 
 
# Export the data frame as a table
write.table(sync_data, "lab_susy_results_Finalplease.csv", sep = ",", row.names = FALSE)

##-----------------------------------------------------------------------------
# Performing t-tests for each variable to compare the means of in-phase vs anti-phase
#
# t-test for Valence
val16 <- result_val_df[, 16]
val17 <- result_val_df[, 17]

# Perform t-test
t_test_val <- t.test(val16, val17, paired = TRUE)

# Print the results
print(t_test_val)

# Now for arousal
aro16 <- result_aro_df[, 16]
aro17 <- result_aro_df[, 17]

# Perform t-test
t_test_aro <- t.test(aro16, aro17, paired = TRUE)

# Print the results
print(t_test_aro)

# Now for pitch
pit16 <- result_pit_df[, 16]
pit17 <- result_pit_df[, 17]

# Perform t-test
t_test_pit<- t.test(pit16, pit17, paired = TRUE)

# Print the results
print(t_test_pit)

# Now for yaw
yaw16 <- result_yaw_df[, 16]
yaw17 <- result_yaw_df[, 17]

# Perform t-test
t_test_yaw<- t.test(yaw16, yaw17, paired = TRUE)

# Print the results
print(t_test_yaw)

# And finally for Roll
rol16 <- result_rol_df[, 16]
rol17 <- result_rol_df[, 17]

# Perform t-test
t_test_rol<- t.test(rol16, rol17, paired = TRUE)

# Print the results
print(t_test_rol)

## For sanity, R produces values as '4.712e-07' which can be read as 4.712 
## multiplied by 10 raised to the power of -7, which is equivalent to 
## 0.0000004712, so well below the p < 0.01 threshold
## This indicates in-phase synchrony across all conditions, showing synchrony!!!

##------------------------------------------------------------------------------
##------------------------------------------------------------------------------
# Performing a multivariate regression model for all five ESabs and ESnoabs 
# and the eight items constituting a 'rapport' measure 

#Firstly read in the SVI data

# Set the file path to the Excel document

file_path <- "C:/Users/tsamu/OneDrive/Documents/SU/Year 2/Thesis/Data/SVI Data/Survey_Stats_Rev.xlsx"

# Read the 'dyad' tab into a data frame
dyad_data <- read_excel(file_path, sheet = "Dyad")

# Setting the column names from row 1
colnames(dyad_data) <- as.character(unlist(dyad_data[1,]))

# Removing the first and last rows to just have the 20 dyads
svi_data <- dyad_data[-c(1,22:25),]

# Convert 'Process' and 'Relationship' columns to numeric
svi_data$Process <- as.numeric(as.character(svi_data$Process))
svi_data$Relationship <- as.numeric(as.character(svi_data$Relationship))

# Aggregating the 'Process' and 'Relationship' sections together for a measure 
# of 'rapport' 
svi_data$rapport <- rowSums(svi_data[, c("Process", "Relationship")], na.rm = TRUE)

mean_rapport <- mean(svi_data$rapport)

# Create a new data frame with the desired variables
merged_data <- cbind(svi_data$rapport, result_val_df$Z, result_aro_df$Z, result_pit_df$Z, result_yaw_df$Z,
                     result_rol_df$Z, result_val_df$Z.noAbs., result_aro_df$Z.noAbs., result_pit_df$Z.noAbs.,
                     result_yaw_df$Z.noAbs., result_rol_df$Z.noAbs.)

# Set appropriate column names
colnames(merged_data) <- c("Rapport", "AbsV", "AbsA", "AbsP", "AbsY",
                           "AbsR", "NonAbsV", "NonAbsA", "NonAbsP",
                           "NonAbsY", "NonAbsR")

# Calculate correlation coefficients
cor_matrix <- cor(merged_data)

# Set upper triangular elements to NA
cor_matrix[upper.tri(cor_matrix)] <- NA

# Round the correlation coefficients to two decimal places
cor_matrix_rounded <- round(cor_matrix, 2)

# Print the rounded correlation matrix
print(cor_matrix_rounded)

##Beautiful correlation matrix
# Generate a lighter palette
col <- colorRampPalette(c("#BB4444", "#EE9988", "#FFFFFF", "#77AADD", "#4477AA"))

corrplot(cor_matrix_rounded, method = "shade", shade.col = NA, tl.col = "black", tl.srt = 45,
         col = col(200), addCoef.col = "black", cl.pos = "n", order = "AOE")

####---------------------------------------------------------------------------
##----------------------------------------------------------------------------
##------------------------------------------------------------------------------
## Checking that the results are reproducible on an individual level, without 
# looping
# Merging the two two data frames from the list, represented as '1' & '2' to form 
# dyad_1

dyad_1 <- merge.data.frame(x = my_data[[1]], y = my_data[[2]], 
                           by = "Video.Time", all = TRUE)

## Making Video.Time a time variable

dyad_1$Video.Time <- strptime(dyad_1$Video.Time, "%H:%M:%OS")

## Performing the SUSY calculation 

library(SUSY)

val1 = susy(dyad_1[, c(10, 85)], segment=30, Hz=15, maxlag=5, 
            permutation=FALSE, restrict.surrogates=FALSE,
            surrogates.total=1000)


val1

##All values are the same for valence in dyad 1, now to test dyad 4 and arousal

dyad_4a <- as.data.frame(dyadlist_new[4])[, c(2, 7)]

dya_res <- susy(dyad_4a, segment=30, Hz=15, maxlag=5, 
                permutation=FALSE, restrict.surrogates=FALSE,
                surrogates.total=1000)
print(dya_res)

# All fine there too! Lastly, checking pitch dyad 11

dyad_11p <- as.data.frame(dyadlist_new[11])[, c(3, 8)]

dyp_res <- susy(dyad_11p, segment=30, Hz=15, maxlag=5, 
                permutation=FALSE, restrict.surrogates=FALSE,
                surrogates.total=1000)
print(dyp_res)

# All fine, the loop appears to generate meaningful values of SUSY.

###----------------------------------------------------------------------------
## -----------------------------------------------------------------------------
## -----------------------------------------------------------------------------

# Attempt to plot SUSY data, specifically Zobs as measure of synchrony
# Create a vector of time points corresponding to each calculation
time <- seq_along(dyadlist_new)

# Calculate Zobs values for each dyad
Zobs_val <- sapply(dyadlist_new, function(dyad) calculate_susy(dyad, col1 = 1, col2 = 6)$Zobs)
Zobs_aro <- sapply(dyadlist_new, function(dyad) calculate_susy(dyad, col1 = 2, col2 = 7)$Zobs)
Zobs_pit <- sapply(dyadlist_new, function(dyad) calculate_susy(dyad, col1 = 3, col2 = 8)$Zobs)
Zobs_yaw <- sapply(dyadlist_new, function(dyad) calculate_susy(dyad, col1 = 4, col2 = 9)$Zobs)
Zobs_rol <- sapply(dyadlist_new, function(dyad) calculate_susy(dyad, col1 = 5, col2 = 10)$Zobs)

# Convert Zobs values to numeric vectors
Zobs_val <- unlist(Zobs_val)
Zobs_aro <- unlist(Zobs_aro)
Zobs_pit <- unlist(Zobs_pit)
Zobs_yaw <- unlist(Zobs_yaw)
Zobs_rol <- unlist(Zobs_rol)

# Plot Zobs values over time
plot(time, Zobs_val, type = "l", xlab = "Time", ylab = "Zobs", col = "red", main = "Zobs over Time")
lines(time, Zobs_aro, col = "blue")
lines(time, Zobs_pit, col = "green")
lines(time, Zobs_yaw, col = "orange")
lines(time, Zobs_rol, col = "purple")
legend("topright", legend = c("Valence", "Arousal", "Pitch", "Yaw", "Roll"),
       col = c("red", "blue", "green", "orange", "purple"), lty = 1)

any(is.na(Zobs_val))

length(time)
length(Zobs_rol)

##------------------------------------------------------------------------------

##Plotting for Values of Z against Z pseudo for valence in dyad 10 
# Firstly close all open graphic devices to ensure the plot generates
dev.off()

# The create a new data frame for valence in dyad 10

dyad_10_valence <- dyadlist_new[[17]][, c(1, 6)]

# Performing the individual 'susy' calculation for valence in dyad 10

res = susy(dyad_10_valence, segment=30, Hz=15, maxlag=5, 
           permutation=FALSE, restrict.surrogates=FALSE,
           surrogates.total=1000)

# Creating new vectors which look at Z and Z pseudo scores for dyad 10

Z.obs <- res$`Valence.x-Valence.y`$data$Valence.x - 
  res$`Valence.x-Valence.y`$data$Valence.y -
  res$`Valence.x-Valence.y`$lagtimes2.data$meanccorrReal

Z.pseudo <- res$`Valence.x-Valence.y`$data$Valence.x - 
  res$`Valence.x-Valence.y`$data$Valence.y -
  res$`Valence.x-Valence.y`$lagtimes2.data$meanccorrPseudo

# Creating a vector of time values to plot valence against
time <- seq(0, length(Z.obs)/15, by = 1/15)

##The subset the data again

Z.obs_subset <- Z.obs[1:length(time)]
Z.pseudobs_subset <- Z.pseudo[1:length(time)]

# Find the minimum length of the two vectors
min_length <- min(length(time), length(Z.pseudo))

# Set the font family
par(family = "serif")

# Plot the two vectors with the same length
plot(time[1:min_length], Z.obs[1:min_length], type = "l", 
     xlab = "Time (seconds)", ylab = "Valence Synchrony", col = "blue", 
     lwd = 2)
lines(time[1:min_length], Z.pseudo[1:min_length], col = "red", lwd = 2)

# Add a legend to the plot
legend("topleft", legend = c("Z", "Z Pseudo"), lty = c(1, 1), 
       col = c("blue", "red"), bty = "n", cex = 0.8, text.font = 2)

## Plotting only the first 200 seconds
# create a logical vector indicating which values are less than or equal to 200
subset_idx <- time <= 200

# subset the time, Z.obs, and Z.pseudo vectors
subset_time <- time[subset_idx]
subset_Z.obs <- Z.obs[subset_idx]
subset_Z.pseudo <- Z.pseudo[subset_idx]

# plot the subsetted data
plot(subset_time, subset_Z.obs, type = "l", xlab = "Time (seconds)", 
     ylab = "Synchrony", col = "blue")
lines(subset_time, subset_Z.pseudo, col = "red")

##------------------------------------------------------------------------------
## Plotting the two absolute values of valence against one another, i.e. 
# Person A vs Person B
# Plotting for the first minute (60 seconds) of the interaction

dyad_10_valence_200 <- dyad_list[[17]][1:900, c("Video.Time", "Valence.x", "Valence.y")]

# Subsetting Video.Time to create
num_frames <- 900  # Number of frames in the subset
fps <- 15  # Frames per second
duration <- num_frames / fps  # Total duration in seconds
time_increment <- 1 / fps  # Time increment between frames

dyad_10_valence_200$Video.Time <- seq(0, duration - time_increment, by = time_increment)

# Now plotting the data
plot(dyad_10_valence_200$Video.Time, dyad_10_valence_200$Valence.x, type = "l", 
     col = "blue", 
     xlab = "Time (seconds)", ylab = "Valence", 
     ylim = c(-0.5, 1),
     main = "")
lines(dyad_10_valence_200$Video.Time, dyad_10_valence_200$Valence.y,
      col = "red")
legend("topright", legend = c("Interactant A", "Interactant B"), 
       col = c("blue",  "red"), lty = 1)
title(main = expression(italic("Valence across the first minute of interaction for Dyad 10")), adj = 0, line = 2)

## Plotting from observation 1000 to observation 1200 for both x and y in 
# absolute values of valence scores

dyad_10_valence_2 <- dyad_list[[10]][900:1799,
                                     c("Video.Time", "Valence.x", "Valence.y")]

# Subsetting Video.Time to create
num_frames <- 900  # Number of frames in the subset
fps <- 15  # Frames per second
duration <- num_frames / fps  # Total duration in seconds
time_increment <- 1 / fps  # Time increment between frames

dyad_10_valence_2$Video.Time <- seq(0, duration - time_increment, by = time_increment)

plot(dyad_10_valence_2$Video.Time, dyad_10_valence_2$Valence.x, type = "l", 
     col = "blue", 
     xlab = "Time (seconds)", ylab = "Valence", 
     ylim = c(-0.5, 1),
     main = "")
lines(dyad_10_valence_2$Video.Time, dyad_10_valence_2$Valence.y,
      col = "red")
legend("topleft", legend = c("Interactant A", "Interactant B"), col = c("blue",
                                                                        "red"), lty = 1)
title(main = expression(italic("Valence in Dyad 10 - The second minute")), adj = 0, line = 2)


## Plotting from observation 2000 to observation 2199 for both x and y in 
# absolute values of valence scores

dyad_10_valence_2000 <- dyad_list[[10]][1800:2599, c("Video.Time", "Valence.x",
                                                     "Valence.y")]

# Subsetting Video.Time to create
num_frames <- 900  # Number of frames in the subset
fps <- 15  # Frames per second
duration <- num_frames / fps  # Total duration in seconds
time_increment <- 1 / fps  # Time increment between frames

dyad_10_valence_2000$Video.Time <- seq(0, duration - time_increment, by = time_increment)
plot(dyad_10_valence_2000$Video.Time, dyad_10_valence_2000$Valence.x, type = "l", 
     col = "blue", 
     xlab = "Time (seconds)", ylab = "Valence", 
     ylim = c(-0.5, 1),
     main = "")
lines(dyad_10_valence_2000$Video.Time, dyad_10_valence_2000$Valence.y,
      col = "red")
legend("topright", legend = c("Interactant A", "Interactant B"), col = c("blue",
                                                                         "red"), lty = 1)
title(main = expression(italic("Valence in Dyad 10 - The third minute")), adj = 0, line = 2)

##------------------------------------------------------------------------------
##------------------------------------------------------------------------------


## Setting R to fetch the data from my saved location
getwd()

setwd("D:/T-Boy/SU/Year 2/Thesis/Facereader/15FPS/Zoom")

##-----------------------------------------------------------------------------
###-----------------------------------------------------------------------------
# Importing the data:

# As multiple files reflecting multiple data frames

# Reading in multiple .txt files to create a list of data frames

my_files_z <- list.files(pattern = "\\.txt$")

my_files_z

my_data_z <- lapply(my_files_z, read.table, sep = '\t', header = TRUE,  skip = 12,
                  na.strings = c(".")) 
##-----------------------------------------------------------------------------
##-----------------------------------------------------------------------------

dyad_list_z <- list()

for (i in seq(1, length(my_data_z), by = 2)) {
  dyad_z <- merge.data.frame(x = my_data_z[[i]], y = my_data_z[[i+1]], 
                           by = "Video.Time", all = TRUE)
  dyad_list_z[[length(dyad_list_z) + 1]] <- dyad_z
  
}

## Concerting the Video.Time columns from a chr string to a time value in 
# format 00:00:00:000

for (i in seq_along(dyad_list_z)) {
  if ("Video.Time" %in% colnames(dyad_list_z[[i]])) {
    dyad_list_z[[i]]$Video.Time <- strptime(dyad_list_z[[i]]$Video.Time, "%H:%M:%OS")
    print(paste("Video.Time column of data frame", i, "converted successfully to time with milliseconds"))
    print(head(format(dyad_list_z[[i]]$Video.Time, "%H:%M:%OS3")))
  }
}

# Sanity check to see if the data has been transformed to time from chr

str(dyad_list_z[[1]]$Video.Time)

## Here is where our variables are selected in both instances of the time series

dyadlist_new_z <- lapply(dyad_list_z, `[`, c("Valence.x", 
                                         "Arousal.x", "Pitch.x", "Yaw.x", 
                                         "Roll.x", "Valence.y", "Arousal.y", 
                                         "Pitch.y", "Yaw.y", "Roll.y"))

dyadlist_new_z[1]

##-----------------------------------------------------------------------------
# Define a function that calculates SUSY for a specific column combination 
# in a data frame
# Define a function that calculates SUSY for a specific column combination 
# in a data frame
calculate_susy <- function(df, col1, col2) {
  susy(df[, c(col1, col2)], segment=30, Hz=15, maxlag = 5,
       permutation=FALSE, restrict.surrogates=FALSE,
       surrogates.total=1000)
}

# Extract p_value for each calculation

result_val_z <- lapply(dyadlist_new_z, calculate_susy, col1 = 1, col2 = 6)
result_aro_z <- lapply(dyadlist_new_z, calculate_susy, col1 = 2, col2 = 7)
result_pit_z <- lapply(dyadlist_new_z, calculate_susy, col1 = 3, col2 = 8)
result_yaw_z <- lapply(dyadlist_new_z, calculate_susy, col1 = 4, col2 = 9)
result_rol_z <- lapply(dyadlist_new_z, calculate_susy, col1 = 5, col2 = 10)
SUSY_results_z <- c(result_val_z, result_aro_z, result_pit_z, result_yaw_z, 
                    result_rol_z)

##-----------------------------------------------------------------------------
##------------------------------------------------------------------------------
# Convert each element of result_val to a data frame and combine them into a 
# single data frame

result_val_df_z <- do.call(rbind, lapply(result_val_z, data.frame))

# Add an index column to keep track of which element of result_val each 
# row corresponds to
result_val_df_z$index <- rep(seq_along(result_val_z), lengths(result_val_z))

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet for result_val
addWorksheet(wb, "result_val_z")

# Write result_val_df to the worksheet
writeData(wb, sheet = "result_val_z", x = result_val_df_z)

# Add a worksheet for each of the other result lists
for (i in c("result_aro_z", "result_pit_z", "result_yaw_z", "result_rol_z")) {
  addWorksheet(wb, i)
  writeData(wb, sheet = i, x = do.call(rbind, lapply(get(i), data.frame)))
}

# Save the workbook as an Excel file
saveWorkbook(wb, "result_allvar_Z15rigen.xlsx")

##-----------------------------------------------------------------------------
# Collapse all results for all dyads into a data frame based upon variable type

result_val_df_z <- do.call(rbind, lapply(result_val_z, data.frame))

print(result_val_df_z)

result_aro_df_z <- do.call(rbind, lapply(result_aro_z, data.frame))

result_pit_df_z <- do.call(rbind, lapply(result_pit_z, data.frame))

result_yaw_df_z <- do.call(rbind, lapply(result_yaw_z, data.frame))

result_rol_df_z <- do.call(rbind, lapply(result_rol_z, data.frame))

# Find the 'true' synchrony on the group level by taking the mean of Z abs
# and Z abs-pseudo and standardizes their difference by the standard deviation 
# of the Zabs−pseudo. As denoted in Tschacher and Meier (2020) - 
# DOI - 10.1080/10503307.2019.1612114

# Beginning with valence

mean_zabs_val_z <- mean(result_val_df_z$Z)
mean_zabs_pseudo_val_z <- mean(result_val_df_z$Z.Pseudo)

v_sd_zabs_z <- sd(result_val_df_z$Z)
v_sd_zabs_pseudo_z <- sd(result_val_df_z$Z.Pseudo)
sync_d_val_z <- (mean_zabs_val_z - mean_zabs_pseudo_val_z) / v_sd_zabs_pseudo_z

mean_znoabs_val_z <- mean(result_val_df_z$Z.noAbs.)
mean_znoabs_pseudo_val_z <- mean(result_val_df_z$Z.Pseudo.noAbs.)

v_sd_znoabs_z <- sd(result_val_df_z$Z.noAbs.)
v_sd_znoabs_pseudo_z <- sd(result_val_df_z$Z.Pseudo.noAbs.)
sync_d_val_no_z <- (mean_znoabs_val_z - mean_znoabs_pseudo_val_z) / v_sd_znoabs_pseudo_z

# Arousal

mean_zabs_aro_z <- mean(result_aro_df_z$Z)
mean_zabs_pseudo_aro_z <- mean(result_aro_df_z$Z.Pseudo)

a_sd_zabs_z <- sd(result_aro_df_z$Z)
a_sd_zabs_pseudo_z <- sd(result_aro_df_z$Z.Pseudo)
sync_d_aro_z <- (mean_zabs_aro_z - mean_zabs_pseudo_aro_z) / a_sd_zabs_pseudo_z

mean_znoabs_aro_z <- mean(result_aro_df_z$Z.noAbs.)
mean_znoabs_pseudo_aro_z <- mean(result_aro_df_z$Z.Pseudo.noAbs.)

a_sd_znoabs_z <- sd(result_aro_df_z$Z.noAbs.)
a_sd_znoabs_pseudo_z <- sd(result_aro_df_z$Z.Pseudo.noAbs.)
sync_d_aro_no_z <- (mean_znoabs_aro_z - mean_znoabs_pseudo_aro_z) / a_sd_znoabs_pseudo_z

# Pitch 

mean_zabs_pit_z <- mean(result_pit_df_z$Z)
mean_zabs_pseudo_pit_z <- mean(result_pit_df_z$Z.Pseudo)

p_sd_zabs_z <- sd(result_pit_df_z$Z)
p_sd_zabs_pseudo_z <- sd(result_pit_df_z$Z.Pseudo)
sync_d_pit_z <- (mean_zabs_pit_z - mean_zabs_pseudo_pit_z) / p_sd_zabs_pseudo_z

mean_znoabs_pit_z <- mean(result_pit_df_z$Z.noAbs.)
mean_znoabs_pseudo_pit_z <- mean(result_pit_df_z$Z.Pseudo.noAbs.)

p_sd_znoabs_z <- sd(result_pit_df_z$Z.noAbs.)
p_sd_znoabs_pseudo_z <- sd(result_pit_df_z$Z.Pseudo.noAbs.)
sync_d_pit_no_z <- (mean_znoabs_pit_z - mean_znoabs_pseudo_pit_z) / p_sd_znoabs_pseudo_z

# Yaw

mean_zabs_yaw_z <- mean(result_yaw_df_z$Z)
mean_zabs_pseudo_yaw_z <- mean(result_yaw_df_z$Z.Pseudo)

y_sd_zabs_z <- sd(result_yaw_df_z$Z)
y_sd_zabs_pseudo_z <- sd(result_yaw_df_z$Z.Pseudo)
sync_d_yaw_z <- (mean_zabs_yaw_z - mean_zabs_pseudo_yaw_z) / y_sd_zabs_pseudo_z

mean_znoabs_yaw_z <- mean(result_yaw_df_z$Z.noAbs.)
mean_znoabs_pseudo_yaw_z <- mean(result_yaw_df_z$Z.Pseudo.noAbs.)

y_sd_znoabs_z <- sd(result_yaw_df_z$Z.noAbs.)
y_sd_znoabs_pseudo_z <- sd(result_yaw_df_z$Z.Pseudo.noAbs.)
sync_d_yaw_no_z <- (mean_znoabs_yaw_z - mean_znoabs_pseudo_yaw_z) / y_sd_znoabs_pseudo_z

# Roll

mean_zabs_rol_z <- mean(result_rol_df_z$Z)
mean_zabs_pseudo_rol_z <- mean(result_rol_df_z$Z.Pseudo)

r_sd_zabs_z <- sd(result_rol_df_z$Z)
r_sd_zabs_pseudo_z <- sd(result_rol_df_z$Z.Pseudo)
sync_d_rol_z <- (mean_zabs_rol_z - mean_zabs_pseudo_rol_z) / r_sd_zabs_pseudo_z

mean_znoabs_rol_z <- mean(result_rol_df_z$Z.noAbs.)
mean_znoabs_pseudo_rol_z <- mean(result_rol_df_z$Z.Pseudo.noAbs.)

r_sd_znoabs_z <- sd(result_rol_df_z$Z.noAbs.)
r_sd_znoabs_pseudo_z <- sd(result_rol_df_z$Z.Pseudo.noAbs.)
sync_d_rol_no_z <- (mean_znoabs_rol_z - mean_znoabs_pseudo_rol_z) / r_sd_znoabs_pseudo_z

##------------------------------------------------------------------------------
##Calculating the mean, SD and p-values for ES of all variables

# Calculate mean for valence
mean_ES_val_z <- mean(result_val_df_z$ES)
mean_ESnoAbs_val_z <- mean(result_val_df_z$ES.noAbs.)

# Calculate standard deviation for valence
sd_ES_val_z <- sd(result_val_df_z$ES)
sd_ESnoAbs_val_z <- sd(result_val_df_z$ES.noAbs.)

# Perform t-test
t_test_ESvalab_z <- t.test(result_val_df_z$ES, mu = 0)
t_test_ESvalnoab_z <- t.test(result_val_df_z$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESvalab_z <- t_test_ESvalab_z$p.value
p_value_ESvalnoab_z <- t_test_ESvalnoab_z$p.value

##----------------------------

# Calculate mean for arousal
mean_ES_aro_z <- mean(result_aro_df_z$ES)
mean_ESnoAbs_aro_z <- mean(result_aro_df_z$ES.noAbs.)

# Calculate standard deviation for arousal
sd_ES_aro_z <- sd(result_aro_df_z$ES)
sd_ESnoAbs_aro_z <- sd(result_aro_df_z$ES.noAbs.)

## Calculate the p-value using the cumulative distribution function (CDF)

# Perform t-test
t_test_ESaroab_z <- t.test(result_aro_df_z$ES, mu = 0)
t_test_ESaronoab_z <- t.test(result_aro_df_z$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESaroab_z <- t_test_ESaroab_z$p.value
p_value_ESaronoab_z <- t_test_ESaronoab_z$p.value

##---------------------------------------

# Calculate mean for yaw
mean_ES_yaw_z <- mean(result_yaw_df_z$ES)
mean_ESnoAbs_yaw_z <- mean(result_yaw_df_z$ES.noAbs.)

# Calculate standard deviation for yaw
sd_ES_yaw_z <- sd(result_yaw_df_z$ES)
sd_ESnoAbs_yaw_z <- sd(result_yaw_df_z$ES.noAbs.)

# Perform t-test
t_test_ESyawab_z <- t.test(result_yaw_df_z$ES, mu = 0)
t_test_ESyawnoab_z <- t.test(result_yaw_df_z$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESyawab_z <- t_test_ESyawab_z$p.value
p_value_ESyawnoab_z <- t_test_ESyawnoab_z$p.value

##----------------------------------------------

# Calculate mean for pit
mean_ES_pit_z <- mean(result_pit_df_z$ES)
mean_ESnoAbs_pit_z <- mean(result_pit_df_z$ES.noAbs.)

# Calculate standard deviation for pit
sd_ES_pit_z <- sd(result_pit_df_z$ES)
sd_ESnoAbs_pit_z <- sd(result_pit_df_z$ES.noAbs.)

# Perform t-test
t_test_ESpitab_z <- t.test(result_pit_df_z$ES, mu = 0)
t_test_ESpitnoab_z <- t.test(result_pit_df_z$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESpitab_z <- t_test_ESpitab_z$p.value
p_value_ESpitnoab_z <- t_test_ESpitnoab_z$p.value

##---------------------------------------------------

# Calculate mean for roll
mean_ES_rol_z <- mean(result_rol_df_z$ES)
mean_ESnoAbs_rol_z <- mean(result_rol_df_z$ES.noAbs.)

# Calculate standard deviation for roll
sd_ES_rol_z <- sd(result_rol_df_z$ES)
sd_ESnoAbs_rol_z <- sd(result_rol_df_z$ES.noAbs.)

# Perform t-test
t_test_ESrolab_z <- t.test(result_rol_df_z$ES, mu = 0)
t_test_ESrolnoab_z <- t.test(result_rol_df_z$ES.noAbs., mu = 0)

# Obtain p-values

# Extract the p-value from the t-test result
p_value_ESrolab_z <- t_test_ESrolab_z$p.value
p_value_ESrolnoab_z <- t_test_ESrolnoab_z$p.value

##------------------------------------------------------------------------------
# Calculate Hodges-Lehmann estimator for all variables as this does not rely on 
# assumptions of normality in the distribution of data
# Starting with valence

hl_result_val_abs_z <- wilcox.test(result_val_df$Z, result_aro_df$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_val_noabs_z <- wilcox.test(result_val_df$Z.noAbs., 
                                   result_val_df$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_vabs_z <- hl_result_val_abs$estimate
hl_estimator_vnoabs_z <- hl_result_val_noabs$estimate

# Print Hodges-Lehmann estimator
print(hl_result_val_noabs_z)

# Arousal

hl_result_aro_abs_z <- wilcox.test(result_aro_df_z$Z, result_aro_df_z$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_aro_noabs_z <- wilcox.test(result_aro_df_z$Z.noAbs., 
                                   result_aro_df_z$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_aabs_z <- hl_result_aro_abs$estimate
hl_estimator_anoabs_z <- hl_result_aro_noabs$estimate

# Print Hodges-Lehmann estimator
print(hl_result_aro_noabs_z)

# Pitch

hl_result_pit_abs_z <- wilcox.test(result_pit_df_z$Z, result_pit_df_z$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_pit_noabs_z <- wilcox.test(result_pit_df_z$Z.noAbs., 
                                   result_pit_df_z$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_pabs_z <- hl_result_pit_abs_z$estimate
hl_estimator_pnoabs_z <- hl_result_pit_noabs_z$estimate

# Yaw

hl_result_yaw_abs_z <- wilcox.test(result_yaw_df_z$Z, result_yaw_df_z$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_yaw_noabs_z <- wilcox.test(result_yaw_df_z$Z.noAbs., 
                                   result_yaw_df_z$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_yabs_z <- hl_result_yaw_abs_z$estimate
hl_estimator_ynoabs_z <- hl_result_yaw_noabs_z$estimate

# Roll

hl_result_rol_abs_z <- wilcox.test(result_rol_df_z$Z, result_rol_df_z$Z.Pseudo, 
                                 conf.int = TRUE, conf.level = 0.95)

hl_result_rol_noabs_z <- wilcox.test(result_rol_df_z$Z.noAbs., 
                                   result_rol_df_z$Z.Pseudo.noAbs.,
                                   conf.int = TRUE, conf.level = 0.95)

# Extract Hodges-Lehmann estimator
hl_estimator_rabs_z <- hl_result_rol_abs_z$estimate
hl_estimator_rnoabs_z <- hl_result_rol_noabs_z$estimate

##------------------------------------------------------------------------------
# Writing a table to represent the group level values of synchrony 

# Create a data frame with the means, standard deviations and p-values
sync_data_z <- data.frame(
  Variable = c("Valence", "Arousal", "Pitch", "Yaw", "Roll"),
  "Mean Z abs" = c(mean_zabs_val_z, mean_zabs_aro_z, mean_zabs_pit_z, 
                   mean_zabs_yaw_z, mean_zabs_rol_z),
  "SD Z abs" = c(v_sd_zabs_z, a_sd_zabs_z, p_sd_zabs_z, y_sd_zabs_z, r_sd_zabs_z),
  "Mean Z abs Pseudo" = c(mean_zabs_pseudo_val_z, mean_zabs_pseudo_aro_z, 
                          mean_zabs_pseudo_pit_z, mean_zabs_pseudo_yaw_z, 
                          mean_zabs_pseudo_rol_z),
  "SD Z abs Pseudo" = c(v_sd_zabs_pseudo_z, a_sd_zabs_pseudo_z, p_sd_zabs_pseudo_z, 
                        y_sd_zabs_pseudo_z, r_sd_zabs_pseudo_z),
  "Cohen's d - Z abs" = c(sync_d_val_z, sync_d_aro_z, sync_d_pit_z, 
                          sync_d_yaw_z, sync_d_rol_z),
  "Hodeges-Lehmann (HL) - Z abs" = c(hl_estimator_vabs_z, hl_estimator_aabs_z, 
                                     hl_estimator_pabs_z, hl_estimator_yabs_z, hl_estimator_rabs_z),
  "Mean ES" = c(mean_ES_val_z, mean_ES_aro_z, mean_ES_pit_z, mean_ES_yaw_z, mean_ES_rol_z),
  "SD ES" = c(sd_ES_val, sd_ES_aro, sd_ES_pit, sd_ES_yaw, sd_ES_rol),
  "P Values ES" = c(p_value_ESvalab_z, p_value_ESaroab_z, p_value_ESpitab_z, 
                    p_value_ESyawab_z, p_value_ESrolab_z),
  "Mean Z noabs" = c(mean_znoabs_val_z, mean_znoabs_aro_z, mean_znoabs_pit_z, 
                     mean_znoabs_yaw_z, mean_znoabs_rol_z),
  "SD Z noabs" = c(v_sd_znoabs_z, a_sd_znoabs_z, p_sd_znoabs_z, y_sd_znoabs_z, 
                   r_sd_znoabs_z),
  "Mean Z noabs Pseudo" = c(mean_znoabs_pseudo_val_z, mean_znoabs_pseudo_aro_z, 
                            mean_znoabs_pseudo_pit_z, mean_znoabs_pseudo_yaw_z, 
                            mean_znoabs_pseudo_rol_z),
  "SD Z noabs Pseudo" = c(v_sd_znoabs_pseudo_z, a_sd_znoabs_pseudo_z, 
                          p_sd_znoabs_pseudo_z, y_sd_znoabs_pseudo_z, 
                          r_sd_znoabs_pseudo_z),
  "Cohen's d - Z noabs" = c(sync_d_val_no_z, sync_d_aro_no_z, sync_d_pit_no_z, 
                            sync_d_yaw_no_z, sync_d_rol_no_z),
  "Hodeges-Lehmann (HL) - Z noabs" = c(hl_estimator_vnoabs_z, hl_estimator_anoabs_z, 
                                       hl_estimator_pnoabs_z, hl_estimator_ynoabs_z, 
                                       hl_estimator_rnoabs_z),
  "Mean ES noabs" = c(mean_ESnoAbs_val_z, mean_ESnoAbs_aro_z, mean_ESnoAbs_pit_z, 
                      mean_ESnoAbs_yaw_z, mean_ESnoAbs_rol_z),
  "SD ES noabs" = c(sd_ESnoAbs_val_z, sd_ESnoAbs_aro_z, sd_ESnoAbs_pit_z, 
                    sd_ESnoAbs_yaw_z, sd_ESnoAbs_rol_z),
  "P Values ES noabs" = c(p_value_ESvalnoab_z, p_value_ESaronoab_z, p_value_ESpitnoab_z, 
                          p_value_ESyawnoab_z, p_value_ESrolnoab_z))


# Export the data frame as a table
write.table(sync_data, "lab_susy_results_ZoomFinr.csv", sep = ",", row.names = FALSE)

##-----------------------------------------------------------------------------
# Performing t-tests for each variable to compare the means of in-phase vs anti-phase
#
# t-test for Valence
val16_z <- result_val_df_z[, 16]
val17_z <- result_val_df_z[, 17]

# Perform t-test
t_test_val_z <- t.test(val16_z, val17_z, paired = TRUE)

# Print the results
print(t_test_val_z)

# Now for arousal
aro16_z <- result_aro_df_z[, 16]
aro17_z <- result_aro_df_z[, 17]

# Perform t-test
t_test_aro_z <- t.test(aro16_z, aro17_z, paired = TRUE)

# Print the results
print(t_test_aro_z)

# Now for pitch
pit16_z <- result_pit_df_z[, 16]
pit17_z <- result_pit_df_z[, 17]

# Perform t-test
t_test_pit_z<- t.test(pit16_z, pit17_z, paired = TRUE)

# Print the results
print(t_test_pit_z)

# Now for yaw
yaw16_z <- result_yaw_df_z[, 16]
yaw17_z <- result_yaw_df_z[, 17]

# Perform t-test
t_test_yaw_z<- t.test(yaw16_z, yaw17_z, paired = TRUE)

# Print the results
print(t_test_yaw_z)

# And finally for Roll
rol16_z <- result_rol_df_z[, 16]
rol17_z <- result_rol_df_z[, 17]

# Perform t-test
t_test_rol_z<- t.test(rol16_z, rol17_z, paired = TRUE)

# Print the results
print(t_test_rol_z)

## For sanity, R produces values as '4.712e-07' which can be read as 4.712 
## multiplied by 10 raised to the power of -7, which is equivalent to 
## 0.0000004712, so well below the p < 0.01 threshold
## This indicates in-phase synchrony across all conditions, showing synchrony!!!

##------------------------------------------------------------------------------
##------------------------------------------------------------------------------
# Performing a multivariate regression model for all five ESabs and ESnoabs 
# and the eight items constituting a 'rapport' measure 

#Firstly read in the SVI data

# Set the file path to the Excel document

setwd("C:/Users/tsamu/OneDrive/Documents/SU/Year 2/Thesis/Data/SVI Data")

# Load the readxl package
library(readxl)

# Assuming your Excel file is named 'data.xlsx' (replace 'data.xlsx' with your actual file name)
Zoom_SVI_z <- read_excel("SVI_data_Zoom.xlsx", skip = 1, col_names = FALSE)


# Set the correct column names
colnames(Zoom_SVI_z) <- c("DateTime", "Name1", "Name2", "Number1", "Gender", "Number2", "Number3", "Number4", "Number5", "Number6", "Number7", "Number8", "Number9")

# View the updated data frame
print(Zoom_SVI_z)

# Load the dplyr package
library(dplyr)
library(corrplot)

# Assuming your data frame is named 'data' (replace 'data' with your actual data frame name)
Zoom_SVI_z <- mutate_at(Zoom_SVI_z, vars(6:13), as.numeric)

# View the updated data frame
print(Zoom_SVI_z)

# Assuming your data frame is named 'data' (replace 'data' with your actual data frame name)
Zoom_SVI_z$Rapport <- rowSums(Zoom_SVI_z[, 6:13])

# View the updated data frame
print(Zoom_SVI_z)

# Assuming your data frame is named 'data' (replace 'data' with your actual data frame name)
# Create a new column to identify the pairs
Zoom_SVI_z <- mutate(Zoom_SVI_z, Pair = ceiling(row_number() / 2))

# Group the data by Pair and summarize to merge the rows within each pair
Zoom_SVI_z <- Zoom_SVI_z %>%
  group_by(Pair) %>%
  summarize_all(sum)

# View the merged data frame
print(Zoom_SVI_z)

# Assuming your data frame is named 'data' (replace 'data' with your actual data frame name)
# Group the data by Pair and calculate the sum of the Rapport scores
rapport_summary_z <- Zoom_SVI_z %>%
  group_by(Pair) %>%
  summarize(Rapport_Sum = sum(Rapport))

# View the summary data frame
print(rapport_summary_z)

# Create a new data frame with the desired variables
cormat_data_z <- cbind(rapport_summary_z$Rapport_Sum, result_val_df_z$Z,
                       result_aro_df_z$Z, result_pit_df_z$Z, result_yaw_df_z$Z,
                     result_rol_df_z$Z, result_val_df_z$Z.noAbs., 
                     result_aro_df_z$Z.noAbs., result_pit_df_z$Z.noAbs.,
                     result_yaw_df_z$Z.noAbs., result_rol_df_z$Z.noAbs.)

# Set appropriate column names
colnames(cormat_data_z) <- c("Rapport", "AbsV", "AbsA", "AbsP", "AbsY",
                             "AbsR", "NonAbsV", "NonAbsA", "NonAbsP",
                             "NonAbsY", "NonAbsR")

# Calculate correlation coefficients
cor_matrix_z <- cor(cormat_data_z)

##Beautiful correlation matrix
# Generate a lighter palette
col_z <- colorRampPalette(c("#BB4444", "#EE9988", "#FFFFFF", "#77AADD", "#4477AA"))

corrplot(cor_matrix_z, method = "shade", shade.col = NA, tl.col = "black", tl.srt = 45,
         col = col(200), addCoef.col = "black", cl.pos = "n", order = "AOE")


##------------------------------------------------------------------------------
## Now to calculate the differences between Face-to-Face and Zoom conditions
## Performing a MANOVA on mean effects sizes for Esabs and ESnoabs

library(car)

# Create a data frame
mydata_manova <- data.frame(valence = c(mean_ES_val, mean_ESnoAbs_val, mean_ES_val_z, mean_ESnoAbs_val_z),
                     arousal = c(mean_ES_aro, mean_ESnoAbs_aro, mean_ES_aro_z, mean_ESnoAbs_aro_z),
                     pitch = c(mean_ES_pit, mean_ESnoAbs_pit, mean_ES_pit_z, mean_ESnoAbs_pit_z),
                     yaw = c(mean_ES_yaw, mean_ESnoAbs_yaw, mean_ES_yaw_z, mean_ESnoAbs_yaw_z),
                     roll = c(mean_ES_rol, mean_ESnoAbs_rol, mean_ES_rol_z, mean_ESnoAbs_rol_z),
                     condition = c("Face-to-Face", "Face-to-Face", "Zoom", "Zoom"))

# Create a data frame
# Create a data frame
mydata_manova <- data.frame(
  valence_absolute = c(mean_ES_val, mean_ES_val_z),
  valence_nonabs = c(mean_ESnoAbs_val, mean_ESnoAbs_val_z),
  arousal_absolute = c(mean_ES_aro, mean_ES_aro_z),
  arousal_nonabs = c(mean_ESnoAbs_aro, mean_ESnoAbs_aro_z),
  pitch_absolute = c(mean_ES_pit, mean_ES_pit_z),
  pitch_nonabs = c(mean_ESnoAbs_pit, mean_ESnoAbs_pit_z),
  yaw_absolute = c(mean_ES_yaw, mean_ES_yaw_z),
  yaw_nonabs = c(mean_ESnoAbs_yaw, mean_ESnoAbs_yaw_z),
  roll_absolute = c(mean_ES_rol, mean_ES_rol_z),
  roll_nonabs = c(mean_ESnoAbs_rol, mean_ESnoAbs_rol_z),
  condition = c("Face-to-Face", "Zoom")
)

# Print the data frame
print(mydata_manova)

#Performing Wilcoxon rank in order to account for potential n

# Wilcoxon rank-sum test for valence ESabs
wilcox.test(valence_absolute ~ condition, data = mydata_manova)

#data:  valence_absolute by condition
#W = 1, p-value = 1
#alternative hypothesis: true location shift is not equal to 0

#The code wilcox.test(valence ~ condition, data = mydata_manova) 
#performs a Wilcoxon rank-sum test to compare the distributions of the 
#variable "valence" between the two conditions ("Face-to-Face" and "Zoom") in 
#the data frame mydata_manova.

#The output of the test indicates the following:
  
#W: The test statistic is equal to 1.
#p-value: The p-value associated with the test is 0.6667.
#Alternative hypothesis: The alternative hypothesis for the test is that the 
#true location shift (median difference) between the two conditions is not 
#equal to 0.
#Interpretation: Based on the p-value of 0.6667, which is greater than the 
#typical significance level of 0.05, we do not have sufficient evidence to 
#reject the null hypothesis. Therefore, we do not have enough evidence to 
#conclude that there is a significant difference in the "valence" variable 
#between the "Face-to-Face" and "Zoom" conditions.

# Wilcoxon rank-sum test for valence ESnoabs
wilcox.test(valence_nonabs ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for arousal Eabs
wilcox.test(arousal_absolute ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for arousal_nonabs
wilcox.test(arousal_nonabs ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for pitch_absolute
wilcox.test(pitch_absolute ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for pitch_nonabs
wilcox.test(pitch_nonabs ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for yaw_absolute
wilcox.test(yaw_absolute ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for yaw_nonabs
wilcox.test(yaw_nonabs ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for roll_absolute
wilcox.test(roll_absolute ~ condition, data = mydata_manova)

# Wilcoxon rank-sum test for roll_nonabs
wilcox.test(roll_nonabs ~ condition, data = mydata_manova)

##------------------------------------------------------------------------------
##Now performing t-tests, firstly for absolute values across conditions

# Independent samples t-test for valence_absolute
t.test(result_val_df$Z, result_val_df_z$Z)

# Independent samples t-test for arousal
t.test(result_aro_df$Z, result_aro_df_z$Z)

# Independent samples t-test for pitch
t.test(result_pit_df$Z, result_pit_df_z$Z)

# Independent samples t-test for yaw
t.test(result_yaw_df$Z, result_yaw_df_z$Z)

# Independent samples t-test for roll
t.test(result_rol_df$Z, result_rol_df_z$Z)

#------------------------------------------------------------------------------

#Now for non-absolute values

# Independent samples t-test for valence without absolute values
t.test(result_val_df$ES.noAbs., result_val_df_z$ES.noAbs.)

# Independent samples t-test for arousal without absolute values
t.test(result_aro_df$ES.noAbs., result_aro_df_z$ES.noAbs.)

# Independent samples t-test for pitch without absolute values
t.test(result_pit_df$ES.noAbs., result_pit_df_z$ES.noAbs.)

# Independent samples t-test for yaw without absolute values
t.test(result_yaw_df$ES.noAbs., result_yaw_df_z$ES.noAbs.)

# Independent samples t-test for roll without absolute values
t.test(result_rol_df$ES.noAbs., result_rol_df_z$ES.noAbs.)

# Perform t-tests and store the results
t_val <- t.test(result_val_df$ES, result_val_df_z$ES)
t_aro <- t.test(result_aro_df$ES, result_aro_df_z$ES)
t_pit <- t.test(result_pit_df$ES, result_pit_df_z$ES)
t_yaw <- t.test(result_yaw_df$ES, result_yaw_df_z$ES)
t_rol <- t.test(result_rol_df$ES, result_rol_df_z$ES)

t_val_noabs <- t.test(result_val_df$ES.noAbs., result_val_df_z$ES.noAbs.)
t_aro_noabs <- t.test(result_aro_df$ES.noAbs., result_aro_df_z$ES.noAbs.)
t_pit_noabs <- t.test(result_pit_df$ES.noAbs., result_pit_df_z$ES.noAbs.)
t_yaw_noabs <- t.test(result_yaw_df$ES.noAbs., result_yaw_df_z$ES.noAbs.)
t_rol_noabs <- t.test(result_rol_df$ES.noAbs., result_rol_df_z$ES.noAbs.)

# Create a data frame to store the results
t_test_results <- data.frame(
  Variable = c("Valence ESAbs", "Arousal Abs", "Pitch Abs", "Yaw Abs", "Roll Abs",
               "Valence ESnoabs", "Arousal ESnoabs", "Pitch ESnoabs", 
               "Yaw ESnoabs", "Roll ESnoabs"),
  t_value = c(t_val$statistic, t_aro$statistic, t_pit$statistic, t_yaw$statistic, t_rol$statistic,
              t_val_noabs$statistic, t_aro_noabs$statistic, t_pit_noabs$statistic, t_yaw_noabs$statistic, t_rol_noabs$statistic),
  df = c(t_val$parameter, t_aro$parameter, t_pit$parameter, t_yaw$parameter, t_rol$parameter,
         t_val_noabs$parameter, t_aro_noabs$parameter, t_pit_noabs$parameter, t_yaw_noabs$parameter, t_rol_noabs$parameter),
  p_value = c(t_val$p.value, t_aro$p.value, t_pit$p.value, t_yaw$p.value, t_rol$p.value,
              t_val_noabs$p.value, t_aro_noabs$p.value, t_pit_noabs$p.value, t_yaw_noabs$p.value, t_rol_noabs$p.value)
)

# Print the data frame
print(t_test_results)

##-----------------------------
## Now performing the Wilcoxon, to test again the possibility of outliers and 
## non-normal distribution of data 

# Perform Wilcoxon signed-rank tests and store the results
wilcox_val <- wilcox.test(result_val_df$ES, result_val_df_z$ES)
wilcox_aro <- wilcox.test(result_aro_df$ES, result_aro_df_z$ES)
wilcox_pit <- wilcox.test(result_pit_df$ES, result_pit_df_z$ES)
wilcox_yaw <- wilcox.test(result_yaw_df$ES, result_yaw_df_z$ES)
wilcox_rol <- wilcox.test(result_rol_df$ES, result_rol_df_z$ES)

wilcox_val_noabs <- wilcox.test(result_val_df$ES.noAbs., result_val_df_z$ES.noAbs.)
wilcox_aro_noabs <- wilcox.test(result_aro_df$ES.noAbs., result_aro_df_z$ES.noAbs.)
wilcox_pit_noabs <- wilcox.test(result_pit_df$ES.noAbs., result_pit_df_z$ES.noAbs.)
wilcox_yaw_noabs <- wilcox.test(result_yaw_df$ES.noAbs., result_yaw_df_z$ES.noAbs.)
wilcox_rol_noabs <- wilcox.test(result_rol_df$ES.noAbs., result_rol_df_z$ES.noAbs.)

# Create a data frame to store the results
wilcox_test_results <- data.frame(
  Variable = c("Valence ESAbs", "Arousal Abs", "Pitch Abs", "Yaw Abs", "Roll Abs",
               "Valence ESnoabs", "Arousal SEnoabs", "Pitch ESnoabs", 
               "Yaw ESnoabs", "Roll ESnoabs"),
  "W value" = c(wilcox_val$statistic, wilcox_aro$statistic, wilcox_pit$statistic, wilcox_yaw$statistic, wilcox_rol$statistic,
              wilcox_val_noabs$statistic, wilcox_aro_noabs$statistic, wilcox_pit_noabs$statistic, wilcox_yaw_noabs$statistic, wilcox_rol_noabs$statistic),
  "P-value" = c(wilcox_val$p.value, wilcox_aro$p.value, wilcox_pit$p.value, wilcox_yaw$p.value, wilcox_rol$p.value,
              wilcox_val_noabs$p.value, wilcox_aro_noabs$p.value, wilcox_pit_noabs$p.value, wilcox_yaw_noabs$p.value, wilcox_rol_noabs$p.value)
)

# Print the data frame
print(wilcox_test_results)

##---------------------------------------------------------------------------
library(BayesFactor)

# Create a data frame
mydata_bayes <- data.frame(
  valence_absolute = c(mean_ES_val, mean_ES_val_z),
  valence_nonabs = c(mean_ESnoAbs_val, mean_ESnoAbs_val_z),
  arousal_absolute = c(mean_ES_aro, mean_ES_aro_z),
  arousal_nonabs = c(mean_ESnoAbs_aro, mean_ESnoAbs_aro_z),
  pitch_absolute = c(mean_ES_pit, mean_ES_pit_z),
  pitch_nonabs = c(mean_ESnoAbs_pit, mean_ESnoAbs_pit_z),
  yaw_absolute = c(mean_ES_yaw, mean_ES_yaw_z),
  yaw_nonabs = c(mean_ESnoAbs_yaw, mean_ESnoAbs_yaw_z),
  roll_absolute = c(mean_ES_rol, mean_ES_rol_z),
  roll_nonabs = c(mean_ESnoAbs_rol, mean_ESnoAbs_rol_z),
  condition = c("Face-to-Face", "Zoom")
)

mydata_bayes$condition <- factor(mydata_bayes$condition)
str(mydata_bayes)

bf_val_abs <- ttestBF(valence_absolute ~ condition, data = mydata_bayes)

str(mydata_bayes)
bf_val_abs <- ttestBF(x = mydata_bayes$valence_absolute, 
                      y = mydata_bayes$condition)
print(bf_val_abs)

bf_val_abs <- ttestBF(x = mydata_bayes$valence_absolute, y = mydata_bayes$condition)
bf_val_noabs <- ttestBF(x = mydata_bayes$valence_nonabs, y = mydata_bayes$condition)
bf_aro_abs <- ttestBF(x = mydata_bayes$arousal_absolute, y = mydata_bayes$condition)
bf_aro_noabs <- ttestBF(x = mydata_bayes$arousal_nonabs, y = mydata_bayes$condition)
bf_pit_abs <- ttestBF(x = mydata_bayes$pitch_absolute, y = mydata_bayes$condition)
bf_pit_noabs <- ttestBF(x = mydata_bayes$pitch_nonabs, y = mydata_bayes$condition)
bf_yaw_abs <- ttestBF(x = mydata_bayes$yaw_absolute, y = mydata_bayes$condition)
bf_yaw_noabs <- ttestBF(x = mydata_bayes$yaw_nonabs, y = mydata_bayes$condition)
bf_rol_abs <- ttestBF(x = mydata_bayes$roll_absolute, y = mydata_bayes$condition)
bf_rol_noabs <- ttestBF(x = mydata_bayes$roll_nonabs, y = mydata_bayes$condition)

str(bf_val_abs)

# Print the Bayes factors
print(bf_val_abs)
print(bf_val_noabs)
print(bf_aro_abs)
print(bf_aro_noabs)
print(bf_pit_abs)
print(bf_pit_noabs)
print(bf_yaw_abs)
print(bf_yaw_noabs)
print(bf_rol_abs)
print(bf_rol_noabs)

# Create a data frame with the results
bayes_results <- data.frame(
  Variable = c("valence_absolute", "valence_nonabs", "arousal_absolute", "arousal_nonabs",
               "pitch_absolute", "pitch_nonabs", "yaw_absolute", "yaw_nonabs",
               "roll_absolute", "roll_nonabs"),
  BF = c(bf_val_abs$bf, bf_val_noabs$bf, bf_aro_abs$bf, bf_aro_noabs$bf, bf_pit_abs$bf,
         bf_pit_noabs$bf, bf_yaw_abs$bf, bf_yaw_noabs$bf, bf_rol_abs$bf, bf_rol_noabs$bf),
  Error = format(c(bf_val_abs$error, bf_val_noabs$error, bf_aro_abs$error, bf_aro_noabs$error, bf_pit_abs$error,
                   bf_pit_noabs$error, bf_yaw_abs$error, bf_yaw_noabs$error, bf_rol_abs$error, bf_rol_noabs$error),
                 scientific = FALSE)
)

# Export the table to Word in APA format
library(apaTables)
write.docx(apa.table(results, title = "Bayes Factor Results", caption = "Bayes factors for each variable"))



 ##-----------------------------------------------------------------------------
#Exporting as an APA table 

# Install the required packages if not already installed
install.packages("flextable")
install.packages("officer")

# Load the required libraries
library(flextable)
library(officer)

# Create the data frame
wilcox_test_results <- data.frame(
  Variable = c("Valence ESAbs", "Arousal ESAbs", "Pitch ESAbs", "Yaw ESAbs", "Roll ESAbs",
               "Valence ESnoabs", "Arousal ESnoabs", "Pitch ESnoabs", "Yaw ESnoabs", "Roll ESnoabs"),
  W_value = c(94, 84, 133, 112, 102, 78, 101, 114, 92, 127),
  P_value = c(0.53, 0.30, 0.36, 0.95, 0.76, 0.20, 0.73, 0.89, 0.48, 0.50)
)

# Create a flextable object
ft <- flextable(wilcox_test_results)

# Create the data frame with rounded p-values
wilcox_test_results$P_value <- round(wilcox_test_results$P_value, 2)

# Create a flextable object
ft <- flextable(wilcox_test_results)

# Set column names
set_header_labels(ft, Variable = "Variable", W.value = "W Value", P.value = "p-Value")

# Format p-values as numeric with 2 decimal places
ft <- format(ft, j = "P.value", digits = 2)

# Apply zebra striping to the table
ft <- theme_zebra(ft, odd = "lightblue", even = "white")

# Autofit the column widths
ft <- autofit(ft)

# Create a Word document
doc <- read_docx()

# Add the table to the Word document
doc <- body_add_flextable(doc, value = ft)

# Save the Word document
print(doc, target = "C:/Users/tsamu/OneDrive/Documents/SU/Year 2/Thesis/Data/SUSY Results/Final.docx")

##Producing and exporting the same table for t-test (Table I)
# Create the data.frame
testy_t <- data.frame(
  Variable = c("Valence ESAbs", "Arousal ESAbs", "Pitch ESAbs", "Yaw ESAbs", "Roll ESAbs",
               "Valence ESnoabs", "Arousal ESnoabs", "Pitch ESnoabs", "Yaw ESnoabs", "Roll ESnoabs"),
  t_value = c(-0.1094770, -0.7399371, 1.1710904, 0.6077409, -0.8301939,
              -1.4353797, -0.6142069, 0.1623168, -1.0765971, 1.1773350),
  df = c(13.20572, 16.00890, 25.67278, 17.96805, 15.86475,
         18.47851, 12.58744, 23.86199, 25.79950, 27.73206),
  p_value = c(0.9144703, 0.4700511, 0.2523177, 0.5509641, 0.4187474,
              0.1678903, 0.5500191, 0.8724231, 0.2916250, 0.2490677)
)

# Create a flextable
ft <- flextable(testy_t)

# Set column names
set_header_labels(ft, Variable = "Variable", t_value = "t-value", df = "df", p_value = "p-value")

# Format p-values as numeric with 2 decimal places
ft <- format(ft, j = "p_value", digits = 2)

# Apply zebra striping to the table
ft <- theme_zebra(ft, odd = "lightblue", even = "white")

# Autofit the column widths
ft <- autofit(ft)

# Create a Word document
doc <- read_docx()

# Add the table to the Word document
doc <- body_add_flextable(doc, value = ft)

# Save the Word document
print(doc, target = "C:/Users/tsamu/OneDrive/Documents/SU/Year 2/Thesis/Data/SUSY Results/Dozen.docx")

# Create the data.frame
data <- data.frame(
  Variable = c("Valence ESAbs", "Arousal ESAbs", "Pitch ESAbs", "Yaw ESAbs", "Roll ESAbs",
               "Valence ESnoabs", "Arousal ESnoabs", "Pitch ESnoabs", "Yaw ESnoabs", "Roll ESnoabs"),
  t_value = c(-0.1094770, -0.7399371, 1.1710904, 0.6077409, -0.8301939,
              -1.4353797, -0.6142069, 0.1623168, -1.0765971, 1.1773350),
  df = c(13.20572, 16.00890, 25.67278, 17.96805, 15.86475,
         18.47851, 12.58744, 23.86199, 25.79950, 27.73206),
  p_value = c(0.9144703, 0.4700511, 0.2523177, 0.5509641, 0.4187474,
              0.1678903, 0.5500191, 0.8724231, 0.2916250, 0.2490677)
)

# Format t-values and degrees of freedom to two decimal places
data$t_value <- sprintf("%.2f", data$t_value)
data$df <- sprintf("%.2f", data$df)

# Create a flextable
ft <- flextable(data)

# Set column names
set_header_labels(ft, Variable = "Variable", t_value = "t-value", df = "df", p_value = "p-value")

# Format p-values as numeric with two decimal places
ft <- format(ft, j = "p_value", digits = 2)

# Apply zebra striping to the table
ft <- theme_zebra(ft, odd = "lightblue", even = "white")

# Autofit the column widths
ft <- autofit(ft)

# Create a Word document
doc <- read_docx()

# Add the table to the Word document
doc <- body_add_flextable(doc, value = ft)

# Save the Word document
print(doc, target = "C:/Users/tsamu/OneDrive/Documents/SU/Year 2/Thesis/Data/SUSY Results/Dozen.docx")

