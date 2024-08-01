#Author: Ze Liu; zeliu_cugb@163.com
###################################################################
if(!require("readxl")){installed.packages("readxl")}
if(!require("dplyr")){installed.packages("dplyrdplyr")}
if(!require("writexl")){installed.packages("writexl")}

library(readxl)
library(dplyr)
library(writexl)

path<-getwd()
file_path <- "C:/Dropbox/Zr isotope/Zr isotope template.xls" 
data.raw <- read_excel(file_path, sheet = "IoliteDataTableExport")

data <- data.raw %>%
  filter(grepl("Output", SelectionLabel))

selected_columns <- data %>%
  select(Comments, StdCorr_Zr94_90, StdCorr_Zr94_90_Int2SE,
         StdCorr_Zr94_91, StdCorr_Zr94_91_Int2SE,
         StdCorr_Zr96_90, StdCorr_Zr96_90_Int2SE)

#change the variable name
result <- selected_columns %>%
  rename(Analyses = Comments) %>%
  mutate(D94_90Zr = NA,
         `2SE_94_90Zr` = NA,
         D94_91Zr = NA,
         `2SE_94_91Zr` = NA,
         D96_90Zr = NA,
         `2SE_96_90Zr` = NA)

# find all row whose Analyses contain “GJ-1”
gj1_rows <- which(grepl("GJ-1", result$Analyses))

#calculation for the row whose Analyses do not contain GJ-1
for (i in 1:nrow(result)) {
  if (!grepl("GJ-1", result$Analyses[i])) {
    upper_gj1 <- tail(gj1_rows[gj1_rows < i], 1)
    lower_gj1 <- head(gj1_rows[gj1_rows > i], 1)
    
    if (length(upper_gj1) > 0 & length(lower_gj1) > 0) {
      #calculate D94_90Zr
      D94_90Zr_val <- ((result$StdCorr_Zr94_90[i] * 2) /
                         (result$StdCorr_Zr94_90[upper_gj1] + result$StdCorr_Zr94_90[lower_gj1]) - 1) * 1000
      result$D94_90Zr[i] <- D94_90Zr_val
      
      #calculate 2SE_94_90Zr
      SE_val_94_90 <- 2 * (1000000 * ((result$StdCorr_Zr94_90_Int2SE[i] / 2 / result$StdCorr_Zr94_90[i])^2 +
                                (result$StdCorr_Zr94_90_Int2SE[upper_gj1] / 2 / result$StdCorr_Zr94_90[upper_gj1])^2 +
                                (result$StdCorr_Zr94_90_Int2SE[lower_gj1] / 2 / result$StdCorr_Zr94_90[lower_gj1])^2))^0.5
      result$`2SE_94_90Zr`[i] <- SE_val_94_90
      
       #calculate D94_91Zr
      D94_91Zr_val <- ((result$StdCorr_Zr94_91[i] * 2) /
                         (result$StdCorr_Zr94_91[upper_gj1] + result$StdCorr_Zr94_91[lower_gj1]) - 1) * 1000
      result$D94_91Zr[i] <- D94_91Zr_val
      
      #calculate 2SE_94_91Zr
      SE_val_94_91 <- 2 * (1000000 * ((result$StdCorr_Zr94_91_Int2SE[i] / 2 / result$StdCorr_Zr94_91[i])^2 +
                                (result$StdCorr_Zr94_91_Int2SE[upper_gj1] / 2 / result$StdCorr_Zr94_91[upper_gj1])^2 +
                                (result$StdCorr_Zr94_91_Int2SE[lower_gj1] / 2 / result$StdCorr_Zr94_91[lower_gj1])^2))^0.5
      result$`2SE_94_91Zr`[i] <- SE_val_94_91
      
      #calculate D96_90Zr
      D96_90Zr_val <- ((result$StdCorr_Zr96_90[i] * 2) /
                         (result$StdCorr_Zr96_90[upper_gj1] + result$StdCorr_Zr96_90[lower_gj1]) - 1) * 1000
      result$D96_90Zr[i] <- D96_90Zr_val
      
      #calculate 2SE_96_90Zr
      SE_val_96_90 <- 2 * (1000000 * ((result$StdCorr_Zr96_90_Int2SE[i] / 2 / result$StdCorr_Zr96_90[i])^2 +
                                (result$StdCorr_Zr96_90_Int2SE[upper_gj1] / 2 / result$StdCorr_Zr96_90[upper_gj1])^2 +
                                (result$StdCorr_Zr96_90_Int2SE[lower_gj1] / 2 / result$StdCorr_Zr96_90[lower_gj1])^2))^0.5
      result$`2SE_96_90Zr`[i] <- SE_val_96_90
   
    }
  }
}
#RM 8299 is from Tissot, F.L.H., Ibañez-Mejia, M., Rabb, S.A., Kraft, R.A., Vocke, R.D., Fehr, M.A., Schönbächler, M., Tang, H., Young, E.D., 2023. A community-led calibration of the Zr isotope reference materials: NIST candidate RM 8299 and SRM 3169. J. Anal. At. Spectrom. 38, 2087-2104. 10.1039/D3JA00167A.

#function for getting input
get_user_input <- function(prompt) {
  cat(prompt, "\n")
  value <- NA
  while(is.na(value)) {
    input <- scan(what = numeric(), n = 1, quiet = TRUE)
    if (length(input) == 0) {
      cat("Invalid input. Please re-enter a numeric value：\n")
    } else {
      value <- input[1]
    }
  }
  cat(sprintf("Your input is: %f\n", value))
  return(value)
}

#main code
main_program <- function(parameter.94_90Zr.standard, parameter.94_90Zr.standard.2SE) {
  tryCatch({
    cat("Continue...\n")
    
    result$DeltaZr94_90_RM8299 = ((result$D94_90Zr/1000+1)*(parameter.94_90Zr.standard/1000+1)-1)*1000
    result$Delta2SE_Zr94_90_RM8299 = ((result$`2SE_94_91Zr`/2)^2+(parameter.94_90Zr.standard.2SE/2)^2)^0.5*2
    
    result <- result %>%
      select(-grep("std", names(result), ignore.case = TRUE))
    
    delete_row <- which(grepl("GJ-1", result$Analyses))
    result <- result[-delete_row,]
    
    cat("The data will be written to an Excel file. Press Enter to continue...\n")
    readline()
    
    write_xlsx(result, paste0(path, "/Zr isotope data.xlsx"))
    
    cat("The Excel file has been written. The program will now exit.\n")
  }, error = function(e) {
    cat("An error occurred during program execution:", conditionMessage(e), "\n")
  }, warning = function(w) {
    cat("Warning：", conditionMessage(w), "\n")
  }, finally = {
    cat("The program execution is complete\n")
  })
}

#Execute the program
tryCatch({
  parameter.94_90Zr.standard <- get_user_input("Please input the parameter.94_90Zr.standard：#Recomamnded value is -0.064; GJ-1 Tissort 2023 JAAS")
  parameter.94_90Zr.standard.2SE <- get_user_input("Please input the parameter.94_90Zr.standard.2SE：#Recomamnded value is 0.032; GJ-1 Tissort 2023 JAAS")
  
  main_program(parameter.94_90Zr.standard, parameter.94_90Zr.standard.2SE)
}, error = function(e) {
  cat("An error occurred during program execution:", conditionMessage(e), "\n")
}, warning = function(w) {
  cat("Warning：", conditionMessage(w), "\n")
}, finally = {
  cat("The program execution is complete...\n")
})

