# Load necessary libraries
library(dplyr)
library(tidyr)
library(openxlsx)

# Read the data file
data <- read.delim("DIA-NN-PROTEINS-LIBRARY MODE-report.pg_matrix.tsv")

# Function to determine if a protein is human or E.coli based on Protein.Ids
is_human <- function(protein_id) {
  !grepl("Ecoli", protein_id, ignore.case = TRUE)
}

is_ecoli <- function(protein_id) {
  grepl("Ecoli", protein_id, ignore.case = TRUE)
}

# Add columns to classify proteins
data <- data %>%
  mutate(Human = ifelse(is_human(Protein.Ids), 1, 0),
         Ecoli = ifelse(is_ecoli(Protein.Ids), 1, 0))

# Get sample columns
sample_columns <- names(data)[grepl("wiff", names(data))]

# Initialize result data frame
results <- data.frame(Sample = sample_columns)

# Calculate counts and missing data for each sample
for (sample in sample_columns) {
  results[results$Sample == sample, "Human_Protein_Count"] <- sum(data$Human & !is.na(data[[sample]]))
  results[results$Sample == sample, "Ecoli_Protein_Count"] <- sum(data$Ecoli & !is.na(data[[sample]]))
  
  results[results$Sample == sample, "Human_Missing_Data"] <- sum(data$Human & is.na(data[[sample]]))
  results[results$Sample == sample, "Ecoli_Missing_Data"] <- sum(data$Ecoli & is.na(data[[sample]]))
  
  total_human <- sum(data$Human)
  total_ecoli <- sum(data$Ecoli)
  
  results[results$Sample == sample, "Human_Missing_Percentage"] <- 100 * results[results$Sample == sample, "Human_Missing_Data"] / total_human
  results[results$Sample == sample, "Ecoli_Missing_Percentage"] <- 100 * results[results$Sample == sample, "Ecoli_Missing_Data"] / total_ecoli
}

# Calculate CV and percentages for protein groups
data_long <- data %>%
  pivot_longer(cols = all_of(sample_columns), names_to = "Sample", values_to = "Intensity")

# Classify samples into groups
sample_groups <- rep(1:3, each = 5)
names(sample_groups) <- sample_columns

# Summarize data by Protein.Ids and group
data_long <- data_long %>%
  mutate(Group = sample_groups[Sample]) %>%
  group_by(Protein.Ids, Group) %>%
  summarize(Mean_Intensity = mean(Intensity, na.rm = TRUE),
            SD_Intensity = sd(Intensity, na.rm = TRUE),
            Human = first(Human),
            Ecoli = first(Ecoli),
            .groups = 'drop') %>%
  mutate(CV = 100 * SD_Intensity / Mean_Intensity)

# Initialize group result data frame
group_results <- data.frame(Group = c(1:3, "Overall"))

# Calculate CV percentages for each group and overall
for (group in 1:3) {
  group_data <- data_long %>% filter(Group == group)
  
  group_results[group, "Ecoli_CV_Under_30_Count"] <- sum(group_data$Ecoli == 1 & group_data$CV < 30, na.rm = TRUE)
  group_results[group, "Human_CV_Under_30_Count"] <- sum(group_data$Human == 1 & group_data$CV < 30, na.rm = TRUE)
  
  group_results[group, "Ecoli_CV_Under_20_Count"] <- sum(group_data$Ecoli == 1 & group_data$CV < 20, na.rm = TRUE)
  group_results[group, "Human_CV_Under_20_Count"] <- sum(group_data$Human == 1 & group_data$CV < 20, na.rm = TRUE)
}

# Calculate overall CV percentages
overall_data <- data_long
group_results[4, "Ecoli_CV_Under_30_Count"] <- sum(overall_data$Ecoli == 1 & overall_data$CV < 30, na.rm = TRUE)
group_results[4, "Human_CV_Under_30_Count"] <- sum(overall_data$Human == 1 & overall_data$CV < 30, na.rm = TRUE)

group_results[4, "Ecoli_CV_Under_20_Count"] <- sum(overall_data$Ecoli == 1 & overall_data$CV < 20, na.rm = TRUE)
group_results[4, "Human_CV_Under_20_Count"] <- sum(overall_data$Human == 1 & overall_data$CV < 20, na.rm = TRUE)

# Create and save the workbook
output_directory <- getwd()
output_file <- file.path(output_directory, "protein_analysis_results.xlsx")
wb <- createWorkbook()
addWorksheet(wb, "Sample_Results")
writeData(wb, "Sample_Results", results)
addWorksheet(wb, "Group_CV_Results")
writeData(wb, "Group_CV_Results", group_results)
addWorksheet(wb, "Data_Long")
writeData(wb, "Data_Long", data_long)

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)

# Output file path
output_file