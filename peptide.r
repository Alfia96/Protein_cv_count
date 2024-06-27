# Load necessary libraries
library(dplyr)
library(tidyr)
library(openxlsx)

# Read the peptides data file
peptides_data <- read.delim("DIA-NN-PEPTIDES-LIBRARY_MODE_report.pr_matrix.tsv")

# Function to determine if a peptide belongs to human or E.coli based on Protein.Ids
is_human <- function(protein_id) {
  !grepl("Ecoli", protein_id, ignore.case = TRUE)
}

is_ecoli <- function(protein_id) {
  grepl("Ecoli", protein_id, ignore.case = TRUE)
}

# Add columns to classify peptides
peptides_data <- peptides_data %>%
  mutate(Human = ifelse(is_human(Protein.Ids), 1, 0),
         Ecoli = ifelse(is_ecoli(Protein.Ids), 1, 0))

# Get sample columns
sample_columns <- names(peptides_data)[grepl("wiff", names(peptides_data))]

# Initialize result data frames
peptide_counts <- data.frame(Sample = sample_columns)
peptides_per_protein <- data.frame()

# Calculate counts for each sample
for (sample in sample_columns) {
  peptide_counts[peptide_counts$Sample == sample, "Human_Peptide_Count"] <- sum(peptides_data$Human & !is.na(peptides_data[[sample]]))
  peptide_counts[peptide_counts$Sample == sample, "Ecoli_Peptide_Count"] <- sum(peptides_data$Ecoli & !is.na(peptides_data[[sample]]))
}

# Calculate the number of peptides per protein for both human and E.coli
peptides_per_protein <- peptides_data %>%
  group_by(Protein.Ids) %>%
  summarize(Human_Peptide_Count = sum(Human & !is.na(rowSums(across(all_of(sample_columns)))), na.rm = TRUE),
            Ecoli_Peptide_Count = sum(Ecoli & !is.na(rowSums(across(all_of(sample_columns)))), na.rm = TRUE),
            Total_Peptide_Count = sum(!is.na(rowSums(across(all_of(sample_columns)))), na.rm = TRUE))

# Define output path in the current working directory
output_directory <- getwd()
output_file <- file.path(output_directory, "peptide_analysis_results.xlsx")

# Write results to an Excel file
wb <- createWorkbook()

# Add sheets and write data
addWorksheet(wb, "Peptide_Counts_Per_Sample")
writeData(wb, "Peptide_Counts_Per_Sample", peptide_counts)

addWorksheet(wb, "Peptides_Per_Protein")
writeData(wb, "Peptides_Per_Protein", peptides_per_protein)

# Save the workbook
saveWorkbook(wb, output_file, overwrite = TRUE)

# Output file path
output_file

