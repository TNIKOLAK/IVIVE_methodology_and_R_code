################################################################################
#
# R Script to Generate a Word Document Report for the IVIVE Example
#
# Example 2: Compound X (Good Agreement Scenario)
#
# Required Packages: officer, flextable, rstudioapi
#
# MODIFICATION: This script now saves the output Word file to the same
#               directory as the script itself, with a timestamp in the filename.
#
################################################################################

# --- 1. SETUP: Install and load required packages ---

# List of required packages
required_packages <- c("officer", "flextable", "rstudioapi")

# Loop through the list, check if installed, and install if not
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg)
  }
}

library(officer)
library(flextable)
library(rstudioapi) # For detecting script path

# Clean the environment to ensure a fresh start
rm(list = ls())


# --- 2. CALCULATIONS: Perform all calculations for Compound X ---

#===============================================================================
#' PART 1: THE IN VITRO EXPERIMENT
#===============================================================================

# --- Given In Vitro Parameters ---
n_hepatocytes <- 0.5e6
vol_cell_single_um3 <- 3400
vol_medium_mL <- 1.0
mass_cell_vitro_nmol <- 0.02   # Changed for Compound X
mass_free_medium_nmol <- 0.20  # Changed for Compound X

# --- Step 1.1: Calculate In Vitro Concentrations ---
vol_cell_vitro_um3 <- n_hepatocytes * vol_cell_single_um3
vol_cell_vitro_L <- vol_cell_vitro_um3 / 1e15
vol_medium_L <- vol_medium_mL / 1000
c_free_medium_nM <- mass_free_medium_nmol / vol_medium_L
c_cell_vitro_nM <- mass_cell_vitro_nmol / vol_cell_vitro_L

# --- Step 1.2: Calculate the In Vitro Partition Coefficient (PC_in_vitro) ---
PC_in_vitro <- c_cell_vitro_nM / c_free_medium_nM

#===============================================================================
#' PART 2: THE IN SILICO PREDICTION
#===============================================================================

# --- Given In Vivo and Physiological Parameters ---
c_blood_obs_ug_L <- 17.5      # Changed for Compound X
c_tissue_obs_ug_kg <- 20.0    # Changed for Compound X
mm_compoundX_g_mol <- 350.0   # Changed for Compound X
Ka_compoundX_albumin_L_mol <- 1.0e5 # Changed for Compound X
c_protein_blood_uM <- 600
c_protein_ISF_uM <- 200
vol_fraction_blood <- 0.05
vol_fraction_ISF <- 0.15
vol_fraction_cell <- 0.80
vol_tissue_L <- 1.0

# --- Step 2.1: Calculate Free Concentration in Blood ---
c_blood_obs_mol_L <- (c_blood_obs_ug_L * 1e-6) / mm_compoundX_g_mol
c_blood_obs_nM <- c_blood_obs_mol_L * 1e9
c_protein_blood_mol_L <- c_protein_blood_uM * 1e-6
c_free_blood_nM <- c_blood_obs_nM / (1 + Ka_compoundX_albumin_L_mol * c_protein_blood_mol_L)

# --- Step 2.2: Apply IVIVE Hypothesis and Calculate Cellular Concentration ---
c_free_ISF_nM <- c_free_blood_nM
c_cell_nM <- c_free_ISF_nM * PC_in_vitro

# --- Step 2.3: Calculate Total Predicted Tissue Concentration ---
c_protein_ISF_mol_L <- c_protein_ISF_uM * 1e-6
c_total_ISF_nM <- c_free_ISF_nM * (1 + Ka_compoundX_albumin_L_mol * c_protein_ISF_mol_L)
mass_blood_nmol <- c_blood_obs_nM * (vol_tissue_L * vol_fraction_blood)
mass_ISF_nmol <- c_total_ISF_nM * (vol_tissue_L * vol_fraction_ISF)
mass_cell_nmol <- c_cell_nM * (vol_tissue_L * vol_fraction_cell)
mass_total_nmol <- mass_blood_nmol + mass_ISF_nmol + mass_cell_nmol
c_tissue_pred_nM <- mass_total_nmol / vol_tissue_L
c_tissue_pred_mol_kg <- c_tissue_pred_nM * 1e-9
c_tissue_pred_ug_kg <- (c_tissue_pred_mol_kg * mm_compoundX_g_mol) * 1e6


# --- 3. REPORT GENERATION: Create and populate the Word document ---

# Initialize a new Word document
doc <- read_docx()

# --- Add Title and Introduction ---
doc <- body_add_par(doc, "IVIVE Report: Compound X Distribution in Human Liver", style = "heading 1")
doc <- body_add_par(doc, "(Good Agreement Scenario)", style = "heading 2")
doc <- body_add_par(doc, "")

doc <- body_add_par(doc, "Objective", style = "heading 2")
doc <- body_add_par(doc, "A research team is investigating a new chemical, 'Compound X,' which is believed to distribute throughout the body primarily via passive diffusion with no significant active transport in the liver. The goal is to validate if a simple in vitro experiment can accurately predict the observed in vivo liver concentration, which would support the passive distribution hypothesis.")
doc <- body_add_par(doc, "")

# --- Part 1: In Vitro Experiment Section ---
doc <- body_add_par(doc, "Part 1: The In Vitro Experiment", style = "heading 2")
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "Step 1.1: Calculate In Vitro Concentrations", style = "heading 3")
doc <- body_add_par(doc, sprintf("Free Compound X Concentration in Medium: %.1f nM", c_free_medium_nM))
doc <- body_add_par(doc, sprintf("Total Compound X Concentration in Cells: %.1f nM", c_cell_vitro_nM))
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "Step 1.2: Calculate the In Vitro Partition Coefficient (PC_in_vitro)", style = "heading 3")
doc <- body_add_par(doc, sprintf("Calculated In Vitro Partition Coefficient (PC_in_vitro): %.1f", PC_in_vitro))
doc <- body_add_par(doc, "")

# --- Part 2: In Silico Prediction Section ---
doc <- body_add_par(doc, "Part 2: The In Silico Prediction", style = "heading 2")
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "Step 2.1: Calculate Free Compound X Concentration in Blood", style = "heading 3")
doc <- body_add_par(doc, sprintf("Observed Blood Concentration: %.1f nM", c_blood_obs_nM))
doc <- body_add_par(doc, sprintf("Calculated Free Blood Concentration: %.3f nM", c_free_blood_nM))
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "Step 2.2: Apply IVIVE Hypothesis and Calculate Cellular Concentration", style = "heading 3")
doc <- body_add_par(doc, sprintf("Predicted Cellular Concentration: %.1f nM", c_cell_nM))
doc <- body_add_par(doc, "")
doc <- body_add_par(doc, "Step 2.3: Calculate Total Predicted Tissue Concentration", style = "heading 3")
doc <- body_add_par(doc, sprintf("Final Predicted Tissue Concentration: %.1f ug/kg", c_tissue_pred_ug_kg), style = "Normal")
doc <- body_add_par(doc, "")

# --- Part 3: Comparison and Conclusion Section ---
doc <- body_add_par(doc, "Part 3: Comparison, Analysis, and Conclusion", style = "heading 2")
doc <- body_add_par(doc, "")

# Create a data frame for the results table
results_df <- data.frame(
  Parameter = c("Liver Concentration (ug/kg)", "Tissue:Blood Ratio"),
  Observed = c(c_tissue_obs_ug_kg, c_tissue_obs_ug_kg / c_blood_obs_ug_L),
  Predicted = c(c_tissue_pred_ug_kg, c_tissue_pred_ug_kg / c_blood_obs_ug_L)
)

# Create a formatted flextable
ft <- flextable(results_df)
ft <- set_header_labels(ft, Parameter = "Parameter", Observed = "Observed Value", Predicted = "Predicted Value")
ft <- colformat_double(ft, j = ~ Observed + Predicted, digits = 2)
ft <- autofit(ft)
ft <- theme_booktabs(ft)

# Add the table to the document
doc <- body_add_par(doc, "Comparison of Results", style = "heading 3")
doc <- body_add_flextable(doc, value = ft)
doc <- body_add_par(doc, "")

# Generate the dynamic conclusion text
tolerance <- 0.2
ratio_pred_obs <- c_tissue_pred_ug_kg / c_tissue_obs_ug_kg

conclusion_text <- if (abs(ratio_pred_obs - 1) <= tolerance) {
  sprintf(
    "The predicted liver concentration (%.1f ug/kg) is in good agreement with the observed value (%.1f ug/kg), falling within the defined %.0f%% tolerance. Conclusion: The hypothesis that a simple in vitro experiment can predict the in vivo liver concentration of Compound X is STRONGLY SUPPORTED by this example. The close match between the predicted and observed values indicates that the model, which only accounts for passive partitioning and protein binding, is sufficient to describe the distribution of Compound X in the liver. This result validates the initial assumption that Compound X is not a significant substrate for active uptake or efflux transporters in this tissue.",
    c_tissue_pred_ug_kg, c_tissue_obs_ug_kg, tolerance * 100
  )
} else if (ratio_pred_obs > 1) {
  sprintf(
    "The model significantly OVER-PREDICTED the liver concentration. The predicted value is %.1f ug/kg, which is %.1f times HIGHER than the observed value of %.1f ug/kg. Conclusion: The hypothesis that a simple in vitro experiment can predict the in vivo liver concentration of Compound X is NOT SUPPORTED by this example. The discrepancy suggests that other biological processes, not included in the model, are limiting accumulation in vivo. The most likely missing factor is the presence of active EFFLUX transporters on the hepatocyte membrane.",
    c_tissue_pred_ug_kg, ratio_pred_obs, c_tissue_obs_ug_kg
  )
} else {
  sprintf(
    "The model significantly UNDER-PREDICTED the liver concentration. The predicted value is %.1f ug/kg, which is only %.2f times the observed value of %.1f ug/kg. Conclusion: The hypothesis that a simple in vitro experiment can predict the in vivo liver concentration of Compound X is NOT SUPPORTED by this example. The discrepancy suggests that other biological processes, not included in the model, are enhancing accumulation in vivo. The most likely missing factor is the presence of active UPTAKE transporters on the hepatocyte membrane.",
    c_tissue_pred_ug_kg, ratio_pred_obs, c_tissue_obs_ug_kg
  )
}

# Add the conclusion to the document
doc <- body_add_par(doc, "Analysis and Conclusion", style = "heading 3")
doc <- body_add_par(doc, conclusion_text)

# --- 4. SAVE THE DOCUMENT ---

# --- NEW DYNAMIC FILE PATH LOGIC ---

# Get the directory of the current script. Fallback to working directory if not in RStudio.
output_folder <- tryCatch({
  # This works reliably when the script is run via the 'Source' button in RStudio
  dirname(rstudioapi::getSourceEditorContext()$path)
}, error = function(e) {
  # If not in RStudio or if an error occurs, use the current working directory
  message("Not running in RStudio or path not available. Saving to working directory.")
  getwd()
})

# Generate a timestamp string in the format YYYYMMDDHHMM
timestamp <- format(Sys.time(), "%Y%m%d%H%M")

# Create the dynamic filename
dynamic_filename <- paste0("IVIVE_Report_CompoundX_", timestamp, ".docx")

# Combine the folder and filename into a full, platform-independent path
file_path <- file.path(output_folder, dynamic_filename)

# Save the document to the determined path
print(doc, target = file_path)

# Notify the user with the full path
cat(sprintf("\nReport has been generated successfully and saved to:\n%s\n", file_path))
cat("------------------------------------------------------------------\n")
cat("End of Script\n")
cat("------------------------------------------------------------------\n")