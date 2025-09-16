# IVIVE_methodology_and_R_code
•	Overview
This repository contains resources for performing In Vitro to In Vivo Extrapolation (IVIVE) to predict the steady-state concentration of a chemical substance (e.g., Compound X) in a target tissue (e.g., liver) based on in vitro experiments. The methodology uses a partition coefficient derived from cryopreserved cells (e.g., human hepatocytes) and extrapolates it to in vivo predictions, accounting for passive partitioning and protein binding.
The IVIVE approach validates models by comparing predicted tissue concentrations to observed in vivo values, helping to identify if active transport mechanisms (e.g., uptake or efflux transporters) are involved.
This is demonstrated through an example scenario for "Compound X" (Good Agreement Scenario), where the model shows close alignment between predictions and observations, supporting passive distribution.
•	Files Included
•	IVIVE_Methodology_20250916.pdf: A detailed PDF document outlining the IVIVE methodology. It includes: 
o	Objective and overview.
o	Step-by-step descriptions of the in vitro experiment, partition coefficient derivation, in silico prediction, comparison/analysis, and conclusion.
o	Notation tables for parameters (known/measured values) and variables (calculated values), with corresponding R code symbols.
o	Equations for all calculations.
o	Date: September 16, 2025.
•	IVIVE_R_Code_v1_2020916.R: An R script that implements the IVIVE methodology for Compound X. 
o	Performs all calculations from the methodology.
o	Generates a Word document report (IVIVE_Report_CompoundX_YYYYMMDDHHMM.docx) with results, tables, and a dynamic conclusion based on prediction-observation agreement.
o	Saves the output Word file to the same directory as the script, with a timestamp.
o	Example-specific parameters are hardcoded for demonstration (e.g., in vitro and in vivo values for Compound X).
•	Requirements
•	R Environment: R version 3.0 or higher (tested with recent versions).
•	Required Packages: 
o	officer: For creating and manipulating Word documents.
o	flextable: For formatting tables in the report.
o	rstudioapi: For detecting the script path (optional; falls back to working directory if not in RStudio).
The script automatically checks and installs these packages if missing.
•	Software: Microsoft Word or a compatible viewer to open the generated .docx report. Running in RStudio is recommended for path detection.
•	How to Use
1.	Clone the Repository: 
git clone https://github.com/your-username/IVIVE_methodology_and_R_code.git
cd IVIVE_methodology_and_R_code

2.	Review the Methodology: 
o	Open IVIVE_Methodology_20250916.pdf to understand the scientific background, equations, and notation.
3.	Run the R Script: 
o	Open IVIVE_R_Code_v1_2020916.R in RStudio or your R environment.
o	Source the script (e.g., via RStudio's "Source" button or source("IVIVE_R_Code_v1_2020916.R") in the console).
o	The script will: 
	Perform IVIVE calculations for Compound X.
	Generate a Word report in the same directory, e.g., IVIVE_Report_CompoundX_202509161200.docx.
	Output a console message with the full file path.
4.	Customize the Script (Optional): 
o	Edit the hardcoded parameters in Section 2 (e.g., n_hepatocytes, c_blood_obs_ug_L, etc.) to test different compounds or scenarios.
o	Adjust the tolerance threshold in the conclusion generation (default: 20% for "good agreement").
o	For other scenarios (e.g., over-prediction or under-prediction), modify values to simulate efflux/uptake effects.
•	Example Output
•	Calculations: In vitro partition coefficient, free concentrations, predicted tissue concentration, etc.
•	Report Structure: 
o	Title and objective.
o	Part 1: In vitro results.
o	Part 2: In silico predictions.
o	Part 3: Comparison table (observed vs. predicted), analysis, and conclusion (e.g., "STRONGLY SUPPORTED" for good agreement).
•	Sample Conclusion (for default Compound X values): The model validates passive distribution, with no significant active transport.
•	Notes
•	The R script is stateful and cleans the environment at the start for reproducibility.
•	The PDF references the R script as "IVIVE R Code v1 2020916.R" in its notation section, ensuring alignment between documentation and code.
•	This is an educational/demonstrative tool; for real-world applications, validate with experimental data and consult domain experts.
•	Contributions: Feel free to open issues or pull requests for improvements, such as adding more examples or visualizations.
•	License
This repository is licensed under the MIT License. See LICENSE for details.
•	Contact
For questions, contact [nikolako@mail.ntua.gr] or open an issue on GitHub.

