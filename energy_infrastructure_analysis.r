# Statistical Analysis of Transmission and Substation Infrastructure Data
# Author: [Your Name]
# Date: [Current Date]
# Description: Analysis of infrastructure project costs, durations, and overruns

# Load required libraries
library(dplyr)
library(ggplot2)
library(flextable)
library(officer)
library(RColorBrewer)
library(scales)
library(openxlsx)

#------------------------------------------------------------------------------
# Data Preprocessing
#------------------------------------------------------------------------------

# Function to prepare dataset with standardized column names
prepare_dataset <- function(data, type) {
  data$CountryGroup <- data$Region
  data$TotalCost <- data$`Total Outturn Project Expenditure (millions, AUD, 2024, PPP)`
  data$ConstructionCost <- data$`Total Outturn Construction Cost (millions, AUD, 2024, PPP)`
  data$Cost_Overrun <- data$`Nominal base cost overrun (%)`
  data$Schedule_Overrun <- data$`Schedule overrun (%)`
  
  if (type == "Transmission") {
    data$Project_Cost_KV <- data$`Project cost per km per kV`
    data$Project_Cost_MVA <- data$`Project cost per km per MVA`
    data$Construction_Cost_KV <- data$`Construction cost per km per kV`
    data$Construction_Cost_MVA <- data$`Construction cost per km per MVA`
    data$Size <- data$`Route Length (km)`
    data$Duration <- data$`Actual duration (years) (date of completion - planned start date)` * 12
  } else if (type == "Substation") {
    data$Project_Cost_KV <- data$`Project cost per kV`
    data$Project_Cost_MVA <- data$`Project cost per MVA`
    data$Construction_Cost_KV <- data$`Construction cost per kV`
    data$Construction_Cost_MVA <- data$`Construction cost per MVA`
    data$Size <- data$`Station Size (m2)`/ data$`No. of stations`
    data$Duration <- data$`Actual construction duration (years) (date of completion - planned start date)` * 12 / data$`No. of stations`
    data$Project_Cost_Station <- data$`Project cost per station`
  }
  
  return(data)
}

# Prepare datasets
Transmission <- prepare_dataset(Transmission, "Transmission")
Substation <- prepare_dataset(Substation, "Substation")

#------------------------------------------------------------------------------
# Summary Statistics Functions
#------------------------------------------------------------------------------

# Function to generate summary statistics
generate_summary_stats <- function(data, variable, digits = 2) {
  data %>%
    group_by(CountryGroup) %>%
    summarise(
      Mean = mean(.data[[variable]], na.rm = TRUE),
      Std_Dev = sd(.data[[variable]], na.rm = TRUE),
      Min = min(.data[[variable]], na.rm = TRUE),
      Max = max(.data[[variable]], na.rm = TRUE),
      Median = median(.data[[variable]], na.rm = TRUE),
      P25 = quantile(.data[[variable]], 0.25, na.rm = TRUE),
      P75 = quantile(.data[[variable]], 0.75, na.rm = TRUE),
      Count = sum(!is.na(.data[[variable]]))
    ) %>%
    ungroup() %>%
    mutate(across(where(is.numeric), ~round(., digits = digits)))
}

# Function to add summary to Word document
add_summary_to_doc <- function(data, variable, title, digits, doc) {
  summary_table <- generate_summary_stats(data, variable, digits)
  overall_stats <- generate_summary_stats(data %>% mutate(CountryGroup = "Overall"), variable, digits)
  
  combined_table <- bind_rows(summary_table, overall_stats)
  colnames(combined_table) <- c("Location", "Mean", "Std. Dev.", "Min", "Max", "Median", "Q25", "Q75", "Count")
  
  doc <- doc %>%
    body_add_par("") %>%
    body_add_par(title) %>%
    body_add_flextable(value = flextable(combined_table))
  
  return(doc)
}

#------------------------------------------------------------------------------
# Visualization Functions
#------------------------------------------------------------------------------

# Function to create and save histogram
create_histogram <- function(data, dataset_name, output_path) {
  num_groups <- length(unique(data$Region))
  color_palette <- brewer.pal(n = num_groups, name = "Blues")
  
  plot <- ggplot(data = data) +
    geom_bar(aes(x = Region, fill = Region), position = "dodge", alpha = 0.7) +
    labs(x = NULL, 
         y = "Number of projects in data sample", 
         title = paste("Overview of projects per region -", dataset_name)) +
    scale_y_continuous(breaks = scales::pretty_breaks()) +
    scale_fill_manual(values = color_palette) +
    theme_minimal() +
    theme(
      legend.title = element_blank(),
      panel.grid.minor = element_blank(),
      axis.line = element_line(colour = "black")
    )
  
  ggsave(
    filename = file.path(output_path, paste0("Histogram_", dataset_name, ".png")),
    plot = plot,
    width = 7.5,
    height = 3,
    dpi = 300
  )
}

# Function to create and save density plot
create_density_plot <- function(data, variable, title, x_label, output_path, dataset_name) {
  valid_obs_count <- sum(!is.na(data[[variable]]))
  
  plot <- ggplot(data = data) +
    geom_density(
      aes(x = !!sym(variable), y = ..scaled..), 
      alpha = 0.7, 
      color = "#7BB6D9", 
      fill = "#9BC8E0"
    ) +
    scale_x_continuous(labels = label_number(scale = 1, big.mark = ",")) +
    labs(
      x = x_label,
      y = "Normalized Density",
      title = title,
      subtitle = paste0("Number of valid observations = ", valid_obs_count)
    ) +
    theme_minimal() +
    theme(
      legend.title = element_blank(),
      panel.grid.minor = element_blank(),
      axis.line = element_line(colour = "black")
    )
  
  ggsave(
    filename = file.path(output_path, 
                        paste0("Density_", dataset_name, "_", gsub(" ", "_", variable), ".png")),
    plot = plot,
    width = 7.5,
    height = 3,
    dpi = 300
  )
}

#------------------------------------------------------------------------------
# Statistical Analysis Functions
#------------------------------------------------------------------------------

# Function to perform statistical tests
perform_statistical_tests <- function(data, dependent_var, grouping_var) {
  # Kruskal-Wallis test
  kw_test <- kruskal.test(formula(paste(dependent_var, "~", grouping_var)), data = data)
  
  # Pairwise Wilcoxon test
  pw_test <- pairwise.wilcox.test(
    data[[dependent_var]], 
    data[[grouping_var]], 
    p.adjust.method = "none"
  )
  
  return(list(kruskal_wallis = kw_test, pairwise_wilcox = pw_test))
}

#------------------------------------------------------------------------------
# Reference Class Analysis
#------------------------------------------------------------------------------

# Function to calculate quantile distributions
calculate_quantiles <- function(data, variable) {
  probs <- seq(0.05, 0.95, by = 0.05)
  
  results <- data.frame(
    Quantile = c(probs, "Average", "Number of observations")
  )
  
  quant_values <- quantile(data[[variable]], probs = probs, na.rm = TRUE)
  avg_value <- mean(data[[variable]], na.rm = TRUE)
  n_obs <- sum(!is.na(data[[variable]]))
  
  results$Value <- c(quant_values, avg_value, n_obs)
  
  return(results)
}

#------------------------------------------------------------------------------
# Main Execution
#------------------------------------------------------------------------------

# Set working directory and output path
# Note: Update these paths according to your environment
setwd("path/to/your/working/directory")
output_path <- "path/to/your/output/directory"

# Generate summaries and visualizations
datasets <- list(Transmission = Transmission, Substation = Substation)

# Create Word document with summaries
doc <- read_docx()
for (dataset_name in names(datasets)) {
  data <- datasets[[dataset_name]]
  variables <- names(data)[grep("^(Project|Construction|Cost|Schedule|Total|Duration|Size)", names(data))]
  
  for (variable in variables) {
    doc <- add_summary_to_doc(data, variable, paste(dataset_name, variable), 2, doc)
  }
}

# Save Word document
print(doc, target = file.path(output_path, "descriptive_statistics.docx"))

# Create visualizations
for (dataset_name in names(datasets)) {
  data <- datasets[[dataset_name]]
  create_histogram(data, dataset_name, output_path)
  
  # Add density plots and other visualizations as needed
}

# Perform statistical analysis
# Add your specific analysis calls here

# Save results
# Add your specific save commands here
