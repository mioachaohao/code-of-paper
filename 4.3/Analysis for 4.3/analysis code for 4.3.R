library(plyr)
# Import libraries and load data
library(tidyverse)
library(igraph)
library(statnet)
library(sysfonts)
library(showtext)
library(Cairo)
library(readxl)
library(circlize)
library(RColorBrewer)

environment_file_path="D:/code/4.3/Analysis for 4.3"
setwd(environment_file_path)
load("Data for the Upstream, Midstream, and Downstream of the Photovoltaic Industry.RData")

# Read the country code data from the Excel file
country_code <- read_excel("country_code.xlsx", sheet = "English")
country_code <- select(country_code, i = country_code, name = iso_3digit_alpha)

# Specify the Excel file path
excel_file_path <- "wb_nodelist.xlsx"

# Get the sheet names from the Excel file
sheet_names <- excel_sheets(excel_file_path)

# Create an empty list to store the data from each sheet
all_tables <- list()

# Loop through each sheet, read its contents, and store it in a variable named after the sheet name
for (sheet_name in sheet_names) {
  table_data <- read_excel(excel_file_path, sheet = sheet_name)
  assign(sheet_name, table_data)
  all_tables[[sheet_name]] <- table_data
}

# Process the data and write to CSV files
for (a in c("_upstream", "_midstream", "_downstream")) {
  for (i in 1995:2020) {
    data <- get(paste("data_", i, a, sep = ""))
    data <- data %>%
      left_join(country_code, by = "i") %>%
      left_join(country_code, by = c("j" = "i")) %>%
      select(export = name.x, import = name.y, weight = v)
    data <- na.omit(data)
    mat <- matrix(0, length(union(unlist(data[1]), unlist(data[2]))), length(union(unlist(data[1]), unlist(data[2]))))
    rownames(mat) <- union(unlist(data[1]), unlist(data[2]))
    colnames(mat) <- union(unlist(data[1]), unlist(data[2]))
    mat[as.matrix(data[, 1:2])] <- data[, 3]
    assign(paste("data_", i, a, "_matrix", sep = ""), mat)
    write.table(get(paste("data_", i, a, "_matrix", sep = "")), paste("Intermediate Data/data_", i, a, "_matrix", ".csv", sep = ""), row.names = TRUE, col.names = TRUE, sep = ",")
    rm(data)
  }
}

# Define a core function for random attack
random_attack <- function(mat, k, a = 0.9, b = 0.6) {
  round_all <- list()
  all_import <- apply(mat, 2, sum)
  mat <- rbind(mat, import = all_import)
  mat <- rbind(mat, all_import = all_import)
  len <- length(colnames(mat))
  round <- k
  round_all <- append(round_all, list(round))
  while (length(round) != 0) {
    mat <- mat[, -which(colnames(mat) %in% round)]
    if (length(colnames(mat)) == 0) {
      break
    }
    mat[round, ] <- (1 - a) * mat[round, ]
    mat["import", ] <- apply(mat[1:(length(rownames(mat)) - 2), ], 2, sum)
    round <- c()
    for (i in colnames(mat)) {
      if (mat["import", i] < (1 - b) * mat["all_import", i]) {
        round <- append(round, i)
      }
    }
    round_all <- append(round_all, list(round))
  }
  return(round_all)
}

# Calculate risk propagation rates and save the results to CSV
output_mean <- data.frame()
output <- data.frame()

# Define parameters a and b for the random attack
for (t in 1995:2019) {
  mat_1 <- get(paste("data_", t, "_upstream_matrix", sep = ""))
  mat_2 <- get(paste("data_", t, "_midstream_matrix", sep = ""))
  mat_3 <- get(paste("data_", t, "_downstream_matrix", sep = ""))
  output <- data.frame()
  for (k in row.names(mat_1)) {
    upstream_infected <- random_attack(mat_1, k)
    midstream_source <- intersect(unlist(upstream_infected), row.names(mat_2))
    midstream_infected <- random_attack(mat_2, midstream_source)
    downstream_source <- intersect(unlist(midstream_infected), row.names(mat_3))
    downstream_infected <- random_attack(mat_3, downstream_source)
    output <- rbind(output, c(k, length(unlist(upstream_infected)), length(unlist(midstream_infected)), length(unlist(downstream_infected))))
  }
  colnames(output) <- c("name", "upstream", "midstream", "downstream")
  output$upstream <- as.numeric(output$upstream)
  output$midstream <- as.numeric(output$midstream)
  output$downstream <- as.numeric(output$downstream)
  output$all <- output$upstream + output$midstream + output$downstream
  output$upstream_ratio <- output$upstream / length(unlist(colnames(mat_1)))
  output$midstream_ratio <- output$midstream / length(unlist(colnames(mat_2)))
  output$downstream_ratio <- output$downstream / length(unlist(colnames(mat_3)))
  output <- arrange(output, -all)
  assign(paste("output_", t, sep = ""), output)
  output_mean <- rbind(output_mean, c(t, mean(as.numeric(output$upstream)), mean(as.numeric(output$midstream)), mean(as.numeric(output$downstream)), mean(as.numeric(output$upstream_ratio)), mean(as.numeric(output$midstream_ratio)), mean(as.numeric(output$downstream_ratio))))
  colnames(output_mean) <- c("time", "upstream_country", "midstream_country", "downstream_country", "upstream_rate", "midstream_rate", "downstream_rate")
  print(t)
}

# Save the average risk propagation rates to a CSV file
write.csv(output_mean, "results/Average Risk Propagation Rate.csv")

library(openxlsx)

# Create a new workbook
wb <- createWorkbook()

# Initialize data frames
years <- c(1995:2019)
upstream_max_30_df <- data.frame(matrix(ncol = length(years), nrow = 30))
midstream_max_30_df <- data.frame(matrix(ncol = length(years), nrow = 30))
downstream_max_30_df <- data.frame(matrix(ncol = length(years), nrow = 30))
max_10_range_df <- data.frame(matrix(ncol = length(years) * 2, nrow = 10))

# Set column names
colnames(upstream_max_30_df) <- as.character(years)
colnames(midstream_max_30_df) <- as.character(years)
colnames(downstream_max_30_df) <- as.character(years)
for (i in 1:length(years)) {
  colnames(max_10_range_df)[(i - 1) * 2 + 1] <- paste(years[i], "_name", sep = "")
  colnames(max_10_range_df)[(i - 1) * 2 + 2] <- paste(years[i], "_num", sep = "")
}

# Loop through years
for (i in years) {
  data <- get(paste("output_", i, sep = ""))
  
  # Process data
  upstream_max_30_df[[as.character(i)]] <- sort(data$upstream_ratio, decreasing = TRUE)[1:30]
  midstream_max_30_df[[as.character(i)]] <- sort(data$midstream_ratio, decreasing = TRUE)[1:30]
  downstream_max_30_df[[as.character(i)]] <- sort(data$downstream_ratio, decreasing = TRUE)[1:30]
  max_10_range_df[paste(i, "_name", sep = "")] <- data$name[1:10]
  max_10_range_df[paste(i, "_num", sep = "")] <- data$all[1:10]
}

# Add worksheets and write data
addWorksheet(wb, "Upstream Max 30")
writeData(wb, "Upstream Max 30", upstream_max_30_df)

addWorksheet(wb, "Midstream Max 30")
writeData(wb, "Midstream Max 30", midstream_max_30_df)

addWorksheet(wb, "Downstream Max 30")
writeData(wb, "Downstream Max 30", downstream_max_30_df)

addWorksheet(wb, "Max 10 Range")
writeData(wb, "Max 10 Range", max_10_range_df)

# Save the workbook
saveWorkbook(wb, "results/Top Thirty Countries with Maximum Risk Propagation Range.xlsx", overwrite = TRUE)

# Create a vector of years from 1995 to 2019
year_vector <- 1995:2019

# Create a list to store correlation results for each year
correlation_list <- list()

# Calculate correlations for each year
for (i in 1995:2019) {
  output <- get(paste("output_", i, sep = ""))
  nodelist_upstream <- get(paste("nodelist_", i, sep = ""))
  names(output)[1] <- "Id"
  data <- inner_join(nodelist_upstream, output, by = "Id")
  
  # Calculate correlations and add them to the list
  correlation_list[[i - 1994]] <- data.frame(
    year = i,
    upstream = cor(data$degree, data$upstream),
    midstream = cor(data$degree, data$midstream),
    downstream = cor(data$degree, data$downstream),
    all = cor(data$degree, data$all)
  )
}

# Merge correlation results into one data frame
corelation <- do.call(rbind, correlation_list)

# Write to a CSV file
write.csv(corelation, "results/Relationship Between Degree and Risk Transmission Range.csv.csv", row.names = FALSE)

# Specific Case Analysis

# Create a new Excel workbook
wb_example <- createWorkbook()

# Loop through countries (CHN and USA)
for (k in c("CHN", "USA")) {
  
  # Loop through years (1995, 2007, 2010, 2019)
  for (t in c(1995, 2007, 2010, 2019)) {
    
    # Get matrices for upstream, midstream, and downstream data for a specific year
    mat_1 <- get(paste("data_", t, "_upstream_matrix", sep = ""))
    mat_2 <- get(paste("data_", t, "_midstream_matrix", sep = ""))
    mat_3 <- get(paste("data_", t, "_downstream_matrix", sep = ""))
    
    # Perform a random attack on the upstream data for the given country
    upstream_infected <- random_attack(mat_1, k)
    
    # Identify the sources in the midstream that are also infected in the upstream
    midstream_source <- intersect(unlist(upstream_infected), row.names(mat_2))
    
    # Perform a random attack on the midstream data using the identified sources
    midstream_infected <- random_attack(mat_2, midstream_source)
    
    # Identify the sources in the downstream that are also infected in the midstream
    downstream_source <- intersect(unlist(midstream_infected), row.names(mat_3))
    
    # Perform a random attack on the downstream data using the identified sources
    downstream_infected <- random_attack(mat_3, downstream_source)
    
    # Convert lists to character vectors of the same length
    max_length <- max(
      sapply(upstream_infected, function(x) if (is.null(x)) 0 else length(x)),
      sapply(midstream_infected, function(x) if (is.null(x)) 0 else length(x)),
      sapply(downstream_infected, function(x) if (is.null(x)) 0 else length(x))
    )
    
    upstream_infected <- sapply(
      upstream_infected,
      function(x) if (is.null(x)) rep(NA, max_length) else c(x, rep(NA, max_length - length(x)))
    )
    midstream_infected <- sapply(
      midstream_infected,
      function(x) if (is.null(x)) rep(NA, max_length) else c(x, rep(NA, max_length - length(x)))
    )
    downstream_infected <- sapply(
      downstream_infected,
      function(x) if (is.null(x)) rep(NA, max_length) else c(x, rep(NA, max_length - length(x)))
    )
    
    # Remove the last column of NA values
    upstream_infected <- upstream_infected[, -c(ncol(upstream_infected))]
    midstream_infected <- midstream_infected[, -c(1, ncol(midstream_infected))]
    downstream_infected <- downstream_infected[, -c(1, ncol(downstream_infected))]
    
    # Convert character vectors to data frames
    df <- data.frame(
      upstream_infected = upstream_infected,
      midstream_infected = midstream_infected,
      downstream_infected = downstream_infected
    )
    
    # Create a new worksheet in the Excel workbook for each country and year combination
    sheet_name <- paste(k, t, sep = "_")
    addWorksheet(wb_example, sheet_name)
    
    # Write the data to the Excel file
    writeData(wb_example, sheet_name, df)
  }
}

# Save the Excel workbook
saveWorkbook(wb_example, "results/Risk Propagation Pathways of China and the United States.xlsx", overwrite = TRUE)
