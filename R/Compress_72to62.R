#' Aggregate MRIO Data
#'
#' This function reads a 73-economy MRIO Excel file, compresses into 63-economies, and saves the compressed file.
#'
#' @param pathin The directory where the MRIO Excel file is stored.
#' @param pathout The directory where the aggregated MRIO file should be saved.
#' @param years The year of the MRIO dataset.
#'
#' @return The file path of the saved aggregated MRIO dataset.
#' @export
#'
#' @examples
#' # Example usage:
#' # aggregate_MRIO("C:/path/to/input", "C:/path/to/output", 2017)
Compress_72to62 <- function(pathin, pathout, years) {
  library(readxl)
  library(dplyr)
  library(data.table)

  # Read MRIOT
  MRIO <- read_excel(paste0(pathin, "/ADB-MRIO-", years, ".xlsx"),
                     range = "E8:DHM2570",
                     col_names = FALSE,
                     col_types = "numeric"
  ) %>%
    replace(is.na(.), 0) %>%
    as.matrix()

  # Define key indices
  num_rows <- nrow(MRIO)
  num_cols <- ncol(MRIO)
  num_excluded <- 48
  num_sectors <- 35
  num_economies <- 10

  # Compute number of rows to aggregate
  num_aggregated_rows <- num_economies * num_sectors
  num_target_rows <- num_sectors

  # Identify relevant row indices
  start_aggregated <- num_rows - num_excluded - num_aggregated_rows + 1
  end_aggregated <- num_rows - num_excluded
  start_target <- end_aggregated + 1
  end_target <- start_target + num_target_rows - 1

  # Aggregate and distribute values into the 35 RoW rows
  aggregated_part <- MRIO[start_aggregated:end_aggregated, ]
  for (i in 1:num_target_rows) {
    target_row <- start_target + (i - 1)
    source_indices <- seq(i, num_aggregated_rows, by = num_target_rows)
    MRIO[target_row, ] <- MRIO[target_row, ] + colSums(aggregated_part[source_indices, , drop = FALSE])
  }

  # Remove the now redundant aggregated rows
  MRIO <- MRIO[-(start_aggregated:end_aggregated), ]

  # Column-Wise Aggregation
  ## Intermediates Matrix
  MRIO_part1 <- MRIO[, 1:2555]
  num_compress1 <- num_aggregated_rows
  num_target_cols1 <- 35

  start_compress1 <- 2555 - 35 - num_compress1
  end_compress1 <- 2555 - 35 - 1
  start_target1 <- 2555 - 34
  end_target1 <- 2555

  for (j in 1:num_target_cols1) {
    target_col <- start_target1 + (j - 1)
    source_indices <- seq(j, num_compress1, by = num_target_cols1) + start_compress1
    MRIO_part1[, target_col] <- MRIO_part1[, target_col] + rowSums(MRIO_part1[, source_indices, drop = FALSE])
  }

  MRIO_part1 <- MRIO_part1[, -((start_compress1 + 1):(end_compress1 + 1))]

  ## Final Demand
  MRIO_part2 <- MRIO[, 2556:2920]
  num_compress2 <- 50
  num_target_cols2 <- 5

  start_compress2 <- 365 - 5 - num_compress2
  end_compress2 <- 365 - 5 - 1
  start_target2 <- 365 - 4
  end_target2 <- 365

  for (j in 1:num_target_cols2) {
    target_col <- start_target2 + (j - 1)
    source_indices <- seq(j, num_compress2, by = num_target_cols2) + start_compress2
    MRIO_part2[, target_col] <- MRIO_part2[, target_col] + rowSums(MRIO_part2[, source_indices, drop = FALSE])
  }

  MRIO_part2 <- MRIO_part2[, -((start_compress2 + 1):(end_compress2 + 1))]

  # Compile Final MRIO
  MRIO_part3 <- MRIO[, 2921, drop = FALSE]
  MRIO_aggregated <- cbind(MRIO_part1, MRIO_part2, MRIO_part3)

  # Save the Output
  output_file <- paste0(pathout, "/MRIO_Aggregated_", years, ".csv")
  fwrite(MRIO_aggregated, output_file, row.names = FALSE)

  return(output_file)
}
