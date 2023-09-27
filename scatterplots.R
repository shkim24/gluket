library(openxlsx)
library(plotly)
library(cowplot)

create_scatterplot <- function(x_values, y_values, plot_title) {
  scatterplot <- plot_ly(x = x_values, y = y_values, type = "scatter", mode = "markers") %>%
    add_markers() %>%
    layout(
      title = plot_title,
      xaxis = list(title = "X"),
      yaxis = list(title = "Y"),
      showlegend = FALSE
    )
  return(scatterplot)
}

plot_excel <- function(excel_file, sheet_name, x_column, y_columns, start_row, end_row) {
  df <- read.xlsx(excel_file, sheet = sheet_name)
  x_values <- as.numeric(df[start_row:end_row, x_column])
  
  plots <- list()
  
  for (col in y_columns) {
    y_values <- as.numeric(df[start_row:end_row, col])
    column_name <- colnames(df)[col]
    plot_title <- as.character(df[3, col])
    scatterplot <- create_scatterplot(x_values, y_values, plot_title)
    plots[[column_name]] <- scatterplot
  }
  
  num_cols <- 3  # Number of columns in the subplot grid
  num_rows <- ceiling(length(plots) / num_cols)  # Number of rows
  plot_list <- list()
  
  plot_index <- 1
  
  for (row in 1:num_rows) {
    subplots <- list()
    
    for (col in 1:num_cols) {
      if (plot_index <= length(plots)) {
        subplots[[col]] <- plots[[names(plots)[plot_index]]]
        plot_index <- plot_index + 1
      }
    }
    
    plot_list[[row]] <- plot_grid(plotlist = subplots)
  }
  
  final_plot <- plot_grid(plotlist = plot_list, ncol = 1)
  return(final_plot)
}

# Excel file, sheet name, column ranges
excel_file <- "Y:/Treadway Pilot/Ketamine Pilot/Regional Analysis/combined_results_FITT2_results_SI_all_k9_anatomic.xlsx"
sheet_name <- "PSS_correlation_no_outliers"
x_column <- 5
y_columns <- 8:27
start_row <- 4
end_row <- 22

# Call function
final_plot <- plot_excel(excel_file, sheet_name, x_column, y_columns, start_row, end_row)
final_plot
