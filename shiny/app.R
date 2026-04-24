# READ ME -----------------------------------------------------------------
#
#      Authors: Sarah Forrest (sforrest@cdc.gov) Paul Scanlon (former CDC)
#      Created: 1 April 2025
#  Last edited: 10 April 2026
# Organization: CDC/NCHS/DRM
#      Purpose: Shiny app for streamlined analysis and reporting of complex 
#               survey data
#
# -------------------------------------------------------------------------

# =========================
# Setup and Libraries
# ========================= 

library(shiny)
library(shinyFeedback)
library(shinyjs)
library(haven) # for the various read_ functions
library(readxl) # for reading in Excel files
library(tidyverse) # for piping
library(glue) # for text manipulation
library(DT) # for aesthetic table display
library(labelled)
library(sjlabelled) # to read data labels if present
library(survey) # to create survey design object
library(srvyr) # to create survey design object
library(flextable) # to generate tables
library(surveytable) # to suppress low-precision estimates
library(ggplot2) # to generate plots
library(quarto) # to render a report Quarto document to a Word document
library(officer) # to output a flextable to a Word document
library(openxlsx) # to output Excel files

# Increase maximum file size for uploads
options(shiny.maxRequestSize = 300 * 1024^2)

# Specify location of the rendered HTML for the user guide
shiny::addResourcePath("docs", "docs")

# =========================
# Functions
# ========================= 

# Read in a data file based on its extension (supports CSV, xls/xlsx, SAS, SPSS, and RData/RDS)
read_file <- function(file_path) {
  
  # Get the file extension
  file_extension <- tools::file_ext(file_path)
  
  # Use if-else statements to determine how to read the file
  if (file_extension == "csv") {
    data <- read.csv(file_path)
    
  } else if (file_extension %in% c("xlsx", "xls")) {
    data <- read_excel(file_path)
    
  } else if (file_extension == "sas7bdat") {
    data <- read_sas(file_path)
    
  } else if (file_extension == "sav") {
    data <- read_sav(file_path)
    
  } else if (file_extension %in% c("rds", "RDS", "Rds")) {
    data <- readRDS(file_path)
    
  } else if (file_extension %in% c("RData", "rdata", "Rdata", "Rda", "rda")) {
    load(file_path)
    data <- ls() %>% tail(1) %>% get() # Return the last loaded object
    
  } else {
    stop("Unsupported file type: ", file_extension)
  }
  
  return(data)
  
}

# Create a codebook (variable table) from data labels
create_variable_table <- function(data) {
  codebook <- enframe(get_label(data))
  colnames(codebook) <- c("Variable", "Variable label") # Rename columns
  return(codebook)
}

is_continuous <- function(data, var, threshold = 10) {
  x <- data[[var]]
  is.numeric(x) && !inherits(x, "labelled") && length(unique(na.omit(x))) > threshold
}

# Get custom NCHS color palette (based on the number of levels)
get_nchs_colors <- function(n) {
  color_palette <- c("#006858", "#008BB0", "#D06F1A", "#FFD200", "#695E4A", "#0057b7") # Define NCHS colors
  return(color_palette[1:n]) # Return the first n colors
}

# Get response options for a given variable
get_response_options <- function(data, vars) {
  lapply(vars, function(v) {
    # Use labelled::to_factor if variable is labelled, else as.factor
    if ("labelled" %in% class(data[[v]])) {
      levels(labelled::to_factor(data[[v]]))
    } else {
      levels(as.factor(data[[v]]))
    }
  })
}


# =========================
# Objects
# ========================= 

# Define a default ggplot2 theme for plots
default_theme = theme(
  axis.text.x = element_text(angle = 45, hjust = 1),
  legend.title = element_blank(),
  plot.caption = element_text(hjust = 0),
  plot.title = element_text(hjust = 0.5, margin = margin(t = 0, r = 0, b = 10, l = 1)),
  plot.subtitle = element_text(hjust = 0.5),
  axis.title.x = element_text(margin = margin(t = 10, r = 0, b = 0, l = 1)),
  axis.title.y = element_text(margin = margin(t = 0, r = 10, b = 0, l = 1)),
  axis.ticks.x = element_blank(),
  axis.ticks.y = element_blank(),
  panel.background = element_blank(),
  # plot.background = element_blank()
)

# Define a custom NCHS-style ggplot2 theme for plots
nchs_theme = theme(
  # text = element_text(family = "serif"),
  # axis.text.x = element_text(angle = 45, hjust = 1),
  # legend.title = element_blank(),
  # plot.caption = element_text(hjust = 0),
  # plot.title = element_text(hjust = 0.5, margin = margin(t = 0, r = 0, b = 10, l = 1)),
  # plot.subtitle = element_text(hjust = 0.5),
  # axis.title.x = element_text(margin = margin(t = 10, r = 0, b = 0, l = 1)),
  # axis.title.y = element_text(margin = margin(t = 0, r = 10, b = 0, l = 1)),
  # axis.ticks.x = element_blank(),
  # axis.ticks.y = element_blank(),
  # panel.background = element_blank(),
  # plot.background = element_blank()
  text = element_text(family = "sans", color = "black", size = 12),
  axis.text.x = element_text(angle = 0, hjust = 0.5, color = "black"),
  axis.text.y = element_text(color = "black"),
  axis.title.x = element_text(margin = margin(t = 10)),
  axis.title.y = element_text(margin = margin(r = 10)),
  axis.ticks.x = element_blank(),
  axis.ticks.y = element_line(color = "black"),
  axis.line = element_line(color = "black"),
  panel.background = element_blank(),
  # plot.background = element_blank(),
  panel.grid.major = element_blank(),
  panel.grid.minor = element_blank(),
  legend.title = element_blank(),
  plot.title = element_text(face = "bold", margin = margin(b = 10)),
  plot.caption = element_text(hjust = 0, size = 10, margin = margin(t = 20))
)

# List of available plot themes
theme_list = c(
  "BW" = "theme_bw",
  "Classic" = "theme_classic",
  "Dark" = "theme_dark",
  "Gray" = "theme_gray",
  "Light" = "theme_light",
  "Line Drawing" = "theme_linedraw",
  "Minimal" = "theme_minimal",
  "NCHS" = "nchs_theme"
)

# =========================
# UI: SurveyLand
# =========================

ui <- fluidPage(
  shinyFeedback::useShinyFeedback(),  # For feedback/warning notices in UI
  shinyjs::useShinyjs(),              # For enabling/disabling buttons
  
  # ---- Custom CSS ----
  tags$head(
    tags$style(HTML("
      a { color: white; }
      .custom-link { color: white; text-decoration: underline; }
      .custom-link:hover { color: #527DF3; }
      body { background-color: #004137; color: white; }
      .custom-box {
        background-color: #006858;
        padding: 15px;
        border-radius: 5px;
        border: 1px solid white;
      }
      .nav-tabs > li > a:hover {
        background-color: white;
        color: black;
      }
      .well { background-color: #006858; }
      .toggle-section {
        border-left: 3px solid rgba(255,255,255,0.3);
        padding-left: 12px;
        margin: 4px 0 12px 4px;
      }
    "))
  ),
  
  titlePanel("SurveyLand"),
  actionButton("user_guide", "User Guide"),
  hr(),
  
  # =========================
  # Data Input & Survey Info
  # =========================
  
  tabsetPanel(
    tabPanel(
      "Data Input and Survey Information",
      sidebarLayout(
        sidebarPanel(
          h4("Step 1: Select a data file for analysis"),
          p("Files must be in CSV, Excel, SAS, SPSS, or R format."),
          fileInput(
            "upload", "Upload data file",
            accept = c(
              ".csv", ".xlsx", ".xls", ".sas7bdat", ".sav",
              ".Rda", ".RData", ".rdata", ".rda", ".Rdata",
              "rds", "RDS", "Rds"
            )
          ),
          verbatimTextOutput("upload_summary"),
          conditionalPanel(
            "output.file_uploaded == true",
            br(),
            strong("To display the codebook, select the box below"),
            checkboxInput("codebook_button", "Display codebook?", value = FALSE),
          ),
          conditionalPanel(
            "input.codebook_button == true",
            div(
              style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
              dataTableOutput("codebook")
            )
          ),
        ),
        
        mainPanel(
          tabsetPanel(
            # ---- Survey Metadata ----
            tabPanel(
              "Survey Metadata",
              br(),
              div(
                class = "custom-box",
                h4("Step 2: Enter survey metadata information"),
                p("Enter applicable details to format the data source caption and title for tables and plots."),
                textInput(
                  "data_producer",
                  label = "Enter the data producer",
                  value = "National Center for Health Statistics",
                  width = "400px"
                ),
                textInput(
                  "survey_name",
                  label = "Enter the survey name",
                  value = "",
                  width = "400px"
                ),
                textInput(
                  "survey_round",
                  label = "Enter the survey round or cycle",
                  value = "",
                  width = "400px"
                ),
                textInput(
                  "survey_date",
                  label = "Enter the data collection date or period",
                  value = "",
                  width = "400px"
                ),
                textInput(
                  "geographic_area",
                  label = "Enter the geographic area of data collection",
                  value = "United States",
                  width = "400px"
                ),
                actionButton("surv_meta_submit", "Preview data source caption"),
                br(), 
                br(),
                textOutput("surv_meta_summary")
              )
            ),
            
            # ---- Data Manipulation ----
            tabPanel(
              "Data Manipulation",
              br(),
              div(
                class = "custom-box",
                h4("Step 3: Filter the data"),
                conditionalPanel(
                  "output.file_uploaded == true",
                  strong("To filter the data, select the box below"),
                  p("Note: For complex survey analyses, filtering should be limited to data cleaning, such as removing incomplete cases. Subsetting the data before creating the survey design object can produce incorrect standard errors. Filtering for subgroup analyses should be used ", strong("only"), " in analyses that do not use survey design information, such as unclustered and unweighted analyses."),
                  checkboxInput("filter_dataset", "Filter the data?", value = FALSE),
                  conditionalPanel(
                    "input.filter_dataset == false",
                    p(em("By default, the complete, unfiltered data file will be used."))
                  ),
                  conditionalPanel(
                    "input.filter_dataset == false",
                    verbatimTextOutput("no_filter_message")
                  ),
                  conditionalPanel(
                    "input.filter_dataset == true",
                    div(
                      class = "toggle-section",
                      selectizeInput(
                        "filtered_var",
                        "Select the variable to filter on",
                        choices = NULL,
                        multiple = FALSE,
                        options = list(
                          placeholder = "Select one",
                          onInitialize = I('function() { this.setValue(""); }')
                        ),
                        width = "400px"
                      ),
                      conditionalPanel(
                        "input.filtered_var != ''",
                        selectizeInput(
                          "filtered_var_value",
                          "Specify the value(s) to filter on",
                          choices = NULL,
                          multiple = TRUE,
                          width = "400px"
                        )
                      )
                    ),
                    conditionalPanel(
                      "input.filtered_var != '' && input.filtered_var_value.length > 0",
                      checkboxInput(
                        "filter_second_var",
                        "Filter on a second variable?",
                        value = FALSE
                      ),
                      conditionalPanel(
                        "input.filter_second_var == true",
                        div(
                          class = "toggle-section",
                          selectizeInput(
                            "filtered_var2",
                            "Select the second variable to filter on",
                            choices = NULL,
                            multiple = FALSE,
                            options = list(
                              placeholder = "Select one",
                              onInitialize = I('function() { this.setValue(""); }')
                            ),
                            width = "400px"
                          ),
                          conditionalPanel(
                            "input.filtered_var2 != ''",
                            selectizeInput(
                              "filtered_var_value2",
                              "Specify the value(s) to filter on for the second variable",
                              choices = NULL,
                              multiple = TRUE,
                              width = "400px"
                            )
                          )
                        )
                      )
                    ),
                    conditionalPanel(
                      condition = "(input.filtered_var != '' && input.filtered_var_value.length > 0 && input.filter_second_var == false) || (input.filter_second_var == true && input.filtered_var2 != '' && input.filtered_var_value2.length > 0)",
                      actionButton("filter_submit", "Submit")
                    ),
                    br(),
                    verbatimTextOutput("filtered_summary")
                  )
                ),
                conditionalPanel(
                  "output.file_uploaded == false",
                  h5(strong("Upload a data file to enable filtering options."))
                )
              )
            ),
            
            # ---- Weighting and Design ----
            tabPanel(
              "Weighting and Design",
              br(),
              div(
                class = "custom-box",
                h4("Step 4: Select weighting and survey design approach"),
                conditionalPanel(
                  "output.file_uploaded == true",
                  strong("To account for complex survey design, select each feature to include and the corresponding variable(s)"),
                  p("Note: If missing values are detected for a selected variable, a warning will display."),
                  
                  # -- Cluster/PSU toggle --
                  checkboxInput("use_ids", "Use a cluster / PSU variable?", value = FALSE),
                  conditionalPanel(
                    "input.use_ids == true",
                    div(
                      class = "toggle-section",
                      selectizeInput(
                        "dynamic_select_ids",
                        "Select the cluster/PSU variable",
                        choices = NULL, multiple = FALSE,
                        options = list(placeholder = "Select one",
                                       onInitialize = I('function() { this.setValue(""); }')),
                        width = "400px"
                      )
                    )
                  ),
                  
                  # -- Strata toggle --
                  checkboxInput("use_strata", "Use a strata variable?", value = FALSE),
                  conditionalPanel(
                    "input.use_strata == true",
                    div(
                      class = "toggle-section",
                      selectizeInput(
                        "dynamic_select_strata",
                        "Select the strata variable",
                        choices = NULL, multiple = FALSE,
                        options = list(placeholder = "Select one",
                                       onInitialize = I('function() { this.setValue(""); }')),
                        width = "400px"
                      )
                    )
                  ),
                  
                  # -- Weights toggle --
                  checkboxInput("use_weights", "Use survey weights?", value = FALSE),
                  conditionalPanel(
                    "input.use_weights == true",
                    div(
                      class = "toggle-section",
                      selectizeInput(
                        "dynamic_select_weight",
                        "Select the weighting variable",
                        choices = NULL, multiple = FALSE,
                        options = list(placeholder = "Select one",
                                       onInitialize = I('function() { this.setValue(""); }')),
                        width = "400px"
                      )
                    )
                  ),
                  conditionalPanel(
                    "input.use_ids == false && input.use_strata == false && input.use_weights == false",
                    p(em("By default, an unweighted, unstratified, and unclustered survey design object will be used."))
                  ),
                  actionButton("surv_design_submit", "Submit"),
                  br(), 
                  br(),
                  verbatimTextOutput("surv_design_summary")
                ),
                conditionalPanel(
                  "output.file_uploaded == false",
                  h5(strong("Upload a data file to enable weighting and design options."))
                )
              )
            )
          )
        )
      )
    ),
    
    # =========================
    # Data Analysis
    # =========================
    
    tabPanel(
      "Data Analysis",
      sidebarLayout(
        sidebarPanel(
          h4("Select analytical approach"),
          p(),
          radioButtons(
            "analysis_type",
            "Select the type of analysis",
            choices = c("One-way (single-variable)", "Two-way (bi-variable)", "Multivariable"),
            selected = character(0)
          )
        ),
        mainPanel(
          tabsetPanel(
            # ---- Tables ----
            tabPanel(
              "Tables",
              tabsetPanel(
                tabPanel(
                  "Output",
                  br(),
                  div(
                    class = "custom-box",
                    h4("Select variable(s) to display the table below"),
                    # Only show options if a data file is uploaded and analysis type is selected
                    conditionalPanel(
                      condition = "output.file_uploaded == true && input.analysis_type != null && input.analysis_type != ''",
                      p("Note: Continuous variables will display summary statistics and categorical variables will display percent distributions."),
                      conditionalPanel(
                        condition = "(input.analysis_type == 'One-way (single-variable)' || input.analysis_type == 'Two-way (bi-variable)')",
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_outcome_table",
                              "Select outcome variable",
                              choices = NULL,
                              multiple = FALSE,
                              options = list(
                                placeholder = "Select one",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 6,
                            textInput(
                              "outcome_label_table",
                              "Optionally, enter outcome variable label",
                              value = ""
                            )
                          )
                        )
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Two-way (bi-variable)'",
                        br(),
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_covariate_table",
                              "Select covariate variable",
                              choices = NULL,
                              multiple = FALSE,
                              options = list(
                                placeholder = "Select one",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 6,
                            textInput(
                              "covariate_label_table",
                              "Optionally, enter covariate variable label",
                              value = ""
                            )
                          )
                        ),
                        p("Note: Covariate variable values will display in rows and outcome variable values will display in columns.")
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Multivariable'",
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_multivariable_table",
                              "Select variables",
                              choices = NULL,
                              multiple = TRUE,
                              options = list(
                                placeholder = "Select",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 7,
                            textInput(
                              "multivariable_label_table",
                              "Optionally, enter variable labels (comma separated)",
                              value = "",
                              width = "100%"
                            )
                          )
                        ),
                        p("Note: Selected variables must be either all continuous or all categorical; mixed variable types are not supported.")
                      ),
                      br(),
                      conditionalPanel(
                        "input.analysis_type == 'One-way (single-variable)' && input.dynamic_select_outcome_table != ''",
                        div(
                          style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                          uiOutput("one_way_table")
                        ),
                        br(),
                        downloadButton("download_one_way_table_image", "Download image"),
                        downloadButton("download_one_way_table_word", "Download Word document"),
                        downloadButton("download_one_way_table_excel", "Download Excel file"),
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Two-way (bi-variable)' && input.dynamic_select_outcome_table != '' && input.dynamic_select_covariate_table != ''",
                        div(
                          style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                          uiOutput("two_way_table")
                        ),
                        br(),
                        downloadButton("download_two_way_table_image", "Download image"),
                        downloadButton("download_two_way_table_word", "Download Word document"),
                        downloadButton("download_two_way_table_excel", "Download Excel file"),
                      )
                    ),
                    conditionalPanel(
                      "input.analysis_type == 'Multivariable' && input.dynamic_select_multivariable_table != ''",
                      div(
                        style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                        uiOutput("multivariable_table")
                      ),
                      br(),
                      downloadButton("download_multivariable_table_image", "Download image"),
                      downloadButton("download_multivariable_table_word", "Download Word document"),
                      downloadButton("download_multivariable_table_excel", "Download Excel file")
                    ),
                    conditionalPanel(
                      condition = "output.file_uploaded == false || (input.analysis_type == null || input.analysis_type == '')",
                      h5(strong("Upload a data file and select analytical approach to enable table options."))
                    )
                  )
                ),
                tabPanel(
                  "Options",
                  br(),
                  div(
                    class = "custom-box",
                    h4("Select data presentation preferences for table generation"),
                    conditionalPanel(
                      condition = "output.file_uploaded == true && (input.analysis_type != null || input.analysis_type != '') &&
                      ((input.analysis_type == 'One-way (single-variable)' && input.dynamic_select_outcome_table != '') ||
                      (input.analysis_type == 'Two-way (bi-variable)' && input.dynamic_select_outcome_table != '' && input.dynamic_select_covariate_table != '') ||
                      (input.analysis_type == 'Multivariable' && input.dynamic_select_multivariable_table != ''))",
                      
                      # ---- One-way / Two-way with categorical outcome ----
                      conditionalPanel(
                        condition = "(input.analysis_type == 'One-way (single-variable)' || input.analysis_type == 'Two-way (bi-variable)') && output.outcome_is_continuous_table == false",
                        h5(strong(HTML(paste0(
                          "To suppress low-precision estimates according to NCHS data presentation standards using the ",
                          tags$a(href = "https://cdcgov.github.io/surveytable/", class = "custom-link", "surveytable package"),
                          ", select the box below")))
                        ),
                        checkboxInput(
                          "nchs_presentation_standard",
                          width = "800px",
                          label = "Suppress low-precision estimates?",
                          value = FALSE
                        ),
                        h5(strong("To display row-level totals, select the box below")),
                        checkboxInput("row_n", "Display Ns?", value = FALSE)
                        ),
                      
                      # ---- One-way / Two-way with continuous outcome ----
                      conditionalPanel(
                        condition = "(input.analysis_type == 'One-way (single-variable)' || input.analysis_type == 'Two-way (bi-variable)') && output.outcome_is_continuous_table == true",
                        h5(strong("To display the percent of known values, select the box below")),
                        checkboxInput("percent_known", "Display percent known?", value = FALSE)
                      ),
                      
                      # ---- Multivariable ----
                      conditionalPanel(
                        condition = "input.analysis_type == 'Multivariable'",
                        # Suppression option — only for categorical multivariable
                        conditionalPanel(
                          condition = "output.multivariable_is_continuous == false && output.multivariable_is_mixed == false",
                          h5(strong(HTML(paste0(
                            "To suppress low-precision estimates according to NCHS data presentation standards using the ",
                            tags$a(href = "https://cdcgov.github.io/surveytable/", class = "custom-link", "surveytable package"),
                            ", select the box below")))
                            ),
                          checkboxInput(
                            "nchs_presentation_standard",
                            width = "800px",
                            label = "Suppress low-precision estimates?",
                            value = FALSE
                          )
                        ),
                        # Percent-known option — only for continuous multivariable
                        conditionalPanel(
                          condition = "output.multivariable_is_continuous == true",
                          checkboxInput("percent_known_multivariable", "Display percent known values?", value = FALSE)
                        ),
                        # Mixed-type warning message in the options panel
                        conditionalPanel(
                          condition = "output.multivariable_is_mixed == true",
                          p(strong("Mixed variable types detected. Please select all continuous or all categorical variables to generate a multivariable table."))
                        )
                      )
                    ),
                    conditionalPanel(
                      condition = "output.file_uploaded == false || (input.analysis_type == null || input.analysis_type == '') ||
                      (input.analysis_type == 'One-way (single-variable)' && (input.dynamic_select_outcome_table == '' || input.dynamic_select_outcome_table == null)) ||
                      (input.analysis_type == 'Two-way (bi-variable)' && ((input.dynamic_select_outcome_table == '' || input.dynamic_select_outcome_table == null) || (input.dynamic_select_covariate_table == '' || input.dynamic_select_covariate_table == null))) ||
                      (input.analysis_type == 'Multivariable' && (input.dynamic_select_multivariable_table == '' || input.dynamic_select_multivariable_table == null))",
                      h5(strong("Upload a data file, select analytical approach, and select variable(s) to enable table options."))
                    )
                  )
                )
              )
            ),
            # ---- Plots ----
            tabPanel(
              "Plots",
              tabsetPanel(
                tabPanel(
                  "Output",
                  br(),
                  div(
                    class = "custom-box",
                    h4("Select variable(s) to display the plot below"),
                    conditionalPanel(
                      condition = "output.file_uploaded == true && input.analysis_type != null && input.analysis_type != ''",
                      conditionalPanel(
                        "(input.analysis_type == 'One-way (single-variable)' || input.analysis_type == 'Two-way (bi-variable)')",
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_outcome_plot",
                              "Select outcome variable",
                              choices = NULL,
                              multiple = FALSE,
                              options = list(
                                placeholder = "Select one",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 6,
                            textInput(
                              "outcome_label_plot",
                              "Optionally, enter outcome variable label",
                              value = ""
                            )
                          )
                        )
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Two-way (bi-variable)'",
                        br(),
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_covariate_plot",
                              "Select covariate variable",
                              choices = NULL,
                              multiple = FALSE,
                              options = list(
                                placeholder = "Select one",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 7,
                            textInput(
                              "covariate_label_plot",
                              "Optionally, enter covariate variable label",
                              value = ""
                            )
                          )
                        ),
                        p("Note: Covariate variable values will display as groups along the x-axis and outcome variable values will display as bar fill colors.")
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Multivariable'",
                        fluidRow(
                          column(
                            width = 5,
                            selectizeInput(
                              "dynamic_select_multivariable_plot",
                              "Select variables",
                              choices = NULL,
                              multiple = TRUE,
                              options = list(
                                placeholder = "Select",
                                onInitialize = I('function() { this.setValue(""); }')
                              ),
                              width = "100%"
                            )
                          ),
                          column(
                            width = 7,
                            textInput(
                              "multivariable_label_plot",
                              "Optionally, enter variable labels (comma separated)",
                              value = "",
                              width = "100%"
                            )
                          )
                        ),
                      ),
                      br(),
                      conditionalPanel(
                        "input.analysis_type == 'One-way (single-variable)' && input.dynamic_select_outcome_plot != ''",
                        div(
                          style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                          plotOutput("one_way_plot", width = 800, height = 600)
                        ),
                        br(),
                        downloadButton("download_one_way_plot", "Download image"),
                        actionButton("add_one_way_plot_to_report", "Add to report")
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Two-way (bi-variable)' && input.dynamic_select_outcome_plot != '' && input.dynamic_select_covariate_plot != ''",
                        div(
                          style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                          plotOutput("two_way_plot", width = 800, height = 600)
                        ),
                        br(),
                        downloadButton("download_two_way_plot", "Download image"),
                        actionButton("add_two_way_plot_to_report", "Add to report")
                      ),
                      conditionalPanel(
                        "input.analysis_type == 'Multivariable' && input.dynamic_select_multivariable_plot != ''",
                        div(
                          style = "background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #dcdcdc;",
                          plotOutput("multivariable_plot", width = 800, height = 600)
                        ),
                        br(),
                        downloadButton("download_multivariable_plot", "Download image"),
                        actionButton("add_multivariable_plot_to_report", "Add to report")
                      )
                    ),
                    conditionalPanel(
                      condition = "output.file_uploaded == false || (input.analysis_type == null || input.analysis_type == '')",
                      h5(strong("Upload a data file and select analytical approach to enable plotting options."))
                    )
                  )
                ),
                tabPanel(
                  "Options",
                  br(),
                  div(
                    class = "custom-box",
                    h4("Select data presentation preferences for plot generation"),
                    conditionalPanel(
                      condition = "output.file_uploaded == true && (input.analysis_type != null || input.analysis_type != '') &&
                      ((input.analysis_type == 'One-way (single-variable)' && input.dynamic_select_outcome_plot != '') ||
                      (input.analysis_type == 'Two-way (bi-variable)' && input.dynamic_select_outcome_plot != '' && input.dynamic_select_covariate_plot != '') ||
                      (input.analysis_type == 'Multivariable' && input.dynamic_select_multivariable_plot != ''))",
                      conditionalPanel(
                        condition = "input.analysis_type == 'Multivariable'",
                        h5(strong("To flip the x and y-axis, select the box below")),
                        checkboxInput("plot_axis_flip", "Flip axes?", value = FALSE),
                        radioButtons(
                          "plot_bar_position",
                          "Select a bar position",
                          choices = c("Stacked", "Side-by-side (dodged)"),
                          inline = TRUE,
                          selected = "Stacked"
                        )
                      ),
                      h5(strong("To display value labels, select the box below")),
                      checkboxInput("value_labels", "Display value labels?", value = FALSE),
                      selectInput(
                        "plot_theme",
                        "Select a ggplot or NCHS theme",
                        choices = theme_list,
                        selected = NULL,
                        width = "400px"
                      ),
                      h5(strong("To overwrite plot labels, enter custom text below")),
                      textInput("plot_title",    "Enter a title",       value = "", width = "400px"),
                      textInput("plot_subtitle", "Enter a subtitle",    value = "", width = "400px"),
                      textInput("plot_xlab",     "Enter x-axis label",  value = "", width = "400px"),
                      textInput("plot_ylab",     "Enter y-axis label",  value = "", width = "400px"),
                      textInput("plot_caption",  "Enter a caption", value = "", width = "400px"),
                      conditionalPanel(
                        condition = "input.analysis_type == 'Two-way (bi-variable)' || input.analysis_type == 'Multivariable'",
                        textInput("plot_legend_title", "Enter legend label", value = "", width = "400px")
                      )
                    ),
                    conditionalPanel(
                      condition = "output.file_uploaded == false || (input.analysis_type == null || input.analysis_type == '') ||
                      (input.analysis_type == 'One-way (single-variable)' && (input.dynamic_select_outcome_plot == '' || input.dynamic_select_outcome_plot == null)) ||
                      (input.analysis_type == 'Two-way (bi-variable)' && ((input.dynamic_select_outcome_plot == '' || input.dynamic_select_outcome_plot == null) || (input.dynamic_select_covariate_plot == '' || input.dynamic_select_covariate_plot == null))) ||
                      (input.analysis_type == 'Multivariable' && (input.dynamic_select_multivariable_plot == '' || input.dynamic_select_multivariable_plot == null))",
                      h5(strong("Upload a data file, select analytical approach, and select variable(s) to enable plot options."))
                    )
                  )
                )
              )
            ),
            # ---- Report ----
            tabPanel(
              "Report",
              br(),
              div(
                class = "custom-box",
                h4("Basic report generation"),
                p("Note: Generating the report may take a moment."),
                h5(strong("Step 1. Generate the report")),
                actionButton("generate_report", "Generate report"),
                br(),
                br(),
                h5(strong("Step 2. Download the report (once generated)")),
                downloadButton("download_report", "Download report")
              )
            )
          )
        )
      )
    )
  )
)

# =========================
# Server
# ========================= 

server <- function(input, output, session) {
  
  # =========================
  # Reactives
  # =========================
  
  # ---- Data Upload ----
  data <- reactive({
    req(input$upload)
    read_file(file_path = input$upload$datapath)
  })
  
  upload_summary <- reactive({
    req(data())
    glue(
      "This file has {nrow(data())} records for {length(names(data()))} variables."
    )
  }) %>% bindEvent(input$upload)
  
  variable_name_list <- reactive({
    req(data())
    names(data())
  })
  
  # ---- Survey Metadata ----
  surv_meta_summary <- reactive({
    glue(
      "SOURCE: {input$data_producer}",
      if (input$survey_name != "") paste0(", ", input$survey_name) else "",
      if (input$survey_round != "") paste0(" ", input$survey_round) else "",
      if (input$survey_date != "") paste0(", ", input$survey_date) else "",
      "."
    )
  }) %>% bindEvent(input$surv_meta_submit)
  
  caption <- reactive({
    glue(
      "SOURCE: {input$data_producer}",
      if (input$survey_name != "") paste0(", ", input$survey_name) else "",
      if (input$survey_round != "") paste0(" ", input$survey_round) else "",
      if (input$survey_date != "") paste0(", ", input$survey_date) else "",
      "."
    )
  })
  
  # ---- Data Manipulation ----
  filtered_data <- reactive({
    req(data())
    if (isTRUE(input$filter_dataset)) {
      req(input$filtered_var, input$filtered_var_value)
      filtered <- data() %>%
        filter(
          !is.na(.data[[input$filtered_var]]) &
            .data[[input$filtered_var]] %in% input$filtered_var_value
        )
      if (isTRUE(input$filter_second_var)) {
        req(input$filtered_var2, input$filtered_var_value2)
        filtered <- filtered %>%
          filter(
            !is.na(.data[[input$filtered_var2]]) &
              .data[[input$filtered_var2]] %in% input$filtered_var_value2
          )
      }
      filtered
    } else {
      data()
    }
  })
  
  filtered_summary <- eventReactive(input$filter_submit, {
    req(filtered_data(), input$filtered_var, input$filtered_var_value)
    summary_text <- glue(
      "The filtered file has {nrow(filtered_data())} records based on the variable '{input$filtered_var}' with value(s): {paste(input$filtered_var_value, collapse = ', ')}."
    )
    if (isTRUE(input$filter_second_var)) {
      req(input$filtered_var2, input$filtered_var_value2)
      summary_text <- paste(
        summary_text,
        glue(
          "Additionally filtered by '{input$filtered_var2}' with value(s): {paste(input$filtered_var_value2, collapse = ', ')}."
        )
      )
    }
    summary_text
  })
  
  # ---- Weighting and Design ----
  ids_active <- reactive({
    isTRUE(input$use_ids) &&
      !is.null(input$dynamic_select_ids) &&
      nzchar(input$dynamic_select_ids)
  })
  
  strata_active <- reactive({
    isTRUE(input$use_strata) &&
      !is.null(input$dynamic_select_strata) &&
      nzchar(input$dynamic_select_strata)
  })
  
  weights_active <- reactive({
    isTRUE(input$use_weights) &&
      !is.null(input$dynamic_select_weight) &&
      nzchar(input$dynamic_select_weight)
  })
  
  valid_id <- reactive({
    if (!ids_active()) return(TRUE)
    has_missing <- sum(is.na(filtered_data()[[input$dynamic_select_ids]])) > 0
    shinyFeedback::feedbackWarning(
      "dynamic_select_ids", has_missing,
      "Missing values detected for selected cluster/PSU variable. Please filter them out in Step 3."
    )
    !has_missing
  })
  
  valid_strata <- reactive({
    if (!strata_active()) return(TRUE)
    has_missing <- sum(is.na(filtered_data()[[input$dynamic_select_strata]])) > 0
    shinyFeedback::feedbackWarning(
      "dynamic_select_strata", has_missing,
      "Missing values detected for selected strata variable. Please filter them out in Step 3."
    )
    !has_missing
  })
  
  valid_weight <- reactive({
    if (!weights_active()) return(TRUE)
    has_missing <- sum(is.na(filtered_data()[[input$dynamic_select_weight]])) > 0
    shinyFeedback::feedbackWarning(
      "dynamic_select_weight", has_missing,
      "Missing values detected for selected weight variable. Please filter them out in Step 3."
    )
    !has_missing
  })
  
  data_survey <- reactive({
    req(filtered_data())
    options(survey.lonely.psu = "adjust")
    
    ids_formula     <- ~1
    strata_formula  <- NULL
    weights_formula <- ~1
    
    if (ids_active()     && valid_id())     ids_formula     <- as.formula(paste("~", input$dynamic_select_ids))
    if (strata_active()  && valid_strata()) strata_formula  <- as.formula(paste("~", input$dynamic_select_strata))
    if (weights_active() && valid_weight()) weights_formula <- as.formula(paste("~", input$dynamic_select_weight))
    
    svydesign(
      ids     = ids_formula,
      strata  = strata_formula,
      data    = filtered_data(),
      weights = weights_formula,
      nest    = TRUE
    )
  })
  
  data_survey_srvyr <- reactive({
    req(data_survey())
    data_survey() %>% as_survey_design()
  })
  
  surv_design_summary <- reactive({
    
    if (ids_active()     && !valid_id())     return("The selected cluster/PSU variable has missing values. Please filter them out in Step 3.")
    if (strata_active()  && !valid_strata()) return("The selected strata variable has missing values. Please filter them out in Step 3.")
    if (weights_active() && !valid_weight()) return("The selected weight variable has missing values. Please filter them out in Step 3.")
    
    survey_design_object <- tryCatch({ data_survey() }, error = function(e) NULL)
    if (is.null(survey_design_object)) {
      return("Failed to create the survey design object. Please check your selections.")
    }
    
    paste(
      paste0(
        "Analysis will be ",
        ifelse(ids_active(), "clustered", "unclustered"),
        ", ",
        ifelse(strata_active(), "stratified", "unstratified"),
        ", and ",
        ifelse(weights_active(), "weighted", "unweighted"),
        "."
      ),
      "\n\nA survey design object has been created using the survey package with the following specifications:",
      "\nids     =", ifelse(ids_active(),     input$dynamic_select_ids,    "~1 (unclustered analysis)"),
      "\nstrata  =", ifelse(strata_active(),  input$dynamic_select_strata, "NULL (unstratified analysis)"),
      "\nweights =", ifelse(weights_active(), input$dynamic_select_weight, "~1 (unweighted analysis)"),
      "\nnest    = TRUE"
    )
    
  }) %>% bindEvent(input$surv_design_submit)
  
  # ---- Data Analysis ----
  codebook <- reactive({
    req(input$upload)
    create_variable_table(data())
  })
  
  analysis_type <- reactive({
    list(one_way = ifelse(input$analysis_type == "One-way (single-variable)", "yes", "no"))
  }) %>% bindEvent(input$analysis_type)
  
  outcome_is_continuous_table <- reactive({
    req(input$dynamic_select_outcome_table, filtered_data())
    is_continuous(filtered_data(), input$dynamic_select_outcome_table)
  })
  
  same_two_way_table_vars <- reactive({
    req(input$analysis_type)
    input$analysis_type == "Two-way (bi-variable)" &&
      nzchar(input$dynamic_select_outcome_table) &&
      nzchar(input$dynamic_select_covariate_table) &&
      identical(input$dynamic_select_outcome_table, input$dynamic_select_covariate_table)
  })
  
  multivariable_var_type <- reactive({
    req(input$dynamic_select_multivariable_table, filtered_data())
    vars  <- input$dynamic_select_multivariable_table
    types <- sapply(vars, function(v) is_continuous(filtered_data(), v))
    if      (all(types))  "continuous"
    else if (!any(types)) "categorical"
    else                   "mixed"
  })
  
  percent_known_multivariable <- reactive({
    isTRUE(input$percent_known_multivariable)
  })
  
  nchs_presentation_standard <- reactive({
    isTRUE(input$nchs_presentation_standard)
  })
  
  row_n <- reactive({
    isTRUE(input$row_n)
  })
  
  percent_known <- reactive({
    isTRUE(input$percent_known)
  })
  
  table_title <- reactive({
    if (input$analysis_type %in% c("One-way (single-variable)", "Two-way (bi-variable)")) {
      paste0(
        "Percent distribution of ",
        input$outcome_label_table,
        if (input$analysis_type == "Two-way (bi-variable)") paste0(", by ", input$covariate_label_table) else "",
        if (input$survey_date != "" | input$geographic_area != "") paste0(": ") else "",
        if (input$geographic_area != "") paste0(input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
    } else if (input$analysis_type == "Multivariable") {
      paste0(
        "Percent distribution of ",
        input$multivariable_label_table,
        if (input$survey_date != "" | input$geographic_area != "") paste0(": ") else "",
        if (input$geographic_area != "") paste0(input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
    }
  })
  
  processed_table_one_way <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_outcome_table)
    
    # ---- Continuous path ----
    if (outcome_is_continuous_table()) {
      if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
      invisible(set_survey(data_survey_srvyr()))
      
      result <- as.data.frame(tab(input$dynamic_select_outcome_table)) %>%
        select(-any_of(c("LL", "UL")))
      
      big_n <- filtered_data() %>%
        filter(!is.na(.data[[input$dynamic_select_outcome_table]])) %>%
        nrow()
      
      title <- paste0(
        "Table. Summary statistics for ",
        input$outcome_label_table,
        if (input$geographic_area != "") paste0(": ", input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
      
      if (percent_known()) {
        result <- result %>%
          rename("Percent\nknown" = any_of("X..known"))
        footnote_text <- paste0(
          "Percent known: Percent of values known; SEM: Standard error of the mean; SD: Standard deviation.\n",
          "NOTES: Total number of complete cases: n = ", format(big_n, big.mark = ","), ".\n",
          caption()
        )
      } else {
        result <- result %>%
          select(-any_of("X..known"))
        footnote_text <- paste0(
          "SEM: Standard error of the mean; SD: Standard deviation.\n",
          "NOTES: Total number of complete cases: n = ", format(big_n, big.mark = ","), ".\n",
          caption()
        )
      }
      
      list(table = result, title = title, footnote = footnote_text)
      
      # ---- Categorical path ----
    } else {
      labelled_data_survey_srvyr <- data_survey_srvyr() %>%
        mutate(!!input$dynamic_select_outcome_table := labelled::to_factor(.data[[input$dynamic_select_outcome_table]]))
      if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
      invisible(set_survey(labelled_data_survey_srvyr))
      subset_table <- tab(input$dynamic_select_outcome_table, max_levels = 50, drop_na = TRUE)
      colnames(subset_table) <- make.unique(colnames(subset_table)) # Ensure unique column names for duplicates
      se_col <- if ("SE.1" %in% colnames(subset_table)) "SE.1" else "SE" # Select correct SE column
      columns_to_keep <- c("Level", "n", "Percent", se_col, "Flags")
      columns_to_keep <- columns_to_keep[columns_to_keep %in% colnames(subset_table)] # Only keep present columns
      subset_table <- subset_table %>%
        select(all_of(columns_to_keep)) %>%
        rename(SE = !!se_col) # rename SE column to "SE" for consistency
      if (nchs_presentation_standard()) {
        columns_to_keep <- c("Level", "n", "n_unsuppressed", "Percent", "SE", "Flags")
        if ("Flags" %in% colnames(subset_table)) {
          subset_table <- subset_table %>%
            mutate(
              n_unsuppressed = n,
              n = if_else(str_detect(Flags, "R|Cx|Px"), -9999, n),
              Percent = if_else(str_detect(Flags, "R|Cx|Px"), -9999, Percent),
              SE = if_else(str_detect(Flags, "R|Cx|Px"), -9999, SE)
            ) %>%
            select(any_of(columns_to_keep)) %>%
            mutate(
              Flags = na_if(Flags, ""),
              flags = max(Flags, na.rm = TRUE),
              N = if_else(!is.na(flags), NA_real_, sum(n_unsuppressed, na.rm = TRUE)),
              big_n = sum(n_unsuppressed, na.rm = TRUE)
            )
        } else {
          subset_table <- subset_table %>%
            mutate(n_unsuppressed = n) %>%
            select(any_of(columns_to_keep)) %>%
            mutate(
              N = sum(n_unsuppressed, na.rm = TRUE),
              big_n = sum(n_unsuppressed, na.rm = TRUE)
            )
        }
      } else {
        columns_to_keep <- c("Level", "n", "Percent", "SE")
        subset_table <- subset_table %>%
          select(any_of(columns_to_keep)) %>%
          mutate(
            N = sum(n, na.rm = TRUE),
            big_n = sum(n, na.rm = TRUE)
          )
      }
      if (!row_n()) {
        subset_table <- subset_table %>% select(Level, N, Percent, SE, big_n)
      } else {
        subset_table <- subset_table %>% select(Level, n, N, Percent, SE, big_n)
      }
      title <- paste0("Table. ", table_title())
      big_n <- max(subset_table$big_n, na.rm = TRUE)
      contains_zero <- any(sapply(subset_table, function(col) any(is.numeric(col) & col == 0.0, na.rm = TRUE)))
      contains_suppressed <- any(sapply(subset_table, function(col) any(is.na(col))))
      footnote_text <- paste0(
        "SE: Standard error",
        if (contains_zero) "; 0.0 Quantity more than zero but less than 0.05" else "",
        if (contains_suppressed) "; * Estimate does not meet National Center for Health Statistics standards of reliability" else "",
        if (!contains_zero & !contains_suppressed) "." else ".",
        '\nNOTES: Percents may not add to 100 due to rounding. Total number of complete cases: n = ',
        format(big_n, big.mark = ','),
        ".\n", caption()
      )
      subset_table <- subset_table %>% select(-big_n, -N)
      if ("n" %in% names(subset_table)) {
        subset_table <- subset_table %>% rename("Number" = "n")
      }
      list(
        table = subset_table,
        title = title,
        footnote = footnote_text
      )
    }
  })
  
  flextable_data_one_way <- reactive({
    req(processed_table_one_way())
    set_flextable_defaults(na_str = "---")
    table_data <- processed_table_one_way()$table
    # ---- Continuous path ----
    if (outcome_is_continuous_table()) {
      numeric_cols <- names(table_data)[sapply(table_data, is.numeric)]
      table_data[numeric_cols] <- lapply(table_data[numeric_cols], function(x) {
        sprintf("%.1f", x)
      })
      ft <- flextable(table_data) %>%
        set_table_properties(width = 0.5, layout = "autofit") %>%
        add_footer_lines(processed_table_one_way()$footnote) %>%
        set_caption(caption = processed_table_one_way()$title)
      # ---- Categorical path ----
    } else {
      if ("Number" %in% names(table_data)) {
        table_data$Number <- ifelse(
          table_data$Number == -9999, "*",
          ifelse(
            is.na(table_data$Number), "---",
            format(as.integer(table_data$Number), big.mark = ",", scientific = FALSE)
          )
        )
      }
      # Format other numeric columns as one decimal, with suppression/missing handling
      other_numeric_cols <- setdiff(names(table_data)[sapply(table_data, is.numeric)], "Number")
      table_data[other_numeric_cols] <- lapply(table_data[other_numeric_cols], function(x) {
        out <- ifelse(x == -9999, "*",
                      ifelse(is.na(x), "---", sprintf("%.1f", x)))
        out
      })
      ft <- flextable(table_data) %>%
        set_table_properties(width = 0.5, layout = "autofit") %>%
        add_footer_lines(processed_table_one_way()$footnote) %>%
        set_caption(caption = processed_table_one_way()$title)
    }
    ft
  })
  
  processed_table_two_way <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_outcome_table, input$dynamic_select_covariate_table)
    
    validate(
      need(
        input$dynamic_select_outcome_table != input$dynamic_select_covariate_table,
        "" # Cannot create a two-way table when the outcome and covariate are the same variable
      )
    )
    
    # ---- Continuous path ----
    if (outcome_is_continuous_table()) {
      # Only convert the covariate (subsetting variable) to factor; leave outcome as numeric
      labelled_data_survey_srvyr <- data_survey_srvyr() %>%
        mutate(!!input$dynamic_select_covariate_table := labelled::to_factor(.data[[input$dynamic_select_covariate_table]]))
      
      if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
      invisible(set_survey(labelled_data_survey_srvyr))
      
      result <- as.data.frame(
        tab_subset(
          input$dynamic_select_outcome_table,
          input$dynamic_select_covariate_table,
          max_levels = 50,
          drop_na = TRUE
        )
      ) %>%
        select(-any_of(c("LL", "UL")))
      
      big_n <- filtered_data() %>%
        filter(
          !is.na(.data[[input$dynamic_select_outcome_table]]) &
            !is.na(.data[[input$dynamic_select_covariate_table]])
        ) %>%
        nrow()
      
      title <- paste0(
        "Table. Summary statistics for ",
        input$outcome_label_table,
        ", by ", input$covariate_label_table,
        if (input$geographic_area != "") paste0(": ", input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
      
      if (percent_known()) {
        result <- result %>%
          rename("Percent\nknown" = any_of("X..known"))
        footnote_text <- paste0(
          "Percent known: Percent of values known; SEM: Standard error of the mean; SD: Standard deviation.\n",
          "NOTES: Total number of complete cases: n = ", format(big_n, big.mark = ","), ".\n",
          caption()
        )
      } else {
        result <- result %>%
          select(-any_of("X..known"))
        footnote_text <- paste0(
          "SEM: Standard error of the mean; SD: Standard deviation.\n",
          "NOTES: Total number of complete cases: n = ", format(big_n, big.mark = ","), ".\n",
          caption()
        )
      }
      
      list(table = result, title = title, footnote = footnote_text)
      
      # ---- Categorical path ----
    } else {
      labelled_data_survey_srvyr <- data_survey_srvyr() %>%
        mutate(
          !!input$dynamic_select_outcome_table := labelled::to_factor(.data[[input$dynamic_select_outcome_table]]),
          !!input$dynamic_select_covariate_table := labelled::to_factor(.data[[input$dynamic_select_covariate_table]])
        )
      if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
      invisible(set_survey(labelled_data_survey_srvyr))
      subset_table_list <- tab_subset(
        input$dynamic_select_outcome_table,
        input$dynamic_select_covariate_table,
        max_levels = 50,
        drop_na = TRUE
      )
      subset_table_list <- lapply(subset_table_list, function(df) {
        colnames(df) <- make.unique(colnames(df)) # Ensure unique column names for duplicates
        se_col <- if ("SE.1" %in% colnames(df)) "SE.1" else "SE" # Select correct SE column
        columns_to_keep <- c("Level", "n", "Percent", se_col, "Flags")
        columns_to_keep <- columns_to_keep[columns_to_keep %in% colnames(df)]
        df %>%
          select(all_of(columns_to_keep)) %>%
          rename(SE = !!se_col) # rename SE column to "SE" for consistency
      })
      if (nchs_presentation_standard()) {
        subset_table_list <- lapply(subset_table_list, function(subset_table) {
          if ("Flags" %in% colnames(subset_table)) {
            columns_to_keep <- c("Level", "n", "n_unsuppressed", "Percent", "SE", "Flags")
            subset_table %>%
              mutate(
                n_unsuppressed = n,
                n = if_else(str_detect(Flags, "R|Cx|Px"), -9999, n),
                Percent = if_else(str_detect(Flags, "R|Cx|Px"), -9999, Percent),
                SE = if_else(str_detect(Flags, "R|Cx|Px"), -9999, SE)
              ) %>%
              select(any_of(columns_to_keep))
          } else {
            columns_to_keep <- c("Level", "n", "n_unsuppressed", "Percent", "SE")
            subset_table %>%
              mutate(n_unsuppressed = n) %>%
              select(any_of(columns_to_keep))
          }
        })
      } else {
        columns_to_keep <- c("Level", "n", "Percent", "SE")
        subset_table_list <- lapply(subset_table_list, function(subset_table) {
          subset_table %>% select(any_of(columns_to_keep))
        })
      }
      outcome_vals <- levels(pull(labelled_data_survey_srvyr, !!rlang::sym(input$dynamic_select_outcome_table)))
      combined <- bind_rows(subset_table_list, .id = "Exposure")
      if (nchs_presentation_standard()) {
        combined <- combined %>%
          group_by(Exposure) %>%
          mutate(N = sum(n_unsuppressed, na.rm = TRUE)) %>%
          ungroup() %>%
          mutate(big_n = sum(n_unsuppressed, na.rm = TRUE))
      } else {
        combined <- combined %>%
          group_by(Exposure) %>%
          mutate(N = sum(n, na.rm = TRUE)) %>%
          ungroup() %>%
          mutate(big_n = sum(n, na.rm = TRUE))
      }
      crosstab_table <- combined %>%
        pivot_wider(
          names_from = Level,
          values_from = c(Percent, SE),
          names_glue = "{Level}_{.value}"
        ) %>%
        group_by(Exposure) %>%
        summarise(
          across(everything(), ~ ifelse(all(is.na(.)), NA, first(na.omit(.))), .names = "{col}"),
          .groups = "drop"
        )
      if (row_n()) {
        crosstab_table <- crosstab_table %>%
          select(Exposure, N, any_of(c(starts_with(outcome_vals))), big_n) %>%
          rename("n" = "N")
      } else {
        crosstab_table <- crosstab_table %>%
          select(Exposure, any_of(c(starts_with(outcome_vals))), big_n)
      }
      title <- paste0("Table. ", table_title())
      big_n <- max(crosstab_table$big_n, na.rm = TRUE)
      contains_zero <- any(sapply(crosstab_table, function(col) any(is.numeric(col) & col == 0.0, na.rm = TRUE)))
      contains_suppressed <- any(sapply(crosstab_table, function(col) any(is.numeric(col) & col == -9999, na.rm = TRUE)))
      contains_missing <- any(sapply(crosstab_table, function(col) {
        is.numeric(col) && any(is.na(col))
      }))
      footnote_text <- paste0(
        "SE: Standard error",
        if (contains_zero) "; 0.0 Quantity more than zero but less than 0.05" else "",
        if (contains_suppressed) "; * Estimate does not meet National Center for Health Statistics standards of reliability" else "",
        if (contains_missing) "; --- Data not available" else "",
        if (!contains_zero & !contains_suppressed & !contains_missing) "." else ".",
        '\nNOTES: Percents may not add to 100 due to rounding. Total number of complete cases: n = ',
        format(big_n, big.mark = ','), ".",
        "\n", caption()
      )
      crosstab_table <- crosstab_table %>% rename("Level" = "Exposure") %>% select(-big_n)
      if ("n" %in% names(crosstab_table)) {
        crosstab_table <- crosstab_table %>% rename("Number" = "n")
      }
      list(
        table = crosstab_table,
        title = title,
        footnote = footnote_text
      )
    }
  })
  
  flextable_data_two_way <- reactive({
    req(processed_table_two_way())
    set_flextable_defaults(na_str = "---")
    table_data <- processed_table_two_way()$table
    # ---- Continuous path ----
    if (outcome_is_continuous_table()) {
      numeric_cols <- names(table_data)[sapply(table_data, is.numeric)]
      table_data[numeric_cols] <- lapply(table_data[numeric_cols], function(x) {
        sprintf("%.1f", x)
      })
      ft <- flextable(table_data) %>%
        set_table_properties(width = 0.75, layout = "autofit") %>%
        add_footer_lines(processed_table_two_way()$footnote) %>%
        set_caption(caption = processed_table_two_way()$title)
      # ---- Categorical path ----
    } else {
      # Format the "Number" column as integer, with suppression/missing handling
      if ("Number" %in% names(table_data)) {
        table_data$Number <- ifelse(
          table_data$Number == -9999, "*",
          ifelse(
            is.na(table_data$Number), "---",
            format(as.integer(table_data$Number), big.mark = ",", scientific = FALSE)
          )
        )
      }
      # Format other numeric columns as one decimal, with suppression/missing handling
      other_numeric_cols <- setdiff(names(table_data)[sapply(table_data, is.numeric)], "Number")
      table_data[other_numeric_cols] <- lapply(table_data[other_numeric_cols], function(x) {
        out <- ifelse(x == -9999, "*",
                      ifelse(is.na(x), "---", sprintf("%.1f", x)))
        out
      })
      ft <- flextable(table_data) %>%
        separate_header() %>%
        set_table_properties(width = 0.75, layout = "autofit") %>%
        add_footer_lines(processed_table_two_way()$footnote) %>%
        set_caption(caption = processed_table_two_way()$title)
    }
    ft
  })
  
  processed_table_multivariable <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_multivariable_table)
    
    var_type <- multivariable_var_type()
    validate(
      need(
        var_type != "mixed",
        "" # Cannot create a multivariable table with mixed variable types. Please select variables that are either all continuous or all categorical.
      )
    )
    
    variable_labels   <- strsplit(input$multivariable_label_table, ",")[[1]] %>% trimws()
    variable_names    <- input$dynamic_select_multivariable_table
    variable_label_map <- setNames(
      if (length(variable_labels) == length(variable_names)) variable_labels else variable_names,
      variable_names
    )
    
    title <- paste0(
      "Table. ",
      if (var_type == "continuous") "Summary statistics for " else "Percent distribution of ",
      input$multivariable_label_table,
      if (input$survey_date != "" | input$geographic_area != "") ": " else "",
      if (input$geographic_area != "") input$geographic_area else "",
      if (input$survey_date != "") paste0(", ", input$survey_date) else ""
    )
    
    # ---- Continuous path ----
    if (var_type == "continuous") {
      
      results_list <- lapply(variable_names, function(var) {
        if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
        invisible(set_survey(data_survey_srvyr()))
        row <- as.data.frame(tab(var)) %>%
          select(-any_of(c("LL", "UL"))) %>%
          mutate(Variable = variable_label_map[[var]])
        row
      })
      
      combined <- bind_rows(results_list) %>%
        relocate(Variable, .before = everything())
      
      if (percent_known_multivariable()) {
        combined <- combined %>%
          rename("Percent known" = any_of("X..known"))
      } else {
        combined <- combined %>%
          select(-any_of("X..known"))
      }
      
      total_n <- sum(
        sapply(variable_names, function(var)
          filtered_data() %>% filter(!is.na(.data[[var]])) %>% nrow()
        )
      )
      
      footnote_text <- paste0(
        if (percent_known_multivariable()) "Percent known: Percent of values known; " else "",
        "SEM: Standard error of the mean; SD: Standard deviation.",
        "\nNOTES: Complete-case counts shown per variable.",
        "\n", caption()
      )
      
      return(list(
        table = combined,
        title = title,
        footnote = footnote_text,
        is_continuous = TRUE
      ))
    }
    
    # ---- Categorical path ----
    labelled_data_survey_srvyr <- data_survey_srvyr()
    subset_table_list <- list()
    
    for (var in variable_names) {
      labelled_data_survey_srvyr <- labelled_data_survey_srvyr %>%
        mutate(!!var := labelled::to_factor(.data[[var]]))
      if (nchs_presentation_standard()) suppressMessages(set_opts(mode = "nchs"))
      invisible(set_survey(labelled_data_survey_srvyr))
      subset_table <- tab(var, max_levels = 50, drop_na = TRUE)
      colnames(subset_table) <- make.unique(colnames(subset_table))
      se_col          <- if ("SE.1" %in% colnames(subset_table)) "SE.1" else "SE"
      columns_to_keep <- c("Level", "n", "Percent", se_col, "Flags")
      columns_to_keep <- columns_to_keep[columns_to_keep %in% colnames(subset_table)]
      subset_table <- subset_table %>%
        select(all_of(columns_to_keep)) %>%
        rename(SE = !!se_col) %>%
        mutate(Variable = var)
      
      if (nchs_presentation_standard()) {
        columns_to_keep <- c("Level", "n", "Percent", "SE", "Flags", "Variable")
        if ("Flags" %in% colnames(subset_table)) {
          subset_table <- subset_table %>%
            mutate(
              n       = if_else(str_detect(Flags, "R|Cx|Px"), -9999, n),
              Percent = if_else(str_detect(Flags, "R|Cx|Px"), -9999, Percent),
              SE      = if_else(str_detect(Flags, "R|Cx|Px"), -9999, SE)
            ) %>%
            select(any_of(columns_to_keep))
        } else {
          subset_table <- subset_table %>% select(any_of(columns_to_keep))
        }
      } else {
        subset_table <- subset_table %>%
          select(any_of(c("Level", "n", "Percent", "SE", "Variable")))
      }
      subset_table_list[[var]] <- subset_table
    }
    
    combined <- bind_rows(subset_table_list, .id = "Variable") %>%
      group_by(Variable) %>%
      mutate(N = sum(n, na.rm = TRUE)) %>%
      ungroup()
    
    combined_table <- combined %>%
      pivot_wider(
        names_from  = Level,
        values_from = c(Percent, SE),
        names_glue  = "{Level}_{.value}"
      ) %>%
      group_by(Variable) %>%
      summarise(
        across(everything(), ~ ifelse(all(is.na(.)), NA, first(na.omit(.))), .names = "{col}"),
        .groups = "drop"
      )
    
    multivariable_vals <- unique(unlist(lapply(variable_names, function(var) {
      levels(pull(labelled_data_survey_srvyr, !!rlang::sym(var)))
    })))
    
    combined_table <- combined_table %>%
      select(Variable, any_of(c(starts_with(multivariable_vals)))) %>%
      mutate(Variable = variable_label_map[Variable])
    
    contains_zero       <- any(sapply(combined_table, function(col) any(is.numeric(col) & col == 0.0,    na.rm = TRUE)))
    contains_suppressed <- any(sapply(combined_table, function(col) any(is.numeric(col) & col == -9999,  na.rm = TRUE)))
    contains_missing    <- any(sapply(combined_table, function(col) is.numeric(col) && any(is.na(col))))
    
    footnote_text <- paste0(
      "SE: Standard error",
      if (contains_zero)       "; 0.0 Quantity more than zero but less than 0.05" else "",
      if (contains_suppressed) "; * Estimate does not meet National Center for Health Statistics standards of reliability" else "",
      if (contains_missing)    "; --- Data not available" else "",
      ".",
      "\nNOTES: Percents may not add to 100 due to rounding.",
      "\n", caption()
    )
    
    if ("n" %in% names(combined_table)) {
      combined_table <- combined_table %>% rename("Number" = "n")
    }
    
    list(
      table = combined_table,
      title = title,
      footnote = footnote_text,
      is_continuous = FALSE
    )
  })
  
  flextable_data_multivariable <- reactive({
    req(processed_table_multivariable())
    set_flextable_defaults(na_str = "---")
    result     <- processed_table_multivariable()
    table_data <- result$table
    
    # ---- Continuous path ----
    if (isTRUE(result$is_continuous)) {
      numeric_cols <- names(table_data)[sapply(table_data, is.numeric)]
      table_data[numeric_cols] <- lapply(table_data[numeric_cols], function(x) {
        sprintf("%.1f", x)
      })
      ft <- flextable(table_data) %>%
        set_table_properties(width = 0.75, layout = "autofit") %>%
        add_footer_lines(result$footnote) %>%
        set_caption(caption = result$title)
      return(ft)
    }
    
    # ---- Categorical path ----
    if ("Number" %in% names(table_data)) {
      table_data$Number <- ifelse(
        table_data$Number == -9999, "*",
        ifelse(
          is.na(table_data$Number), "---",
          format(as.integer(table_data$Number), big.mark = ",", scientific = FALSE)
        )
      )
    }
    other_numeric_cols <- setdiff(names(table_data)[sapply(table_data, is.numeric)], "Number")
    table_data[other_numeric_cols] <- lapply(table_data[other_numeric_cols], function(x) {
      ifelse(x == -9999, "*", ifelse(is.na(x), "---", sprintf("%.1f", x)))
    })
    
    flextable(table_data) %>%
      separate_header() %>%
      set_table_properties(width = 0.75, layout = "autofit") %>%
      add_footer_lines(result$footnote) %>%
      set_caption(caption = result$title)
  })
  
  # ---- Plot Generation ----
  plot_title <- reactive({
    if (input$plot_title != "") {
      input$plot_title
    } else if (input$analysis_type %in% c("One-way (single-variable)", "Two-way (bi-variable)")) {
      paste0(
        "Figure. ", "Percent distribution of ",
        input$outcome_label_plot,
        if (input$analysis_type == "Two-way (bi-variable)") paste0(", by ", input$covariate_label_plot) else "",
        if (input$survey_date != "" | input$geographic_area != "") paste0(": ") else "",
        if (input$geographic_area != "") paste0(input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
    } else if (input$analysis_type == "Multivariable") {
      paste0(
        "Figure. ", "Percent distribution of ",
        input$multivariable_label_plot,
        if (input$survey_date != "" | input$geographic_area != "") paste0(": ") else "",
        if (input$geographic_area != "") paste0(input$geographic_area) else "",
        if (input$survey_date != "") paste0(", ", input$survey_date) else ""
      )
    }
  })
  
  plot_subtitle <- reactive({
    if (input$plot_subtitle != "") {
      input$plot_subtitle
    } else {
      NULL
    }
  })
  
  plot_xlab <- reactive({
    req(input$analysis_type)
    if (input$analysis_type == "One-way (single-variable)") {
      if (input$plot_xlab != "") {
        input$plot_xlab
      } else {
        input$dynamic_select_outcome_plot
      }
    } else if (input$analysis_type == "Two-way (bi-variable)") {
      if (input$plot_xlab != "") {
        input$plot_xlab
      } else {
        input$dynamic_select_covariate_plot
      }
    } else if (input$analysis_type == "Multivariable") {
      if (input$plot_xlab != "") {
        input$plot_xlab
      }
    } else {
      NULL
    }
  })
  
  plot_ylab <- reactive({
    if (input$plot_ylab != "") {
      input$plot_ylab
    } else {
      "Percent"
    }
  })
  
  plot_caption <- reactive({
    if (input$plot_caption != "") {
      input$plot_caption
    } else {
      caption()
    }
  })
  
  plot_legend_title <- reactive({
    req(input$analysis_type)
    if (input$plot_legend_title != "") {
      input$plot_legend_title
    } else if (input$analysis_type %in% c("One-way (single-variable)", "Two-way (bi-variable)")) {
      input$dynamic_select_outcome_plot
    } else if (input$analysis_type == "Multivariable") {
      "Response"
    } else {
      NULL
    }
  })
  
  one_way_plot_obj <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_outcome_plot, caption())
    data_survey_srvyr() %>%
      mutate(outcome_lbl = labelled::to_factor(.data[[input$dynamic_select_outcome_plot]])) %>%
      filter(!is.na(outcome_lbl)) %>%
      group_by(outcome_lbl) %>%
      summarise(mean = survey_mean(vartype = "se"), .groups = "drop") %>%
      ggplot(aes(x = outcome_lbl, y = mean, fill = outcome_lbl)) +
      {
        if (input$plot_theme == "nchs_theme")
          geom_bar(stat = "identity", fill = "#006858")
        else
          geom_bar(stat = "identity")
      } +
      {
        if (isTRUE(input$value_labels)) {
          geom_text(aes(label = sprintf("%.1f", mean * 100)), vjust = -0.5, size = 3)
        } else {
          NULL
        }
      } +
      xlab(plot_xlab()) +
      scale_y_continuous(name = plot_ylab(), labels = scales::percent) +
      labs(
        title = plot_title(),
        caption = plot_caption(),
        subtitle = plot_subtitle()
      ) +
      {
        if (input$plot_theme == "nchs_theme")
          nchs_theme
        else
          match.fun(str_sub(input$plot_theme))()
      } +
      theme(legend.position = "none")
  })
  
  two_way_plot_obj <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_outcome_plot, input$dynamic_select_covariate_plot, caption())
    data_survey_srvyr() %>%
      mutate(
        outcome_lbl = labelled::to_factor(.data[[input$dynamic_select_outcome_plot]]),
        covariate_lbl = labelled::to_factor(.data[[input$dynamic_select_covariate_plot]])
      ) %>%
      filter(!is.na(outcome_lbl), !is.na(covariate_lbl)) %>%
      group_by(outcome_lbl, covariate_lbl) %>%
      summarise(mean = survey_mean(vartype = "se"), .groups = "drop") %>%
      ggplot(aes(x = covariate_lbl, y = mean, fill = outcome_lbl)) +
      geom_bar(stat = "identity", position = position_dodge()) +
      {
        if (isTRUE(input$value_labels)) {
          geom_text(
            aes(label = sprintf("%.1f", mean * 100), group = outcome_lbl),
            position = position_dodge(width = 0.9),
            vjust = -0.5,
            size = 3
          )
        } else {
          NULL
        }
      } +
      xlab(plot_xlab()) +
      scale_y_continuous(name = plot_ylab(), labels = scales::percent) +
      labs(
        title = plot_title(),
        caption = plot_caption(),
        subtitle = plot_subtitle(),
        fill = plot_legend_title()
      ) +
      {
        if (input$plot_theme == "nchs_theme")
          nchs_theme
        else
          match.fun(str_sub(input$plot_theme))()
      } +
      {
        if (input$plot_theme == "nchs_theme")
          scale_fill_manual(values = get_nchs_colors(
            length(unique(filtered_data()[[input$dynamic_select_outcome_plot]]))
          ))
        else
          scale_fill_discrete()
      }
  })
  
  multivariable_plot_obj <- reactive({
    req(data_survey_srvyr(), input$dynamic_select_multivariable_plot, plot_caption())
    
    variable_labels <- strsplit(input$multivariable_label_plot, ",")[[1]] %>% trimws()
    variable_names <- input$dynamic_select_multivariable_plot
    variable_label_map <- setNames(
      if (length(variable_labels) == length(variable_names)) variable_labels else variable_names,
      variable_names
    )
    
    data_survey_srvyr <- data_survey_srvyr() %>%
      mutate(across(all_of(input$dynamic_select_multivariable_plot), labelled::to_factor)) %>%
      filter(if_all(all_of(input$dynamic_select_multivariable_plot), ~ !is.na(.)))
    
    input$dynamic_select_multivariable_plot %>%
      map(~ svymean(as.formula(paste0("~", .x)), design = data_survey_srvyr, na.rm = TRUE)) %>%
      set_names(input$dynamic_select_multivariable_plot) %>%
      imap_dfr(~ {
        tibble(
          variable = variable_label_map[.y],
          response = str_remove(names(.x), paste0("^", .y)),
          mean = as.numeric(.x)
          # se = as.numeric(SE(.x))
        )
      }) %>%
      ggplot(aes(x = variable, y = mean, fill = response)) +
      {
        if (!is.null(input$plot_bar_position) && input$plot_bar_position == "Side-by-side (dodged)")
          geom_bar(stat = "identity", position = "dodge")
        else
          geom_bar(stat = "identity", position = "stack")
      } +
      # Data labels
      {
        if (isTRUE(input$value_labels)) {
          if (!is.null(input$plot_bar_position) && input$plot_bar_position == "Side-by-side (dodged)") {
            # Check if axes are flipped
            if (isTRUE(input$plot_axis_flip)) {
              geom_text(
                aes(label = sprintf("%.1f", mean * 100)),
                position = position_dodge(width = 0.9),
                hjust = -0.25, # to the right of the bar for horizontal bars
                size = 3
              )
            } else {
              geom_text(
                aes(label = sprintf("%.1f", mean * 100)),
                position = position_dodge(width = 0.9),
                vjust = -0.5, # above bar for vertical bars
                size = 3
              )
            }
          } else {
            geom_text(
              aes(label = sprintf("%.1f", mean * 100)),
              position = position_stack(vjust = 0.5),
              size = 3
            )
          }
        } else {
          NULL
        }
      } +
      xlab(plot_xlab()) +
      scale_y_continuous(name = plot_ylab(), labels = scales::percent) +
      labs(
        title = plot_title(),
        caption = plot_caption(),
        subtitle = plot_subtitle(),
        fill = plot_legend_title()
      ) +
      {
        if (isTRUE(input$plot_axis_flip))
          coord_flip()
        else
          NULL
      } +
      {
        if (input$plot_theme == "nchs_theme")
          nchs_theme
        else
          match.fun(str_sub(input$plot_theme))()
      } +
      {
        if (input$plot_theme == "nchs_theme")
          scale_fill_manual(values = get_nchs_colors(
            length(unique(input$dynamic_select_multivariable_plot))
          ))
        else
          scale_fill_discrete()
      }
  })
  
  # ---- Report Generation ----
  report_items <- reactiveValues(
    tables = list(),
    plots = list()
  )
  
  report_path <- reactiveVal(NULL)
  
  # =========================
  # Observers and Outputs
  # =========================
  
  # ---- User Guide Modal ----
  observeEvent(input$user_guide, {
    showModal(
      modalDialog(
        title = "SurveyLand User Guide",
        size = "l",
        easyClose = TRUE,
        footer = NULL,
        tags$iframe(
          src = "docs/user-guide.html",
          width = "100%",
          height = "700px",
          style = "border:none;"
        )
      )
    )
  })
  
  # ---- Data Manipulation Observers ----
  observe({
    updateSelectizeInput(session, "filtered_var", choices = variable_name_list())
    updateSelectizeInput(session, "filtered_var2", choices = variable_name_list())
  })
  
  observeEvent(input$filtered_var, {
    req(data(), input$filtered_var)
    updateSelectizeInput(
      session,
      "filtered_var_value",
      choices = sort(unique(data()[[input$filtered_var]]))
    )
  })
  
  observeEvent(input$filtered_var2, {
    req(data(), input$filtered_var2)
    updateSelectizeInput(
      session,
      "filtered_var_value2",
      choices = sort(unique(data()[[input$filtered_var2]]))
    )
  })
  
  # Reactive value to track if the submit button has been pressed
  submit_pressed <- reactiveVal(0)
  
  # Observe the submit button
  observeEvent(input$filter_submit, {
    submit_pressed(submit_pressed() + 1) # Increment the count when the submit button is pressed
  })
  
  # ---- Weighting and Design Observers ----
  observe({
    updateSelectizeInput(session, "dynamic_select_ids", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_strata", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_weight", choices = variable_name_list())
  })
  
  observeEvent(input$design_select, {
    if (input$design_select == "No") {
      updateRadioButtons(session, "weighting_select", selected = "No")
      updateSelectizeInput(session, "dynamic_select_ids", selected = "")
      updateSelectizeInput(session, "dynamic_select_strata", selected = "")
    }
  })
  
  observeEvent(input$weighting_select, {
    if (input$weighting_select == "No") {
      updateSelectizeInput(session, "dynamic_select_weight", selected = "")
    }
  })
  
  # ---- Analysis Observers ----
  
  # ---- Table Generation Observers ----
  observe({
    updateSelectizeInput(session, "dynamic_select_outcome_table", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_covariate_table", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_multivariable_table", choices = variable_name_list())
  })
  
  observeEvent(input$dynamic_select_outcome_table, {
    req(input$dynamic_select_outcome_table)
    updateTextInput(session, "outcome_label_table", value = input$dynamic_select_outcome_table)
  })
  
  observeEvent(input$dynamic_select_covariate_table, {
    req(input$dynamic_select_covariate_table)
    updateTextInput(session, "covariate_label_table", value = input$dynamic_select_covariate_table)
  })
  
  observeEvent(input$dynamic_select_multivariable_table, {
    req(input$dynamic_select_multivariable_table)
    updateTextInput(
      session,
      "multivariable_label_table",
      value = gsub(",", ", ", paste(input$dynamic_select_multivariable_table, collapse = ","))
    )
  })
  
  observe({
    shinyFeedback::feedbackWarning(
      inputId = "dynamic_select_outcome_table",
      show = same_two_way_table_vars(),
      text = "Outcome and covariate cannot be the same variable for a two-way table."
    )
    
    shinyFeedback::feedbackWarning(
      inputId = "dynamic_select_covariate_table",
      show = same_two_way_table_vars(),
      text = "Outcome and covariate cannot be the same variable for a two-way table."
    )
  })
  
  observe({
    req(input$dynamic_select_multivariable_table)
    var_type <- multivariable_var_type()
    mixed_var_type <- var_type == "mixed"
    
    shinyFeedback::feedbackWarning(
      inputId = "dynamic_select_multivariable_table",
      show    = mixed_var_type,
      text    = "Selected variables must all be continuous or all categorical; mixed types are not supported."
    )
  })
  
  # ---- Plot Generation Observers ----
  observeEvent(input$dynamic_select_outcome_plot, {
    req(input$dynamic_select_outcome_plot)
    updateTextInput(session, "outcome_label_plot", value = input$dynamic_select_outcome_plot)
  })
  
  observeEvent(input$dynamic_select_covariate_plot, {
    req(input$dynamic_select_covariate_plot)
    updateTextInput(session, "covariate_label_plot", value = input$dynamic_select_covariate_plot)
  })
  
  observeEvent(input$dynamic_select_multivariable_plot, {
    req(input$dynamic_select_multivariable_plot)
    updateTextInput(
      session,
      "multivariable_label_plot",
      value = gsub(",", ", ", paste(input$dynamic_select_multivariable_plot, collapse = ","))
    )
  })
  
  observe({
    updateSelectizeInput(session, "dynamic_select_outcome_plot", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_covariate_plot", choices = variable_name_list())
    updateSelectizeInput(session, "dynamic_select_multivariable_plot", choices = variable_name_list())
  })
  
  # ---- Report Generation Observers ----
  observeEvent(input$add_one_way_plot_to_report, {
    plt <- one_way_plot_obj()
    label <- paste("Distribution of", input$dynamic_select_outcome_plot)
    report_items$plots <- append(report_items$plots, list(list(label = label, plot = plt)))
    showNotification("Plot will be added to the report.", type = "message")
  })
  
  observeEvent(input$add_two_way_plot_to_report, {
    plt <- two_way_plot_obj()
    label <- paste("Distribution of", input$dynamic_select_outcome_plot, "by", input$dynamic_select_covariate_plot)
    report_items$plots <- append(report_items$plots, list(list(label = label, plot = plt)))
    showNotification("Plot will be added to the report.", type = "message")
  })
  
  observeEvent(input$add_multivariable_plot_to_report, {
    plt <- multivariable_plot_obj()
    label <- paste("Multivariable distributions")
    report_items$plots <- append(report_items$plots, list(list(label = label, plot = plt)))
    showNotification("Plot will be added to the report.", type = "message")
  })
  
  observeEvent(input$generate_report, {
    shinyjs::disable("download_report")
    if (is.null(quarto::quarto_path())) {
      showNotification("Quarto command-line tools path not found.", type = "error")
      return()
    }
    showNotification("Generating report, please wait...", type = "message")
    report_file <- "survey-report.docx"
    report_plots_with_files <- lapply(report_items$plots, function(x) {
      tmpfile <- tempfile(fileext = ".png")
      ggsave(tmpfile, plot = x$plot, width = 8, height = 6, dpi = 600)
      list(label = x$label, file = tmpfile)
    })
    tryCatch({
      quarto::quarto_render(
        "report.qmd",
        execute_params = list(
          survey_name = input$survey_name,
          survey_round = input$survey_round,
          survey_date = input$survey_date,
          report_plots = report_plots_with_files
        ),
        output_file = report_file
      )
      report_path(report_file)
      shinyjs::enable("download_report")
      showNotification("Report generated successfully! Click 'Download report' to get the file.", type = "message")
    }, error = function(e) {
      showNotification(paste("Error generating report:", e$message), type = "error")
    })
  })
  
  # ---- Output Bindings ----
  output$upload_summary <- renderText({ upload_summary() })
  output$surv_meta_summary <- renderText({ surv_meta_summary() })
  
  output$filtered_summary <- renderText({
    req(submit_pressed() > 0)
    filtered_summary()
  })
  
  output$no_filter_message <- renderText({
    if (input$filter_dataset == "No") {
      glue("The entire file with {nrow(filtered_data())} records will be used for analysis.")
    }
  })
  
  output$weighting_design_selections_summary <- renderPrint({
    paste0("design: ", input$design_select, " weighting: ", input$weighting_select)
  })
  
  output$surv_design_summary <- renderText({
    surv_design_summary()
  })
  
  output$no_surv_design_message <- renderText({
    req(input$design_select)
    if (input$design_select == "No") {
      paste(
        "Analysis will be unweighted and design information will not be used.",
        "\n\nAn unweighted and unclustered survey design object has been created using the survey package with the following specifications:",
        "\nids = ~1 (unclustered design)",
        "\nstrata = NULL (no strata specified)",
        "\nweights = ~1 (unweighted analysis)",
        "\nnest = TRUE (nested design)"
      )
    }
  })
  
  output$codebook <- renderDT({
    req(codebook())
    codebook()
  })
  
  output$one_way_table <- renderUI({
    flextable_data_one_way() %>% htmltools_value()
  })
  
  output$two_way_table <- renderUI({
    flextable_data_two_way() %>% htmltools_value()
  })
  
  output$multivariable_table <- renderUI({
    flextable_data_multivariable() %>% htmltools_value()
  })
  
  output$outcome_is_continuous_table <- reactive({
    req(input$dynamic_select_outcome_table, filtered_data())
    outcome_is_continuous_table()
  })
  outputOptions(output, "outcome_is_continuous_table", suspendWhenHidden = FALSE)
  
  output$multivariable_is_continuous <- reactive({
    req(input$dynamic_select_multivariable_table)
    multivariable_var_type() == "continuous"
  })
  outputOptions(output, "multivariable_is_continuous", suspendWhenHidden = FALSE)
  
  output$multivariable_is_mixed <- reactive({
    req(input$dynamic_select_multivariable_table)
    multivariable_var_type() == "mixed"
  })
  outputOptions(output, "multivariable_is_mixed", suspendWhenHidden = FALSE)
  
  output$one_way_plot <- renderPlot({
    one_way_plot_obj()
  })
  
  output$two_way_plot <- renderPlot({
    two_way_plot_obj()
  })
  
  output$multivariable_plot <- renderPlot({
    multivariable_plot_obj()
  })
  
  # ---- Download Handlers ----
  output$download_one_way_table_image <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_table, ".png", sep = "")
    },
    content = function(file) {
      ft <- flextable_data_one_way()
      save_as_image(ft, path = file, dpi = 600)
    }
  )
  
  output$download_two_way_table_image <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_table, "_", input$dynamic_select_covariate_table, ".png", sep = "")
    },
    content = function(file) {
      ft <- flextable_data_two_way()
      save_as_image(ft, path = file, dpi = 600)
    }
  )
  
  output$download_multivariable_table_image <- downloadHandler(
    filename = function() {
      paste("multivariable", ".png", sep = "")
    },
    content = function(file) {
      ft <- flextable_data_multivariable()
      save_as_image(ft, path = file, dpi = 600)
    }
  )
  
  output$download_one_way_table_word <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_table, ".docx", sep = "")
    },
    content = function(file) {
      doc <- read_docx()
      ft <- flextable_data_one_way()
      doc <- body_add_flextable(doc, ft)
      print(doc, target = file)
    }
  )
  
  output$download_two_way_table_word <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_table, "_", input$dynamic_select_covariate_table, ".docx", sep = "")
    },
    content = function(file) {
      doc <- read_docx()
      ft <- flextable_data_two_way()
      doc <- body_add_flextable(doc, ft)
      print(doc, target = file)
    }
  )
  
  output$download_multivariable_table_word <- downloadHandler(
    filename = function() {
      paste("multivariable", ".docx", sep = "")
    },
    content = function(file) {
      doc <- read_docx()
      ft <- flextable_data_multivariable()
      doc <- body_add_flextable(doc, ft)
      print(doc, target = file)
    }
  )
  
  output$download_one_way_table_excel <- downloadHandler(
    filename = function() {
      paste0(input$dynamic_select_outcome_table, ".xlsx")
    },
    content = function(file) {
      # Get processed data and metadata
      processed <- processed_table_one_way()
      subset_table <- processed$table
      title <- processed$title
      footnote_text <- processed$footnote
      
      # Format all numeric columns to 1 decimal place except 'Number'
      subset_table <- subset_table %>%
        mutate(across(where(is.numeric) & !matches("^Number$"), ~ sprintf("%.1f", .)))
      
      # Replace values that begin with -9999 with "*", and NA with a dash
      subset_table <- subset_table %>%
        mutate(across(
          everything(),
          ~ case_when(
            str_detect(as.character(.), "^\\-9999") ~ "*",
            is.na(.) | . == "NA" ~ "---",
            TRUE ~ as.character(.)
          )
        ))
      
      wb <- createWorkbook()
      addWorksheet(wb, "Table")
      n_cols <- ncol(subset_table)
      
      # Title (merged across all columns)
      writeData(wb, "Table", title, startCol = 1, startRow = 1)
      mergeCells(wb, "Table", cols = 1:n_cols, rows = 1)
      
      # Table (header + data)
      writeData(wb, "Table", subset_table, startCol = 1, startRow = 2)
      
      # Footnote (split by lines, merged across all columns)
      footnote_lines <- unlist(strsplit(footnote_text, "\n"))
      footnote_row <- nrow(subset_table) + 3
      writeData(wb, "Table", footnote_lines, startCol = 1, startRow = footnote_row, colNames = FALSE)
      for (i in seq_along(footnote_lines)) {
        mergeCells(wb, "Table", cols = 1:n_cols, rows = footnote_row + i - 1)
      }
      
      # Styles
      baseStyle <- createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000")
      wrapStyle <- createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000", wrapText = TRUE)
      borderStyle <- createStyle(border = "bottom", borderColour = "#000000", borderStyle = "thin")
      
      addStyle(wb, "Table", wrapStyle, rows = 1, cols = 1:n_cols, gridExpand = TRUE)                                      # Title
      addStyle(wb, "Table", wrapStyle, rows = 2, cols = 1:n_cols, gridExpand = TRUE)                                      # Header
      addStyle(wb, "Table", baseStyle, rows = 3:(nrow(subset_table) + 2), cols = 1:n_cols, gridExpand = TRUE)             # Body
      addStyle(wb, "Table", wrapStyle, rows = footnote_row:(footnote_row + length(footnote_lines) - 1), cols = 1:n_cols, gridExpand = TRUE) # Footer
      
      addStyle(wb, "Table", borderStyle, rows = 1:2, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE)                   # Top & header border
      addStyle(wb, "Table", borderStyle, rows = nrow(subset_table) + 2, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE) # Bottom border
      
      setColWidths(wb, "Table", cols = 1:n_cols, widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  output$download_two_way_table_excel <- downloadHandler(
    filename = function() {
      paste0(input$dynamic_select_outcome_table, "_by_", input$dynamic_select_covariate_table, ".xlsx")
    },
    content = function(file) {
      processed <- processed_table_two_way()
      crosstab_table <- processed$table
      title <- processed$title
      footnote_text <- processed$footnote
      
      # Format all numeric columns to 1 decimal place except 'Number'
      crosstab_table <- crosstab_table %>%
        mutate(across(where(is.numeric) & !matches("^Number$"), ~ sprintf("%.1f", .)))
      
      # Replace values that begin with -9999 with "*", and NA with a dash
      crosstab_table <- crosstab_table %>%
        mutate(across(
          everything(),
          ~ case_when(
            str_detect(as.character(.), "^\\-9999") ~ "*",
            is.na(.) | . == "NA" ~ "---",
            TRUE ~ as.character(.)
          )
        ))
      
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "Table")
      n_cols <- ncol(crosstab_table)
      n_rows <- nrow(crosstab_table)
      
      # Prepare two-row header
      col_names <- colnames(crosstab_table)
      header_split <- stringr::str_match(col_names, "^(.*)_(Percent|SE)$")
      header1 <- ifelse(!is.na(header_split[, 2]), header_split[, 2], col_names)
      header2 <- ifelse(!is.na(header_split[, 3]), header_split[, 3], "")
      header1[header2 == ""] <- col_names[header2 == ""]
      
      # Title
      openxlsx::writeData(wb, "Table", title, startCol = 1, startRow = 1)
      openxlsx::mergeCells(wb, "Table", cols = 1:n_cols, rows = 1)
      
      # Two-row header
      openxlsx::writeData(wb, "Table", t(header1), startCol = 1, startRow = 2, colNames = FALSE)
      openxlsx::writeData(wb, "Table", t(header2), startCol = 1, startRow = 3, colNames = FALSE)
      
      # Merge vertically for single-row headers
      for (i in which(header2 == "")) {
        openxlsx::mergeCells(wb, "Table", cols = i, rows = 2:3)
      }
      # Merge horizontally for multi-column groups
      for (grp in unique(header1[header2 != ""])) {
        idx <- which(header1 == grp & header2 != "")
        if (length(idx) > 1) {
          openxlsx::mergeCells(wb, "Table", cols = idx, rows = 2)
        }
      }
      
      # Data
      openxlsx::writeData(wb, "Table", crosstab_table, startCol = 1, startRow = 4, colNames = FALSE)
      
      # Footnote
      footnote_lines <- unlist(strsplit(footnote_text, "\n"))
      footnote_row <- n_rows + 4
      openxlsx::writeData(wb, "Table", footnote_lines, startCol = 1, startRow = footnote_row, colNames = FALSE)
      for (i in seq_along(footnote_lines)) {
        openxlsx::mergeCells(wb, "Table", cols = 1:n_cols, rows = footnote_row + i - 1)
      }
      
      # Styles
      baseStyle <- openxlsx::createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000")
      wrapStyle <- openxlsx::createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000", wrapText = TRUE)
      borderStyle <- openxlsx::createStyle(border = "bottom", borderColour = "#000000", borderStyle = "thin")
      
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = 1, cols = 1:n_cols, gridExpand = TRUE)                                              # Title
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = 2:3, cols = 1:n_cols, gridExpand = TRUE)                                            # Header
      openxlsx::addStyle(wb, "Table", baseStyle, rows = 4:(n_rows + 3), cols = 1:n_cols, gridExpand = TRUE)                                 # Body
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = footnote_row:(footnote_row + length(footnote_lines) - 1), cols = 1:n_cols, gridExpand = TRUE) # Footer
      
      openxlsx::addStyle(wb, "Table", borderStyle, rows = 1:3, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE)                            # Top & header border
      openxlsx::addStyle(wb, "Table", borderStyle, rows = n_rows + 3, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE)                     # Bottom border
      
      openxlsx::setColWidths(wb, "Table", cols = 1:n_cols, widths = "auto")
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  output$download_multivariable_table_excel <- downloadHandler(
    filename = function() {
      paste0("multivariable", ".xlsx")
    },
    content = function(file) {
      processed <- processed_table_multivariable()
      multivariable_table <- processed$table
      title <- processed$title
      footnote_text <- processed$footnote
      
      # Format all numeric columns to 1 decimal place except 'Number'
      multivariable_table <- multivariable_table %>%
        mutate(across(where(is.numeric) & !matches("^Number$"), ~ sprintf("%.1f", .)))
      
      # Replace values that begin with -9999 with "*", and NA with a dash
      multivariable_table <- multivariable_table %>%
        mutate(across(
          everything(),
          ~ case_when(
            str_detect(as.character(.), "^\\-9999") ~ "*",
            is.na(.) | . == "NA" ~ "---",
            TRUE ~ as.character(.)
          )
        ))
      
      wb <- openxlsx::createWorkbook()
      openxlsx::addWorksheet(wb, "Table")
      n_cols <- ncol(multivariable_table)
      n_rows <- nrow(multivariable_table)
      
      # Prepare two-row header
      col_names <- colnames(multivariable_table)
      header_split <- stringr::str_match(col_names, "^(.*)_(Percent|SE)$")
      header1 <- ifelse(!is.na(header_split[, 2]), header_split[, 2], col_names)
      header2 <- ifelse(!is.na(header_split[, 3]), header_split[, 3], "")
      header1[header2 == ""] <- col_names[header2 == ""]
      
      # Title
      openxlsx::writeData(wb, "Table", title, startCol = 1, startRow = 1)
      openxlsx::mergeCells(wb, "Table", cols = 1:n_cols, rows = 1)
      
      # Two-row header
      openxlsx::writeData(wb, "Table", t(header1), startCol = 1, startRow = 2, colNames = FALSE)
      openxlsx::writeData(wb, "Table", t(header2), startCol = 1, startRow = 3, colNames = FALSE)
      
      # Merge vertically for single-row headers
      for (i in which(header2 == "")) {
        openxlsx::mergeCells(wb, "Table", cols = i, rows = 2:3)
      }
      # Merge horizontally for multi-column groups
      for (grp in unique(header1[header2 != ""])) {
        idx <- which(header1 == grp & header2 != "")
        if (length(idx) > 1) {
          openxlsx::mergeCells(wb, "Table", cols = idx, rows = 2)
        }
      }
      
      # Data
      openxlsx::writeData(wb, "Table", multivariable_table, startCol = 1, startRow = 4, colNames = FALSE)
      
      # Footnote
      footnote_lines <- unlist(strsplit(footnote_text, "\n"))
      footnote_row <- n_rows + 4
      openxlsx::writeData(wb, "Table", footnote_lines, startCol = 1, startRow = footnote_row, colNames = FALSE)
      for (i in seq_along(footnote_lines)) {
        openxlsx::mergeCells(wb, "Table", cols = 1:n_cols, rows = footnote_row + i - 1)
      }
      
      # Styles
      baseStyle <- openxlsx::createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000")
      wrapStyle <- openxlsx::createStyle(fontName = "Arial", fontSize = 8, fontColour = "#000000", wrapText = TRUE)
      borderStyle <- openxlsx::createStyle(border = "bottom", borderColour = "#000000", borderStyle = "thin")
      
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = 1, cols = 1:n_cols, gridExpand = TRUE)                                              # Title
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = 2:3, cols = 1:n_cols, gridExpand = TRUE)                                            # Header
      openxlsx::addStyle(wb, "Table", baseStyle, rows = 4:(n_rows + 3), cols = 1:n_cols, gridExpand = TRUE)                                 # Body
      openxlsx::addStyle(wb, "Table", wrapStyle, rows = footnote_row:(footnote_row + length(footnote_lines) - 1), cols = 1:n_cols, gridExpand = TRUE) # Footer
      
      openxlsx::addStyle(wb, "Table", borderStyle, rows = 1:3, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE)                            # Top & header border
      openxlsx::addStyle(wb, "Table", borderStyle, rows = n_rows + 3, cols = 1:n_cols, gridExpand = TRUE, stack = TRUE)                     # Bottom border
      
      openxlsx::setColWidths(wb, "Table", cols = 1:n_cols, widths = "auto")
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  output$download_one_way_plot <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_plot, ".png", sep = "")
    },
    content = function(file) {
      ggsave(file, plot = last_plot(), width = 8, height = 6, dpi = 600)
    }
  )
  
  output$download_two_way_plot <- downloadHandler(
    filename = function() {
      paste(input$dynamic_select_outcome_plot, "_", input$dynamic_select_covariate_plot, ".png", sep = "")
    },
    content = function(file) {
      ggsave(file, plot = last_plot(), width = 8, height = 6, dpi = 600)
    }
  )
  
  output$download_multivariable_plot <- downloadHandler(
    filename = function() {
      paste("multivariable", ".png", sep = "")
    },
    content = function(file) {
      ggsave(file, plot = last_plot(), width = 8, height = 6, dpi = 600)
    }
  )
  
  output$download_report <- downloadHandler(
    filename = function() {
      "survey-report.docx"
    },
    content = function(file) {
      req(report_path())
      file.copy(report_path(), file)
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )
  
  # ---- Output Options ----
  output$file_uploaded <- reactive({
    !is.null(input$upload)
  })
  outputOptions(output, "file_uploaded", suspendWhenHidden = FALSE)
  
}

shinyApp(ui = ui, server = server)