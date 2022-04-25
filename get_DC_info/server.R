# This shiny app is made to extract the relevant information from an excel-
# document related to the starting PhDs at the Ghent University FPPW.
 

library(shiny)
library(officer)
source("utils.R")

# Define server logic required to draw a histogram
server <- function(input, output) {
  
    data <- reactive({
      # read-file
      file <- input$file
      ext <- tools::file_ext(file$datapath)
      
      req(file)
      validate(need(ext == "xlsx", "Please upload a .xlsx file"))
      
      data <- readxl::read_xlsx(file$datapath, skip = 6)
      # data <- readxl::read_xlsx(file, skip = 6)
      
      # rename cols
      data <- rename_cols(data)
      
      # preprocess data
      preprocess_data(data)
    })

    output$text <- renderText({
        extract_info(data(), type = "html")
    })
    
    output$doctoral_info.docx <- downloadHandler(
      filename = function() {
        paste0("doctoral_info_", Sys.Date(), '.docx')
        }, 
      content = function(con) {
        out <- extract_info(data(), type = ".docx")
        print(out, target = con)
        }
      )
}

