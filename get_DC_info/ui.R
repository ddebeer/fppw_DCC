# This shiny app is made to extract the relevant information from an excel-
# document related to the starting PhDs at the Ghent University FPPW.
 

library(shiny)

# Define UI for application that draws a histogram
ui <- basicPage(
  fileInput(
    inputId = "file",
    label = "Upload the doctoral information file",
    multiple = FALSE,
    accept = ".xlsx",
    width = NULL,
    buttonLabel = "Browse to the file",
    placeholder = "<file name>"),
  downloadLink("doctoral_info.docx", "Download .docx"),
  htmlOutput("text")
  )
