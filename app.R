library(shiny)
library(png)
library(openxlsx)

# Define UI for application that draws a histogram
ui <- fluidPage(
  titlePanel("Your photo in an Excel sheet!"),
   sidebarLayout(
      sidebarPanel(
        fileInput("file1", "Choose a PNG file", multiple=FALSE, accept=c("image/png", ".png")),
        fluidRow(
          column(6, align="left", downloadButton("downloadData", "Download")),
          column(6, align="right", tags$a(href="https://github.com/ssayols/png2xlsx", "Source here!"))
        )
      ),
      mainPanel(
        plotOutput("raster")
      )
   )
)

# Define server logic required to draw a histogram
server <- function(input, output) {
   
  # read png image and convert to grayscale
  img <- reactive({
    req(input$file1)
    img <- readPNG(input$file1$datapath)
    img.gray <- matrix(nrow=nrow(img), ncol=ncol(img))
    for(i in 1:nrow(img)) {
      for(j in 1:ncol(img)) {
        img.gray[i, j] <- mean(img[i, j, ])
      }
    }
    img.gray
  })
  
  # display image on the screen
  output$raster <- renderPlot({
    plot.new()
    rasterImage(img(), 0, 0, 1, 1)
  })
  
  # download image, after converting to an excel workbook with proper background formatting
  output$downloadData <- downloadHandler(
    filename= "yourphoto.xlsx",
    content = function(file) {
      wb <- createWorkbook()
      modifyBaseFont(wb, fontSize=1, fontColour="white", fontName="Calibri")
      addWorksheet(wb, "img", gridLines=FALSE, zoom=50)
      writeData(wb, sheet=1, x=img())
      setColWidths(wb, sheet=1, cols=1:ncol(img()), widths=1)
      setRowHeights(wb, sheet=1, rows=1:nrow(img()), heights=1)
      conditionalFormatting(wb, sheet=1, cols=1:ncol(img()), rows=1:nrow(img()),
                            style=c("black", "white"),
                            rule =c(0, 1),
                            type ="colourScale")
      saveWorkbook(wb, file=file, overwrite=TRUE)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)

