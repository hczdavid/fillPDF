

#---------------------------------------------------------------------------#
#                                                                           #
#  Please Click "Run App" button in the top right to go the interface       #
#                                                                           #
#---------------------------------------------------------------------------#



#-------------Don't Change anything below. Thank you!-----------------------#

library(shiny)
library(xlsx)
library(staplr)


source(paste0(getwd(),"/Source/generatePDF.R"))

# Define UI for application that draws a histogram
ui <- fluidPage(

    # Application title
    titlePanel("Generate Forms"),

    # Sidebar with a slider input for number of bins 
    sidebarLayout(
        sidebarPanel(
            # Input: Select a file 
            fileInput("audittool", "Choose an Audit Tool (.xlsx file)",
                      multiple = FALSE,
                      accept = c("xlsx")),
            
            #selectInput("cyl", "Choose the cycle", choices = c("Cycle2", "Cycle3")),
            selectInput("act_neg", "Choose the sheet", choices = c("Active", "Negative","Inquiry")),
            #submitButton(text = "Update individual"),

            tags$hr(),
            selectInput("ind", label = "Choose individual(s)", NULL,multiple = T),# multiple choices
            tags$h6(),
            textInput("mypath", "Output Path", value = "Default"),
            tags$hr("Click ONCE to generate files"),
            #submitButton("Generate",width = 150)
            actionButton("generate", "Generate", width = 200),
            
        ),

        # Show a plot of the generated distribution
        mainPanel(
            tags$h2("Welcome to Use This Tool"),
            tags$h3("Progress Log"),
            textOutput("wd"),# the directory appears on the main panel
            tags$hr(),
            textOutput("upfile"),
            tags$hr(),
            textOutput("prog1"),
            
            textOutput("prog2"),
            tags$hr(),
            textOutput("prog3"),
            tags$hr(),
            span(textOutput("message"), style="color:red"),# color words
            tags$hr(),
            textOutput("outpath")
            
            
            
        )
    )
)

# Define server logic required to draw a histogram
server <- function(input, output,session) {
    
    
    act_neg <- reactive({
        if(input$act_neg == "Active")return("QA ACTIVE REVIEW")
        else if(input$act_neg == "Negative")return("QA NEGATIVE REVIEW")
        return("QA INQUIRY REVIEW")
        })
    
    myind <- reactive({
        if(input$act_neg == "Active")return(c(2, 5:100))
        else if(input$act_neg == "Negative")return(c(2, 4:100))
        return(c(3,5:100))
    })
    
    data <- reactive({
        inFile <- input$audittool
        if (is.null(inFile)) return(NULL)
        xlsx::read.xlsx(inFile$datapath,sheetName = act_neg(), rowIndex = myind())
    })
    
    mypath <- reactive({
        if(input$mypath == "Default")return(NULL)
        input$mypath
    })
    
    #output$outpath <- renderText({getwd()})
    
    output$upfile <- renderText({
        "1. Please upload an Audit tool (.xlsx) file."})
    
   
    output$prog1 <- renderText({
        req(input$audittool)
        "2. Please Select the sheet and individual."})
    
    output$prog2 <- renderText({
        req(input$ind)
        paste("You have selected", input$ind, "from", input$act_neg,"case.")})
    
    output$prog3 <- renderText({
        req(input$ind)
        "3. It's ready to generate documents. Please click 'GENERATE' only ONCE."
    })
    
    mychoice <- reactive(data()[!is.na(data()[,3]), 3])
    
    output$wd <- renderText({
        mydesk <- strsplit(getwd(), split = "[/]")[[1]]
        mydesk <- mydesk[-length(mydesk)]
        paste0("My Desktop Path:  ", paste(mydesk, collapse = "/"))})

     observeEvent(data(), {
         updateSelectInput(session, "ind", choices = mychoice())
     })
    
     
     message <- eventReactive(input$generate, {
         generatePDF(mydata = data(), act_neg = act_neg(), individual = input$ind, mypath = mypath())
     })
    
    output$message <- renderText({
        message()
    })
    
    
}

# Run the application 
shinyApp(ui = ui, server = server)
                                                                                                                                                                         