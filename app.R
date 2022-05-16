rm(list = ls())

cat("\f") 

packages <- c("tidyverse", "readxl", "data.table", "utils", "writexl", "openxlsx", "Rcpp", "officer", "kableExtra", "rmdformats", "flexdashboard", "shiny", "shinyWidgets", "shinydashboard", "shinyWidgets", "shinythemes")    

for (i in 1:length(packages)) {
  if (!packages[i] %in% rownames(installed.packages())) {
    install.packages(packages[i])
  }
  library(packages[i], character.only = TRUE)
}
rm(packages)

the_dir <- "C:/Users/keoha/OneDrive/Documents/BC Graduate Assistant/WC_Dashboard" #set this to the location of the relevant data sheet
setwd(the_dir)

read_excel_allsheets <- function(filename, tibble = FALSE) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  names(x) <- sheets
  x
}

mysheets <- read_excel_allsheets("WC Historic Stats.xlsx")
mysheets <- purrr::map(mysheets, tibble::as_tibble)
list2env(mysheets, envir = .GlobalEnv)


#Stdnt.Body.Comp.Total <- Stdnt.Body.Comp.FY[30,]
#Stdnt.Body.Comp.FY <- Stdnt.Body.Comp.FY[1:29,]
Stdnt.Body.Comp.FY$`Fiscal Year` <- as.integer(Stdnt.Body.Comp.FY$`Fiscal Year`)
WC.Course.Format.Offers.Class$`Fiscal Year` <- as.integer(WC.Course.Format.Offers.Class$`Fiscal Year`)
Cert.Stdnt.Body.Comp$`Fiscal Year` <- as.integer(Cert.Stdnt.Body.Comp$`Fiscal Year`)
Grad.Stdnt.Body.Comp$`Fiscal Year` <- as.integer(Grad.Stdnt.Body.Comp$`Fiscal Year`)
Undrgrd.Stdnt.Body.Comp$`Fiscal Year` <- as.integer(Undrgrd.Stdnt.Body.Comp$`Fiscal Year`)
Grad.Stdnt.Body.Comp.SUMM.Only$`Fiscal Year` <- as.integer(Grad.Stdnt.Body.Comp.SUMM.Only$`Fiscal Year`)
Undrgrd.Stdnt.Body.Comp.SUMM$`Fiscal Year` <- as.integer(Undrgrd.Stdnt.Body.Comp.SUMM$`Fiscal Year`)
Cert.Stdnt.1.Course.Format$`Fiscal Year` <- as.integer(Cert.Stdnt.1.Course.Format$`Fiscal Year`)
Grad.Stdnt.1.Course.Format$`Fiscal Year` <- as.integer(Grad.Stdnt.1.Course.Format$`Fiscal Year`)
Undrgrd.Stdnt.1.Course.Format$`Fiscal Year` <- as.integer(Undrgrd.Stdnt.1.Course.Format$`Fiscal Year`)
Undrgrd.Stdnt.Body.Comp.fy$`Fiscal Year` <- as.integer(Undrgrd.Stdnt.Body.Comp.fy$`Fiscal Year`)
Grad.Stdnt.Body.Comp.fy$`Fiscal Year` <- as.integer(Grad.Stdnt.Body.Comp.fy$`Fiscal Year`)
Cert.Stdnt.Body.Comp.fy$`Fiscal Year` <- as.integer(Cert.Stdnt.Body.Comp.fy$`Fiscal Year`)


#Base R solution to formatting elements of a list into seperate *tibbles*--------------------------------------------------------------------------------------------------------
#new_list <- setNames(Map(function(x, y) setNames(data.frame(x), y), 
#                         list_a,names(list_a)), paste0("df", 1:3))
#--------------------------------------------------------------------------------------------------------------------
#----------------------------Shiny--------------------------------------------------------------------------------------------------
library(shiny)
# Define UI ----
ui <- fluidPage(theme = bslib::bs_theme(bootswatch = "journal", bg = "#4D4C4C"
                                        , fg = "white", primary = "#962232"
                                        , secondary = "#DBCCA4", heading_font = "Scala")
                 #, titlePanel("Woods College of Advancing Studies Program Statistics")
       , fluidRow(
           mainPanel(
             img(src = "WC_Logo.png", height = 85, width = 573, align = 'center')
       , navlistPanel(
         column(width = 12
                , sliderInput("slider1", label = "Fiscal Years"
                              , min = min(Stdnt.Body.Comp.FY$`Fiscal Year`)
                              , max = max(Stdnt.Body.Comp.FY$`Fiscal Year`)
                              , value = c(min(Stdnt.Body.Comp.FY$`Fiscal Year`), max(Stdnt.Body.Comp.FY$`Fiscal Year`))
                              , sep = "", step = 1, ticks = F)
         ),
        column( width = 12, pickerInput("var1", "Certificate Program", choices = unique(Cert.Stdnt.Body.Comp$Program)
                                                   , options = list(`action-box` = T)
                                                   , multiple = T
                                                   , selected = unique(Cert.Stdnt.Body.Comp$Program))), 
        column(width = 12, pickerInput("var2", "Graduate Major", choices = unique(Grad.Stdnt.Body.Comp$Major)
                                                   , options = list(`action-box` = T)
                                                   , multiple = T
                                                   , selected = unique(Grad.Stdnt.Body.Comp$Major))), 
        column(width = 12, pickerInput("var3", "Undergraduate Major", choices = unique(Undrgrd.Stdnt.Body.Comp$`Major`)
                                                   , options = list(`action-box` = T)
                                                   , multiple = T
                                                   , selected = unique(Undrgrd.Stdnt.Body.Comp$`Major`))),
        "Woods College Aggregate",
        "Tables", 
        tabPanel("Woods College Students Served", DT::dataTableOutput("table")), 
        tabPanel("Woods College Student Body Composition FY", DT::dataTableOutput("mytable1")
                , selected = T),
        tabPanel("Woods College Course Format Offerings", DT::dataTableOutput("mytable2")),
        tabPanel("Woods College Course Format Offerings by Class", DT::dataTableOutput("mytable3")),
        "Student Body Composition",
        tabPanel("Certificate Student Body Composition", DT::dataTableOutput("mytable4")),
        tabPanel("Certificate Student Body Composition FY", DT::dataTableOutput("mytable5")),
        tabPanel("Graduate Student Body Composition", DT::dataTableOutput("mytable6")),
        tabPanel("Graduate Student Body Composition FY", DT::dataTableOutput("mytable7")),
        tabPanel("Undergraduate Student Body Composition", DT::dataTableOutput("mytable8")),
        tabPanel("Undergraduate Student Body Composition FY", DT::dataTableOutput("mytable9")),
        tabPanel("Graduate Summer Only Student Body Composition", DT::dataTableOutput("mytable10")),
        tabPanel("Undergraduate Summer Only Student Body Composition", DT::dataTableOutput("mytable11")),
        "Students Enrolled in Only",
        "One Course Format", 
        tabPanel("Certificate Students Only Enrolled in 1 Course Format", DT::dataTableOutput("mytable12")),
        tabPanel("Graduate Students Only Enrolled in 1 Course Format", DT::dataTableOutput("mytable13")),
        tabPanel("Undergraduate Student Only Enrolled in 1 Course Format", DT::dataTableOutput("mytable14"))
      )
      
      #column(width = 7
      #       , pickerInput("var1", "Certificate Program", choices = unique(Cert.Stdnt.Body.Comp$Major)
      #                     , options = list(`action-box` = T), multiple = T)
      #       , pickerInput("var2", "Graduate Major", choices = unique(Grad.Stdnt.Body.Comp$Major)
      #                     , options = list(`action-box` = T), multiple = T)
      #       , pickerInput("var3", "Undergraduate Major", choices = unique(Undrgrd.Stdnt.Body.Comp$`Major 1`)
      #                     , options = list(`action-box` = T), multiple = T)
      #)
   )
  )
 )

# Define server logic ----
server <- function(input, output) {
  
  output$range <- renderPrint({ input$slider1 })
  
  observe({
    print(input$var1)
  })
  observe({
    print(input$var2)
  })
  observe({
    print(input$var3)
  })
  # choosing tables to display
  output$table <- DT::renderDataTable({
    DT::datatable(Stdnts.Served, rownames = F)
  })
  output$mytable1 <- DT::renderDataTable({
    DT::datatable(Stdnt.Body.Comp.FY %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F)
  })
  output$mytable2 <- DT::renderDataTable({
    DT::datatable(WC.Course.Format.Offers %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F)
  })
  output$mytable3 <- DT::renderDataTable({
    DT::datatable(WC.Course.Format.Offers.Class %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F)
  })
  output$mytable4 <- DT::renderDataTable({
    DT::datatable(Cert.Stdnt.Body.Comp %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
      dplyr::filter(`Program` %in% input$var1), rownames = F)
    
  })
  output$mytable5 <- DT::renderDataTable({
    DT::datatable(Cert.Stdnt.Body.Comp.fy %>%
                    dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
                    dplyr::filter(`Program` %in% input$var2), rownames = F)
  })
  output$mytable6 <- DT::renderDataTable({
    DT::datatable(Grad.Stdnt.Body.Comp %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
      dplyr::filter(`Major` %in% input$var2), rownames = F)
  })
  output$mytable7 <- DT::renderDataTable({
    DT::datatable(Grad.Stdnt.Body.Comp.fy %>%
                    dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
                    dplyr::filter(`Major` %in% input$var2), rownames = F)
  })
  output$mytable8 <- DT::renderDataTable({
    DT::datatable(Undrgrd.Stdnt.Body.Comp %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
      dplyr::filter(`Major` %in% input$var3), rownames = F)
  })
  output$mytable9 <- DT::renderDataTable({
    DT::datatable(Undrgrd.Stdnt.Body.Comp.fy %>%
                    dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
                    dplyr::filter(`Major` %in% input$var3), rownames = F)
  })
  output$mytable10 <- DT::renderDataTable({
    DT::datatable(Grad.Stdnt.Body.Comp.SUMM.Only %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]) %>%
      dplyr::filter(`Major` %in% input$var2), rownames = F)
  })
  output$mytable11 <- DT::renderDataTable({
    DT::datatable(Undrgrd.Stdnt.Body.Comp.SUMM %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2])%>%
      dplyr::filter(`Major` %in% input$var3), rownames = F)
  })
  output$mytable12 <- DT::renderDataTable({
    DT::datatable(Cert.Stdnt.1.Course.Format %>%
      dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F)
  })
  output$mytable13 <- DT::renderDataTable({
    DT::datatable(Grad.Stdnt.1.Course.Format %>%
                    dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F)
  })
  output$mytable14 <- DT::renderDataTable({
    DT::datatable(Undrgrd.Stdnt.1.Course.Format %>%
                    dplyr::filter(`Fiscal Year` >= input$slider1[1], `Fiscal Year` <= input$slider1[2]), rownames = F) 
                     
  })
  
}
# Run the app ----
shinyApp(ui, server)


