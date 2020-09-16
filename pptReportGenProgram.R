library(tidyverse)
library(leaflet)
library(readxl)
library(magrittr)
library(leafpop)
library(mapview)
library(png)
library(grid)
library(OpenStreetMap)
library(ggrepel)
library(ggpubr)
library(officer)
library(gridExtra)


## function to read in the data from a .xlsx file
func_readFile <- function(fileName){
  data_frame <- readxl::read_xlsx(fileName)
  return(data_frame)
}

## Function to bin the continuous variables into buckets
func_dataTransform <- function(data_frame, nested = T){
  
  ## Change 'Value_1' etc to the relevant variables
  data_frameTransformed <- data_frame %>%
    mutate(Value_1 = if_else(`Value 1` < -85, "-85 and lower",
                             if_else(`Value 1` <= -75, "-75 to -85",
                                     if_else(`Value 1` <= -65, "-65 to -75",
                                             if_else(`Value 1` <= -55, "-55 to -65",
                                                     "-55 and higher"))))) %>%
    mutate(Value_2 = if_else(`Value 2` < -85, "-85 and lower",
                             if_else(`Value 1` <= -75, "-75 to -85",
                                     if_else(`Value 1` <= -65, "-65 to -75",
                                             if_else(`Value 1` <= -55, "-55 to -65",
                                                     "-55 and higher"))))) %>%
    mutate(Value_3 = if_else(`Value 3` < -85, "-85 and lower",
                             if_else(`Value 1` <= -75, "-75 to -85",
                                     if_else(`Value 1` <= -65, "-65 to -75",
                                             if_else(`Value 1` <= -55, "-55 to -65",
                                                     "-55 and higher"))))) %>%
    select(-matches("Value [0-9]+$")) %>%
    pivot_longer(cols = matches("Value_[0-9]+$"), names_to = "specs")
  
  if (nested == T) {
    data_frameNested <- data_frameTransformed %>% group_by(specs, Mobile) %>%
      nest()
    return(data_frameNested)
  }
  return(data_frameTransformed)
}

## Function for creating the base street map

func_baseStreetMap <- function(data_frame){
  lat <- data_frame$Lat
  lon <- data_frame$Lon
  upperLeft = c(max(lat) + 0.001, min(lon) - 0.0001)
  lowerRight = c(min(lat) - 0.001, max(lon) + 0.0001)
  
  streetMap <- openmap(upperLeft = upperLeft, lowerRight = lowerRight,
                       mergeTiles = T, minNumTiles = 8)
  streetMap.latlon <- openproj(streetMap,
                               projection = "+proj=longlat +ellps=WGS84 +datum=WGS84   +no_defs")
  return(streetMap.latlon)
}

## Function for creating the plots : inputs -> data_frame, baseStreet map

func_plotGen <- function(data_frame, streetMap, ttheme = "minimal", plotTitle){
  
  ##Street Map
  
  streetMapPlot <- autoplot.OpenStreetMap(streetMap) + 
    geom_point(data = data_frame, aes(x = Lon, y = Lat, fill = value),
               size = 3, show.legend = F, alpha = 0.3, shape = 21) +
    theme_void() #+ 
    #theme(plot.margin = margin(c(1,1,1,1, unit = "pt")))
  
  ##Pie chart
  
  
  dataSumm <-  data_frame %>% 
    select(value) %>% group_by(value) %>% na.omit() %>% summarize(counts = n())  %>%
    ungroup() %>% 
    mutate(pct = counts*100/sum(counts)) %>% 
    # mutate(colors = if_else(specs == "-55 and higher", "#32a852",
    #                                     if_else(RSSI == "-55 to -65", "#25dbd2",
    #                                     if_else(RSSI == "-65 to -75", "#f2c613",
    #                                     if_else(RSSI == "-75 to -85", "#e61938",
    #                                             "#0f0f0f"))))) %>% 
    arrange((counts)) %>%
    mutate(cumVal = cumsum(counts),
           midPoint = cumVal - counts/2,
           specs = as.factor(value))
  
  
  
  donutPlot1 <- dataSumm %>%
    ggplot(aes(x = 1, y = counts, fill = fct_inorder(value))) +
    geom_bar(stat = "identity", position = "stack", width = 1, show.legend = F) +
    coord_polar(theta = "y", start = 0) +
    geom_text_repel(aes(label = paste0(round(pct, 2), "%"), y = sum(counts) - midPoint),
                    col = "black", size = 3.5,
                    show.legend = F, nudge_x = 2.5, segment.size = 0.2,
                    direction = "y", hjust = 0.1) +
    theme_void() +
    theme(plot.margin = margin(0,0,0,0, unit = "pt"),
          text = element_text(size = 100))
  ## Datatable Grob
  
  
  
  if (ttheme == "default") {
    tableGrobVal <- tableGrob(dataSumm %>%
                                select(value, counts), rows = NULL,
                              theme = ttheme_default(base_size = 11,
                                                     padding = unit(c(2,2), "mm")))
  } else {
    tableGrobVal <- tableGrob(dataSumm %>%
                                select(value, counts), rows = NULL,
                              theme = ttheme_minimal(base_size = 10,
                                                     padding = unit(c(1,1), "mm")))
    tableGrobVal <- gtable::gtable_add_grob(tableGrobVal,
                                            grobs = grid::rectGrob(gp =
                                                                     grid::gpar(fill = NA,
                                                                                lwd = 2)),
                                            t = 2, b = nrow(tableGrobVal), l = 1,
                                            r = ncol(tableGrobVal))
    tableGrobVal <- gtable::gtable_add_grob(tableGrobVal,
                                            grobs = grid::rectGrob(gp = 
                                                                     grid::gpar(fill = NA,
                                                                                lwd = 2)),
                                            t = 1, l = 1, r = ncol(tableGrobVal))
  }
  
  tableGrobVal %<>%  as_ggplot()
  
  # dataSummGrob <- dataSumm %>% select(value, counts) %>% gt() %>%
  #   as_ggplot()
  
  # layoutMatrix <- rbind(c(1,2),c(3,2))
  # arranged_gridPLot <- grid.arrange(tableGrobVal, streetMapPlot, donutPlot1,
  #                                   ncol = 2, top = plotTitle,
  #                                   layout_matrix = layoutMatrix,
  #                                   widths = c(1,3), heights = c(1,1),
  #                                   padding = unit(0, "line")) %>%
  #   ggpubr::as_ggplot()
  
  arranged_gridPLot <- ggarrange(ggarrange(tableGrobVal, donutPlot1, ncol = 1),
                                 ggarrange(streetMapPlot, ncol = 1, nrow = 1),
                                 ncol = 2, align = "hv", labels = plotTitle,
                                 widths = c(1,2), common.legend = F)
  
  return(arranged_gridPLot)
}


## Function for aggregating 4 plots in 1

func_plotAgg <- function(plotList){
  ggarrange(plotlist = plotList$plots, ncol = 2, nrow = 2, align = "hv") %>%
    return()
}

## Function for creating PPT slides
func_createSlides <- function(plots, ppt){
  ppt <- add_slide(ppt, layout = "Blank", master = "Office Theme")
  ppt <- ph_with(ppt, value = plots, ph_location_fullsize())
  return(ppt)
}

## Function for execution

func_main <- function(outputPPTName){
  ## Read in the data frame
  filePath <- file.choose()
  data_frame <- func_readFile(filePath)
  
  ## Transform and nest the data : nests by the variable specs
  nestedData <- func_dataTransform(data_frame, nested = T)
  
  ## Create Base street map : common to all plots
  baseMap <- func_baseStreetMap(data_frame)
  
  ## Create plots for nested data
  nestedData %<>% mutate(plots = map(.x = data, .f = func_plotGen, baseMap,
                                     plotTitle = Mobile))
  
  ## Aggregate nested data another level
  nestedData2 <- nestedData %>% group_by(specs) %>% nest()
  
  ## Plot aggregation for 4 plots
  nestedData2 %<>% mutate(plots4 = map(.x = data, .f = func_plotAgg))
  
  ## Create PPT and print plots
  ppt <- read_pptx()
  pptOutput <- lapply(nestedData2$plots4, func_createSlides, ppt)
  
  ## Save PPT
  print(pptOutput, target = outputPPTName)
}