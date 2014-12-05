wd <- getwd()
setwd('C:/Users/MaartenH/Documents/work/HIVA/acv_or')
load('rapportering/descr_tables.RData')
setwd(wd)

library(openxlsx)



#write.xlsx(t.wg.centrale, file = "writeXLSXTable1.xlsx", asTable = TRUE) # error, geen data.frame

# options: simple data, or automatically formatedd as a table
# write.xlsx(tab, file = "OR1_descr_tables.xlsx", asTable = TRUE, # issue, "row.name"s als colname
#            col.names=TRUE, row.names=TRUE) 


# General styling
# ---------------

# make custom table header style
hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", 
                   textDecoration = "Italic", border = "Bottom")

# stel algemene stijlelementen in
options("openxlsx.borderColour" = "#4F80BD")
options("openxlsx.borderStyle" = "thin")


# Setup workbook, sheet
# ---------------------

wb <- createWorkbook()
sheet_name <- "OR_tables"
addWorksheet(wb = wb, sheetName = sheet_name)
setColWidths(wb, sheet = 1, cols=1:100, widths = "auto") # set cols to automatically resize

# write multiple tables in single sheet
# -------------------------------------

# Table 1: cbind()'ed table

tab <- as.data.frame(t.wg.centrale) # omzetten naar dataframe
class(tab[,3] ) <- "percentage"
  
coords <- c(1,1)
writeData(wb = wb, sheet = sheet_name, x = tab, xy=coords,
          borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")

# Table 2: svytable()

tab <- as.data.frame.matrix(t.hhr.centrale)

coords <- c(1,nrow(tab) + 3) # hoogte tab 1 + 3 rows => 1 spatie
writeData(wb = wb, 
          sheet = sheet_name, x = tab, 
          xy=coords,
          borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")


openXL(wb) # toon excel bestand zonder weg te schrijven

saveWorkbook(wb, "OR1_descr_tables.xlsx", overwrite = TRUE)

# TODO wat is een header/footer?


