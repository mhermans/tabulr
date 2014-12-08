source('table_export_openxlsx.r')


# BASIC EXAMPLES #
# ============== #

d <- as.data.frame(Titanic)
tab1 <- table(d$Class, d$Sex)
tab2 <- table(d$Class, d$Survived)
tab3 <- table(d$Sex, d$Survived)


# write single table
# ------------------

wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = 'OR_tables')

wb <- print.tabular.xlsx(wb, 1, c(1,1), tab1)
openXL(wb)


# write multiple tables
# ---------------------

wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = 'OR_tables')

wb <- print.tabulars.xlsx(wb, 1, list(tab1) )
#wb <- print.tabulars.xlsx(wb, 1, list(tab1, tab2, tab3) )
# Error in .self$updateCellStyles(sheet = r$sheet, rows = r$rows, cols = r$cols,  : 
#                                   CHAR() can only be applied to a 'CHARSXP', not a 'NULL'
openXL(wb)




# EXAMPLES: OR TABELLEN # 
# ===================== #

wd <- getwd()
setwd('C:/Users/MaartenH/Documents/work/HIVA/acv_or')
load('rapportering/descr_tables.RData')
setwd(wd)
rm(wd)


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
addWorksheet(wb = wb, sheetName = 'OR_tables')
setColWidths(wb, sheet = 1, cols=1:30, widths = "auto") # set cols to automatically resize

# write multiple tables in single sheet
# -------------------------------------

# Table 1: cbind()'ed table

tab1 <- as.data.frame(t.wg.centrale) # omzetten naar dataframe
#class(tab1[,3] ) <- "percentage"

# Table 2: svytable()
tab2 <- as.data.frame.matrix(t.hhr.centrale)

wb <- print.tabulars.xlsx(wb, 1, list(tab1, tab2))
openXL(wb) # toon excel bestand zonder weg te schrijven

saveWorkbook(wb, "OR1_descr_tables.xlsx", overwrite = TRUE)



# writeData(wb = wb, sheet = sheet_name, x = tab, xy=coords,
#           borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")
# writeData(wb = wb, 
#           sheet = sheet_name, x = tab, 
#           xy=coords,
#           borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")

