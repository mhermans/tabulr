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

# Table 1: cbind()'ed table
tab1 <- as.data.frame(t.wg.centrale) # omzetten naar dataframe
#class(tab1[,3] ) <- "percentage"

# Table 2: svytable()
tab2 <- as.data.frame.matrix(t.hhr.centrale)

# Table 3: result svy.datatable()
tab3 <- t.invoed.tech


# write single table in single sheet
# ----------------------------------

wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = 'OR_tables')
setColWidths(wb, sheet = 1, cols=1:30, widths = "auto") # set cols to automatically resize

wb <- print.tabular.xlsx(wb, sheet=1, coords=c(1,1), tabular=tab1, 
                         add.caption=TRUE,
                         add.row.margin=TRUE,
                         add.col.margin=TRUE)

wb <- print.tabular.xlsx(wb, sheet=1, coords=c(1,1), tabular=tab1, 
                         add.caption=TRUE,
                         add.row.margin=FALSE)
openXL(wb)

# write multiple tables in single sheet
# -------------------------------------

wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = 'OR_tables')
setColWidths(wb, sheet = 1, cols=1:30, widths = "auto") # set cols to automatically resize

wb <- print.tabulars.xlsx(wb, 1, list(tab1, tab2, tab3))
openXL(wb) # toon excel bestand zonder weg te schrijven

saveWorkbook(wb, "OR1_descr_tables.xlsx", overwrite = TRUE)



# writeData(wb = wb, sheet = sheet_name, x = tab, xy=coords,
#           borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")
# writeData(wb = wb, 
#           sheet = sheet_name, x = tab, 
#           xy=coords,
#           borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")



# VOORBEELDTABELLEN
# =================

wb <- createWorkbook()


# Tabel met enkel data en row/colnames
addWorksheet(wb = wb, sheetName = 'basic')
setColWidths(wb, sheet = 1, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=1, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=FALSE,
                         add.row.margin=FALSE,
                         add.col.margin=FALSE)

# Tabel met data, R/C names en caption
addWorksheet(wb = wb, sheetName = 'caption')
setColWidths(wb, sheet = 2, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=2, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=TRUE,
                         add.row.margin=FALSE,
                         add.col.margin=FALSE)

# Tabel met data, R/C names, caption en R margins
addWorksheet(wb = wb, sheetName = 'row_margin')
setColWidths(wb, sheet = 3, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=3, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=TRUE,
                         add.row.margin=TRUE,
                         add.col.margin=FALSE)

# Tabel met data, R/C names, caption en col margins
addWorksheet(wb = wb, sheetName = 'c_margins')
setColWidths(wb, sheet = 4, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=4, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=TRUE,
                         add.row.margin=FALSE,
                         add.col.margin=TRUE)

# Tabel met data, R/C names, caption en R/C margins
addWorksheet(wb = wb, sheetName = 'rc_margins')
setColWidths(wb, sheet = 5, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=5, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=TRUE,
                         add.row.margin=TRUE,
                         add.col.margin=TRUE)

# Tabel met data, R/C names en caption en comment
addWorksheet(wb = wb, sheetName = 'comment')
setColWidths(wb, sheet = 6, cols=1:30, widths = "auto")
wb <- print.tabular.xlsx(wb, sheet=6, coords=c(1,1), tabular=t.invoed.tech, 
                         add.caption=TRUE,
                         add.comment=TRUE,
                         add.row.margin=FALSE,
                         add.col.margin=FALSE)

# Tabel met data, R/C names, caption en R/C margins en comment
addWorksheet(wb = wb, sheetName = 'OR_tables')
setColWidths(wb, sheet = 1, cols=1:30, widths = "auto")

openXL(wb)
