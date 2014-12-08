library(openxlsx)

wd <- getwd()
setwd('C:/Users/MaartenH/Documents/work/HIVA/acv_or')
load('rapportering/descr_tables.RData')
setwd(wd)
rm(wd)

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
sheet_name <- 'OR_tables'
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



# basis-elementen die rondom de data geplaatst worden
#
# * caption   	  # col onder of boven toevoegen, spanning?
# * row.margin 		# col rechts toevoegen
# * col.margin 		# row onder toevoegen
# * comment			  # row onder toevoegen [italic?)

print.tabular.xlsx <- function(wb, sheet, coords, tabular, 
                               add.caption=FALSE, add.row.margin=FALSE, add.col.margin=FALSE, add.comment=FALSE, 
                               style=None) {
  
  # make sure the object is a data.frame, convert if needed
  # -------------------------------------------------------
  
  if ('matrix' %in% class(tabular) ) { tabular <- as.data.frame(tabular) }
  if ('svytable' %in% class(tabular)) { tabular <- as.data.frame.matrix(tabular) }
  
  # consider coords (r,c)
  start_r <- coords[2]
  start_c <- coords[1]
  
  n_data_rows <- nrow(tabular)
  n_data_cols <- ncol(tabular)
  #n_total_cols <- 
  #n_total_rows <- 

  # if caption, start the table one row lower
  if (add.caption) { start_r <- start_r + 1 }
  
  # write out data rows/cols, including row & col names
  # ---------------------------------------------------
  
  writeData(
    wb = wb, sheet = sheet, startCol = start_c, startRow = start_r,
    x = tabular, rowNames=TRUE, colNames=TRUE,
    borders = "surrounding")
  
  
  # add caption-row and caption (on top by default)
  # -----------------------------------------------
  
  if (add.caption) {
    caption_text <- 'Table 1: This is a very long static table caption that needs to be fixed with a variable input (in %)'
    
    # add row contents
    # ----------------
    
    # row for comment is one above start_r 
    comment_r <- start_r - 1
    writeData(
      wb = wb, sheet = sheet, startCol = start_c, startRow = comment_r,
      x = caption_text, rowNames=FALSE, colNames=FALSE)
    
    # for caption, merge the cells to span multiple cols, either width table, or 
    # some minimum nr. of cols
    
    comment_merge_width <- ifelse(7 - start_c >= 6, 6)
    mergeCells(wb, sheet=sheet, 
               cols = start_c:comment_merge_width, 
               rows = comment_r)
    
    # TODO: auto col width should not be affected
    # cf. https://github.com/awalker89/openxlsx/issues/43
    
  }
  

  
  
  # if margins, add row/and or col margin
  # -------------------------------------
  
  row.margin.name
  col.margin.name
  
  margin.table(tab)
    
  # if comment, add additional row with comment
  # -------------------------------------------
  
    
  
  # return wb (save if filename?)
  # -----------------------------
  wb
  
}




wb <- createWorkbook()
sheet_name <- 'OR_tables'
addWorksheet(wb = wb, sheetName = sheet_name)
#setColWidths(wb, sheet = 1, cols=1:100, widths = "auto") # set cols to automatically resize
setColWidths(wb, sheet = 1, cols=1:100, widths = "auto") # set cols to automatically resize
wb <- print.tabular.xlsx(wb, sheet='OR_tables', coords=c(1,1), tabular=t.wg.centrale, add.caption=TRUE)
wb <- print.tabular.xlsx(wb, 'OR_tables', c(1, nrow(t.wg.centrale)+4), t.hhr.centrale)
openXL(wb)

# TODO store table coords on wb object (per sheet)?

# print.tabular.xlsx()
# print.tabular.tex()
# print.tabular.md()
#
# write.tabular.list.xlsx <- function(wb, tables) {
#   
# }
# 
# add.strings(tabular, 'nl') <- [caption, col/row names, comment]
# get.strings(tabular, 'nl')


# toprule
# midrule
# bottomrule
# 