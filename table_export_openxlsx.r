library(openxlsx)


# GENERAL AND HELPER FUNCTIONS #
# ============================ #


# basis-elementen die rondom de data geplaatst worden
#
# * caption       # col onder of boven toevoegen, spanning?
# * row.margin 		# col rechts toevoegen
# * col.margin 		# row onder toevoegen
# * comment			  # row onder toevoegen [italic?)

print.tabular.xlsx <- function(wb, sheet, coords, tabular, 
                               add.caption=FALSE, add.comment=FALSE, 
                               add.row.margin=FALSE, add.col.margin=FALSE,
                               style=None) {
  
  # make sure the object is a data.frame, convert if needed
  # -------------------------------------------------------
  
  if ('matrix' %in% class(tabular) ) { tabular <- as.data.frame(tabular) }
  if ('svytable' %in% class(tabular)) { tabular <- as.data.frame.matrix(tabular) }
  
  # consider coords (r,c)
  start_r <- coords[1]
  start_c <- coords[2]
  
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
  
  #   row.margin.name
  #   col.margin.name
  #   
  #   margin.table(tab)
  
  
  # if comment, add additional row with comment
  # -------------------------------------------
  
  
  
  # return wb (save if filename?)
  # -----------------------------
  wb
  
}


last_filled_row <- function(wb, sheet_number) {
  # returns the indexnumber of the last row with data on it
  # returns 0 if no rows contain data
  
  sheet_data <- wb$sheetData[[sheet_number]]
  if (length(sheet_data) == 0 ) {
    max_row <- 0
  }else{
    max_row <- max(as.integer(names(sheet_data)))  
  }
  
  max_row
}


print.tabulars.xlsx <- function(wb, sheet, tabulars, start_c=1) {
  # todo: specify by sheet index or name

  for (tabular in tabulars) {
    
    # check the last 
    previous_r <- last_filled_row(wb, sheet)
    
    if (previous_r == 0) { 
      start_r <- 1
    }else {
      start_r <- previous_r + 2
    }    
    
    coords <- c(start_r, start_c)
    print.tabular.xlsx(wb, sheet, coords, tabular) # modify in place!
    #openXL(wb)
  }
  
  wb
  
}


#write.xlsx(t.wg.centrale, file = "writeXLSXTable1.xlsx", asTable = TRUE) # error, geen data.frame

# options: simple data, or automatically formatedd as a table
# write.xlsx(tab, file = "OR1_descr_tables.xlsx", asTable = TRUE, # issue, "row.name"s als colname
#            col.names=TRUE, row.names=TRUE) 
