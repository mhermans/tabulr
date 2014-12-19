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
                               add.caption=TRUE, add.comment=TRUE, 
                               add.row.margin=TRUE, add.col.margin=TRUE,
                               style=None) {
  #print(str(tabular))
  
  if ( is.null(caption(tabular)) ) { add.caption  <- FALSE}
  if ( is.null(col.margin(tabular)) ) { add.col.margin  <- FALSE}
  
  # make sure the object is a data.frame, convert if needed
  # -------------------------------------------------------

  caption <- attr(tabular, 'caption')
  col.margin <- attr(tabular, 'col.margin')
  
  if ('matrix' %in% class(tabular) ) { tabular <- as.data.frame(tabular) }
  
  if ('svytable' %in% class(tabular)) { tabular <- as.data.frame.matrix(tabular) }
  
  attr(tabular, 'caption') <- caption
  attr(tabular, 'col.margin') <- col.margin
  rm(caption, col.margin)
  
  
  # dimensions
  # ----------
  
  #   1 caption row       x   [spanning caption colums]
  #   2 white row         x   1 rownames col  + 3 ncol cols + 1 row margin col
  #   3 colnames row      x   1 whitespace    + 3 ncol rows + 1 row margin col
  #   4 data row          x   1 rowname col   + 3 ncol rows + 1 row margin col
  #   5 data row          x   1 rowname col   + 3 ncol rows + 1 row margin col
  #   6 data row          x   1 rowname col   + 3 ncol rows + 1 row margin col
  #   7 col margin row    x   1 whitespace    + 3 ncol rows + 1 row margin col
  #   8 comment row       x   [spanning comment columns]
  
  
  # determine position index of rows
  # --------------------------------
  
  caption_table_space <- 1 # optional whiteline => 1 ipv 2?
  
  n_data_rows <- nrow(tabular)
  
  table_start_r <- coords[1] # consider coords (r,c)
  caption_r <- table_start_r
  
  # data includes rownames, colnames TODO: change to make option?
  data_start_r <- table_start_r
  
  # if caption, start the table two rows lower
  if (add.caption) { data_start_r <- data_start_r + caption_table_space + 1 } # 
  
  #col_names_r <- data_start_r - 1 # colnames one above data # <-> currently included in data
  col_names_r <- data_start_r 
  
  data_end_r <- data_start_r + n_data_rows
  
  table_end_r <- data_end_r
  
  col_margin_r <- data_end_r + 1  

  if (add.col.margin == TRUE) { table_end_r <- table_end_r + 1 } 
  
  # caption and comment are outside of table start-end dimensions?
  
  if ( add.comment == TRUE ) { 
    comment_r <- data_end_r + 1 }
  if ( add.col.margin == TRUE ) { comment_r <- col_margin_r + 1 }

  
  
#   n_total_rows <- n_data_rows
#   if (add.caption = TRUE) { n_total_rows + 1 }
#   if (add.col.margin = TRUE) { n_total_rows + 1 }
#   if (add.comment = TRUE) { n_total_rows + 1 }
  
  # determine position index of cols
  # --------------------------------

  n_data_cols <- ncol(tabular)
  table_start_c <- coords[2]
  data_start_c <- table_start_c

  # TODO: ook mogelijk maken dat er geen rownames zijn?
  row_margin_c <- 1 + ncol(tabular) + 1
  
  # TODO: ook mogelijk maken dat er geen colnames zijn?
  col_margin_c <- data_start_c + 1

  table_end_c <- 1 + ncol(tabular)
  if (add.row.margin == TRUE) { table_end_c <- table_end_c + 1 } 

  


  # write out data rows/cols, including row & col names
  # ---------------------------------------------------
  
  writeData(
    wb = wb, sheet = sheet, startCol = data_start_c, startRow = data_start_r,
    x = tabular, rowNames=TRUE, colNames=TRUE,
    borders = "none")
  
  
  # add caption-row and caption (on top by default)
  # -----------------------------------------------
  
  if (add.caption) {
    caption_text <- caption(tabular)
    
    # add row contents
    # ----------------
    
    # row for comment is the most top one, i.e. table_start_r
    writeData(
      wb = wb, sheet = sheet, startCol = table_start_c, startRow = caption_r,
      x = caption_text, rowNames=FALSE, colNames=FALSE)
    
    # for caption, merge the cells to span multiple cols, either width table, or 
    # some minimum nr. of cols
    
    caption_merge_width <- ifelse(7 - table_start_c >= 6, 6)
    mergeCells(wb, sheet=sheet, 
               cols = table_start_c:caption_merge_width, 
               rows = table_start_r)
    
    # TODO: auto col width should not be affected
    # cf. https://github.com/awalker89/openxlsx/issues/43
    
  }
  
  
  # if margins, add row/and or col margin
  # -------------------------------------
  
  if (add.row.margin) {
    row.margin.contents <- t(rep(100, nrow(tabular))) # TODO, parametriseer
    
    for (i in 1:length(row.margin.contents)) {
      writeData(
        wb = wb, sheet = sheet, startCol = row_margin_c, startRow = data_start_r+i,
        x = row.margin.contents[i], rowNames=FALSE, colNames=FALSE)  
    }
    
    
  }
  
  if (add.col.margin) {
    col.margin.contents <- t(as.data.frame(as.vector(col.margin(tabular))))
    
    writeData(
      wb = wb, sheet = sheet, startCol = col_margin_c, startRow = col_margin_r,
      x = col.margin.contents, rowNames=FALSE, colNames=FALSE)  
     
  } 
  
  #   row.margin.name
  #   col.margin.name
  #   
  #   margin.table(tab)
  
  
  # if comment, add additional row with comment
  # -------------------------------------------
  
  if (add.comment) {
    comment.contents <- 'N = 1200 (NA = 100), chi^2 = 3, p 0.000. Gewogen steekproef ' # TODO, parametriseer
    
    writeData(
      wb = wb, sheet = sheet, startCol = table_start_c, startRow = comment_r,
      x = comment.contents, rowNames=FALSE, colNames=FALSE)  
    
    comment_merge_width <- ifelse(7 - table_start_c >= 6, 6)
    mergeCells(wb, sheet=sheet, 
               cols = table_start_c:comment_merge_width, 
               rows = comment_r)
    
  } 
  
  
  # add horizontal line styling
  # ---------------------------
  
  table_top_row_style <- createStyle(border="Top", borderStyle = "medium", borderColour="#000000", halign='right')
  table_bottom_row_style <- createStyle(border="Top", borderStyle = "medium", borderColour="#000000")
  table_mid_row_style <- createStyle(border="Top", borderStyle = "thin", borderColour="#000000")
  #rc_names_style <- createStyle(textDecoration="bold")
  
  # table top rule -> add to start_table_row
  addStyle(wb, sheet = sheet, table_top_row_style, rows = col_names_r, cols = table_start_c:table_end_c, gridExpand = TRUE)
  addStyle(wb, sheet = sheet, table_mid_row_style, rows = col_names_r+1, cols = table_start_c:table_end_c, gridExpand = TRUE)

  # no col margin -> only bottom line
  if (add.col.margin == FALSE) {
    addStyle(wb, sheet = sheet, table_bottom_row_style, rows = table_end_r+1, cols = table_start_c:table_end_c, gridExpand = TRUE)
  }else{
  # col margin -> mid lid en bottom line
    addStyle(wb, sheet = sheet, table_mid_row_style, rows = table_end_r, cols = table_start_c:table_end_c, gridExpand = TRUE)
    addStyle(wb, sheet = sheet, table_bottom_row_style, rows = table_end_r+1, cols = table_start_c:table_end_c, gridExpand = TRUE)
  }

#   addStyle(wb, sheet = 1, table_mid_row_style, rows = 12, cols = 1:6, gridExpand = TRUE)
#   addStyle(wb, sheet = 1, table_bottom_row_style, rows = 13, cols = 1:6, gridExpand = TRUE)
  
  
  
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


print.tabulars.xlsx <- function(wb, sheet, tabulars, start_c=1, spacer_rows=2) {
  # todo: specify by sheet index or name

  for (tabular in tabulars) {
    
    # check the last 
    previous_r <- last_filled_row(wb, sheet)
    
    if (previous_r == 0) { 
      start_r <- 1
    }else {
      start_r <- previous_r + 1 + spacer_rows
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
