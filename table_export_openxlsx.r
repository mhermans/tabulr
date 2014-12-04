wd <- getwd()
setwd('C:/Users/MaartenH/Documents/work/HIVA/acv_or')
load('rapportering/descr_tables.RData')
setwd(wd)

library(openxlsx)

options("openxlsx.borderColour" = "#4F80BD")
options("openxlsx.borderStyle" = "thin")

#write.xlsx(t.wg.centrale, file = "writeXLSXTable1.xlsx", asTable = TRUE) # error, geen data.frame


tab <- as.data.frame(t.wg.centrale)

write.xlsx(tab, file = "OR1_descr_tables.xlsx", asTable = TRUE, col.names=TRUE, row.names=TRUE) # issue, "row.name"s als colname

write.xlsx(tab, file = "OR1_descr_tables.xlsx", asTable = FALSE, row.names=TRUE)

?write.xlsx


hs1 <- createStyle(fgFill = "#DCE6F1", halign = "CENTER", textDecoration = "Italic",
                   border = "Bottom")

wb <- createWorkbook()
test.n <- "data.frame"
my.df <- tab
addWorksheet(wb = wb, sheetName = test.n)
setColWidths(wb, sheet = 1, cols=2:100, widths = "auto")
writeData(wb = wb, sheet = test.n, x = my.df, xy=c(1,1),
          borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")

writeData(wb = wb, sheet = test.n, x = my.df, xy=c(1,nrow(tab) + 3),
          borders = "surrounding", rowNames=TRUE, headerStyle = hs1, borderStyle = "dashed")




openXL(wb)
#saveWorkbook(wb, "setHeaderFooterExample.xlsx", overwrite = TRUE)

# wat is een header/footer?


