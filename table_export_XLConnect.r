library(XLConnect)

# schrijffunctie
# --------------

writeTable <- function(table, excel_fn) {
  wbFilename <- excel_fn
  wb = loadWorkbook(wbFilename, create = TRUE)
  sheet <- 'tables'
  createSheet(wb, sheet)
  nameLocation <- paste(sheet, "$A$1", sep='!')
  createName(wb, name = 'tabName', formula = nameLocation)
  writeNamedRegion(wb, data = table, name = tabName, header = TRUE, rownames = '')
  saveWorkbook(wb)
  
}


# Voorbeeld van table i/d vorm van een data.frame
# -----------------------------------------------

t.wg.centrale <- cbind(
  table(OR1.mg.real$centrale),
  svytable(~centrale, OR1.svydesign),
  round(table(OR1.mg.real$centrale)/sum(table(OR1.mg.real$centrale))*100,1),
  round(svytable(~centrale, OR1.svydesign) / sum(svytable(~centrale, OR1.svydesign))*100,1),
  round(prop.table(table(OR1.pop[!duplicated(OR1.pop$id_frc),]$centrale))*100,1))

colnames(t.wg.centrale) <- c('Gereal. (N)', 'Gewogen (N)', 'Gereal. (%)', 'Gewogen (%)', 'Populatie (%)')
rownames(t.wg.centrale) <- c('BIE', 'Met', 'VeD', 'Tra', 'LBC', 'CNE', 'OwN', 'OwF', 'OpD')

writeTable(t.wg.centrale, 'ACV_OR_tables2.xlsx')


# voorbeeld van een table i/d vorm van een svytable object
# --------------------------------------------------------

# XLConnect verwacht standaard data.frame-objecten. Het 
# object v svytable() dient eerst omgezet te worden

clbls <- c('Altijd', 'Meestal', 'Soms', 'Zelden', 'Nooit', 'W/n')
vertrouw.tab <- svy.crosstab(
  rowvar='sector_c3', colvar='v11c', 
  design=OR1.svydesign, 
  rownames=sector_3c_lbls, colnames=clbls, tex=FALSE,
  caption='Hoe vaak bestempelt de werkgever onder tegenkanting informatie als vertrouwelijk? (naar sector, \\%)')

vertrouw.tab <- as.data.frame.matrix(vertrouw.tab) # zet om naar dataframe

writeTable(vertrouw.tab, 'ACV_OR_tables2.xlsx')
