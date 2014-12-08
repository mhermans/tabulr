# TABULR NOTES 

Goal: robust Excel export for core two-dimenional tabular datastructures

tab1 <- svytable()
caption(tab1) <- 'Tabel caption'


# toprule
# midrule
# bottomrule
# 

Two dimensional tabular datastructures:

* data.frame
* table
* svytable¨
* structable? [3+ dimensions]
 

Targeted output formats:

* LaTeX
* Excel
* 


## Available packages

* `xlsx` https://code.google.com/p/rexcel/
* `openxlsx` https://github.com/awalker89/openxlsx
* `XLConnect` https://github.com/miraisolutions/xlconnect

See also:

* http://cran.r-project.org/web/packages/tables/index.html
* kable() in `knitr`

openxlsx vs XLConnect vs xlsx
	-> openxlsx heeft geen dependency op Java, meer high level, minder goed gedocumenteerd als XLConnect


## 

# TODO wat is een header/footer? in Excel
 
## Design of tables

* http://www.r-project.org/conferences/useR-2007/program/posters/weigand.pdf
* Banner table: http://www.woelfelresearch.com/bannerTables.html
* http://stat405.had.co.nz/lectures/19-tables.pdf
* http://marcoghislanzoni.com/blog/2013/10/11/pivot-tables-in-r-with-melt-and-cast/
* http://www.stata.com/stata-news/news29-1/export-tables-to-excel/
* http://blog.stata.com/2013/09/25/export-tables-to-excel/
* http://statmethods.wordpress.com/2014/06/19/quickly-export-multiple-r-objects-to-an-excel-workbook/
* [Small Guide to Making Nice Tables](http://www.inf.ethz.ch/personal/markusp/teaching/guides/guide-tables.pdf) (latex-georienteerd)
* http://truben.no/table/



define TableTemplate? e.g. bannertable, crosstable, datatable

tabulr | write tabular datastructures

Related: xtable, pander

# not possible to seperate style and content in table construction

A Tabular object should contain all the structural information and values needed so that it can be passeed to different backends...

convert table(), svytable(), etc. into Tabular objects

as.tabular(my_svy_tab, margins="sum"|c(10), col.class, comment='test|N', caption='mycaption', ) # default cross table

a some point objet should no be more modifyable, e.g. calculate custom margins, specific tests, etc.

# style
# alignment

Tabular object
	rectangular data frame
	strings
		en
			caption
			row.labels
			col.labels
		nl
	

	


* Most frequently, but not necessary crosstabs, e.g. "svy.datatable"
* frequencies vs R/C percentages prop.table
* R/C margin totals margin.table, addmargins()
* Captions
* "Note-line" with N, statistical test
* significatietoets/sterretjes per R/C? via colour?
* stacked headers?


# Multilangual labels support?

labels list of
	- 'nl'
		- caption string
		- rownames list of strings
		- colnames list of strings
	- 'fr'
		- caption string
		- rownames
		- colnames
	
-> if only one => select those strings
-> if multiple, check for global option, e.g. tabulr.lang.default
-> if global default lang is set set in document, use system lang setting
-> if selected language is undefined => error or warning with fallback to other lang?


print.tabulr 


Standard tabular forms in Excel rapport Liagre: