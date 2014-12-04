# TABULR NOTES 

Goal: robust Excel export for core two-dimenional tabular datastructures

tab1 <- svytable()
caption(tab1) <- 'Tabel caption'


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

openxlsx vs XLConnect vs xlsx
	-> openxlsx heeft geen dependency op Java, meer high level, minder goed gedocumenteerd als XLConnect




 
Design of tables

* http://www.r-project.org/conferences/useR-2007/program/posters/weigand.pdf
* Banner table: http://www.woelfelresearch.com/bannerTables.html
* http://stat405.had.co.nz/lectures/19-tables.pdf
* http://marcoghislanzoni.com/blog/2013/10/11/pivot-tables-in-r-with-melt-and-cast/
* http://www.stata.com/stata-news/news29-1/export-tables-to-excel/
* http://blog.stata.com/2013/09/25/export-tables-to-excel/
* http://statmethods.wordpress.com/2014/06/19/quickly-export-multiple-r-objects-to-an-excel-workbook/

tabulr | write tabular datastructures

Related: xtable, pander




	


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



Standard tabular forms in Excel rapport Liagre: