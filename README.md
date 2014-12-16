tabulr
======

**tabulr** -- Write out tabular R datatructures to Excel.

Goal is are a set of robust functions that can take a list of R-objects containing a tabular structure, and write it to Excel using sensible layout and styling defaults. 

Such abular structure are for instance crosstables, output from `syvtable()` and `table()`, a table of descriptives or [model parameters](https://github.com/dgrtwo/broom), a dataframe, etc.

# Example

```R
library(openxlsx)
source('table_export_openxlsx.r')
wb <- createWorkbook()
addWorksheet(wb = wb, sheetName = 'OR_tables', gridLines = FALSE)
wb <- print.tabulars.xlsx(wb, 1, list(tab1, tab2, tab3))
saveWorkbook(wb, "OR1_descr_tables.xlsx",
```

# TODO

[X] rewrite functions using [openxlsx](https://github.com/awalker89/openxlsx).
[X] basic layout routine for row & col names, row and col margins, captions and comments.
[X] minimal styling based on LaTex tables
[ ] auto column resizing (in progress, cf. [issue #43](https://github.com/awalker89/openxlsx/issues/43))
[ ] easy setters and getters for tabular properties (e.g. `caption()`, `comment()`)
[ ] optional: cell highlighting for significance-tests?
[ ] optional: multilingual strings support?
