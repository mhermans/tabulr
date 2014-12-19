# CAPTION FUNCTIONS #
# ================= #

# table
# -----

"caption<-" <- function(x,value) UseMethod("caption<-")
"caption<-.table" <- function(x,value) {
  if (length(value)>2)
    stop("\"caption\" must have length 1 or 2")
  attr(x,"caption") <- value
  return(x)
}
caption <- function(x,...) UseMethod("caption")
caption.table <- function(x,...) {
  return(attr(x,"caption",exact=TRUE))
}

# matrix
# ------

"caption<-.matrix" <- function(x,value) {
  if (length(value)>2)
    stop("\"caption\" must have length 1 or 2")
  attr(x,"caption") <- value
  return(x)
}

caption.matrix <- function(x,...) {
  return(attr(x,"caption",exact=TRUE))
}


# data.frame
# ------

"caption<-.data.frame" <- function(x,value) {
  if (length(value)>2)
    stop("\"caption\" must have length 1 or 2")
  attr(x,"caption") <- value
  return(x)
}

caption.data.frame <- function(x,...) {
  return(attr(x,"caption",exact=TRUE))
}


# COL MARGIN FUNCTIONS
# ====================

# table
# -----

"col.margin<-" <- function(x,value) UseMethod("col.margin<-")
"col.margin<-.table" <- function(x,value) {
  attr(x,"col.margin") <- value
  return(x)
}
col.margin <- function(x,...) UseMethod("col.margin")
col.margin.table <- function(x,...) {
  return(attr(x,"col.margin",exact=TRUE))
}

# matrix
# ------

"col.margin<-.matrix" <- function(x,value) {
  attr(x,"col.margin") <- value
  return(x)
}

col.margin.matrix <- function(x,...) {
  return(attr(x,"col.margin",exact=TRUE))
}


# data.frame
# ------

"col.margin<-.data.frame" <- function(x,value) {
  attr(x,"col.margin") <- value
  return(x)
}

col.margin.data.frame <- function(x,...) {
  return(attr(x,"col.margin",exact=TRUE))
}


# ROW MARGIN FUNCTIONS
# ====================

# table
# -----

"row.margin<-" <- function(x,value) UseMethod("row.margin<-")
"row.margin<-.table" <- function(x,value) {
  attr(x,"row.margin") <- value
  return(x)
}
row.margin <- function(x,...) UseMethod("row.margin")
row.margin.table <- function(x,...) {
  return(attr(x,"row.margin",exact=TRUE))
}

# matrix
# ------

"row.margin<-.matrix" <- function(x,value) {
  attr(x,"row.margin") <- value
  return(x)
}

row.margin.matrix <- function(x,...) {
  return(attr(x,"row.margin",exact=TRUE))
}


# data.frame
# ------

"row.margin<-.data.frame" <- function(x,value) {
  attr(x,"row.margin") <- value
  return(x)
}

row.margin.data.frame <- function(x,...) {
  return(attr(x,"row.margin",exact=TRUE))
}


# multilingual captions
# options("tabular.strings.lang" = "nl")
# getOption("openxlsx.borderColour", "black")
# getOption("tabular.strings.lang")
# 
# # multilingual caption
# 
# "caption<-.matrix" <- function(x,value, lang=getOption("openxlsx.borderColour", "en") )  {
#   if (length(value)>2)
#     stop("\"caption\" must have length 1 or 2")
#   attr(x,"caption") <- value
#   return(x)
# }
# 
# caption.matrix <- function(x,...) {
#   return(attr(x,"caption",exact=TRUE))
# }
# 
