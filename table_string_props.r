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

"caption<-data.frame" <- function(x,value) {
  if (length(value)>2)
    stop("\"caption\" must have length 1 or 2")
  attr(x,"caption") <- value
  return(x)
}

caption.data.frame <- function(x,...) {
  return(attr(x,"caption",exact=TRUE))
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
