#' @title Add 'gt' table into a Word document
#'
#' @description Add a 'gt' table into a Word document. The table will be processed
#'   using the \link[gt]{as_word} function then inserted either after, before,
#'   or on the cursor location.
#'
#' @param x `rdocx` object
#' @param value `gt_tbl` object
#' @param align left, center (default) or right.
#' @param caption_location top (default), bottom, or embed indicating if the
#'   title and subtitle should be listed above, below, or be embedded in the
#'   table
#' @param caption_align left (default), center, or right. Alignment of caption
#'   (title and subtitle). Used when `caption_location` is not "embed".
#' @param split set to TRUE if you want to activate Word option 'Allow row to
#'   break across pages'.
#' @param keep_with_next Word option 'keep rows together' can be activated when
#'   TRUE. It avoids page break within tables.
#' @param pos where to add the gt table relative to the cursor, one of "after"
#'   (default), "before", "on" (end of line).
#'
#' @export
#'
#' @returns An updated rdocx object with a 'gt' table inserted
#'
#' @seealso flextable::body_add_flextable()
#'
#' @examples
#'
#'  library(officer)
#'  library(gt)
#'
#'  gt_tbl <- gt(head(exibble))
#'
#'  doc <- read_docx()
#'  doc <- body_add_gt(doc, value = gt_tbl)
#'  fileout <- tempfile(fileext = ".docx")
#'  print(doc, target = fileout)
#'
#'
#' @importFrom rlang is_installed arg_match
#' @importFrom officer body_add_xml
#' @importFrom gt as_word
#' @importFrom xml2 read_xml xml_children
#'
body_add_gt <- function(
  x,
  value,
  align = "center",
  pos = c("after","before","on"),
  caption_location = c("top","bottom","embed"),
  caption_align = "left",
  split = FALSE,
  keep_with_next = TRUE
) {

  ## check that officer is available
  if (!is_installed("officer")) {
    stop("{officer} package is necessary to add gt tables to word documents.")
  }

  ## check that inputs are an officer rdocx and gt tbl
  stopifnot(inherits(x, "rdocx"))
  stopifnot(inherits(value, "gt_tbl"))

  pos <- arg_match(pos)
  caption_location <- arg_match(caption_location)

  # Build all table data objects through a common pipeline
  tbl_ooxml <- as_word(
    data = value,
    align = align,
    caption_location = caption_location,
    caption_align = caption_align,
    split = split,
    keep_with_next = keep_with_next
    )

  tbl_nodes <- paste0(
    "<tablecontainer>",
    tbl_ooxml,
    "</tablecontainer>"
    ) %>%
    read_xml() %>%
    xml_children()  %>%
    suppressWarnings()

  order_to_add_nodes <- seq_along(tbl_nodes)
  if(pos == "before"){
    order_to_add_nodes <- rev(order_to_add_nodes)
  }

  for(i in order_to_add_nodes){
    x <- body_add_xml(x = x, str = tbl_nodes[[i]], pos = pos)
  }

  x
}
