# Function to skip tests if Suggested packages not available on system
check_suggests_xml <- function() {
  skip_if_not_installed("officer")
  skip_if_not_installed("xml2")
}

## this is used across the tests, so simplest approach is to load the packages

suppressWarnings({
suppressPackageStartupMessages({
  library(gt)
  library(dplyr)
})})

test_that("tables can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <-
    gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() %>%
    body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table caption
  docx_table_caption_text <- xml2::xml_text(docx_contents[1:2])

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    xml2::xml_text(xml2::xml_find_all(docx_table_body_header, ".//w:p")),
    c(
      "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables can be added to a word doc - position 'before'", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <-
    gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() %>%
    officer::body_add_par("this is paragraph 1") %>%
    officer::body_add_par("this is paragraph 2") %>%
    officer::cursor_end() %>%
    body_add_gt(
      gt_exibble_min,
      align = "center",
      pos = "before"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  ## drop extra
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))
  docx_table_contents <- docx_contents[c(2:4)]
  docx_previous_inserts <- docx_contents[c(1,5)]

  ## test "previous" contents
  expect_equal(
    xml2::xml_text(docx_previous_inserts),
    c("this is paragraph 1","this is paragraph 2")
  )

  ## extract table caption
  docx_table_caption_text <- xml2::xml_text(docx_table_contents[1:2])

  ## extract table contents
  docx_table_body_header <-
    docx_table_contents[3] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_table_contents[3] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    xml2::xml_text(xml2::xml_find_all(docx_table_body_header, ".//w:p")),
    c(
      "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})

test_that("tables with special characters can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <-
    gt::exibble[1,] %>%
    dplyr::mutate(special_characters = "><&\"'") %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() %>%
    body_add_gt(
      gt_exibble_min,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table caption
  docx_table_caption_text <- xml2::xml_text(docx_contents[1:2])

  ## extract table contents
  docx_table_body_header <-
    docx_contents[3] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[3] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    xml2::xml_text(xml2::xml_find_all(docx_table_body_header, ".//w:p")),
    c(
      "num", "char", "fctr", "date", "time",
      "datetime", "currency", "row", "group",
      "special_characters"
    )
  )

  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a",
        "><&\"'"
      )
    )
  )
})

test_that("tables with embedded titles can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <-
    gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() %>%
    body_add_gt(
      gt_exibble_min,
      caption_location = "embed",
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:t") %>%
      xml2::xml_text(),
    c(
      "table title", "table subtitle", "num", "char", "fctr",
      "date", "time","datetime", "currency", "row", "group"
    )
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text()),
    list(
      c(
        "0.1111",
        "apricot",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "2.2220",
        "banana",
        "two",
        "2015-02-15",
        "14:40",
        "2018-02-02 14:33",
        "17.95",
        "row_2",
        "grp_a"
      )
    )
  )
})
