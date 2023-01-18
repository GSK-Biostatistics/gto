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

test_that("tables with spans can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    ) %>%
    ## add spanner across columns 1:5
    gt::tab_spanner(
      "My Column Span",
      columns = 3:5
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <- docx_contents[1:2] %>%
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <- docx_contents[3] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[3] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c( "","","My Column Span", "","","","",
       "num", "char","fctr", "date", "time","datetime", "currency", "row","group")
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

test_that("tables with multi-level spans can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_header(
      title = "table title",
      subtitle = "table subtitle"
    ) %>%
    ## add spanner across columns 1:5
    gt::tab_spanner(
      "My 1st Column Span L1",
      columns = 1:5
    ) %>%
    gt::tab_spanner(
      "My Column Span L2",
      columns = 2:5,level = 2
    ) %>%
    gt::tab_spanner(
      "My 2nd Column Span L1",
      columns = 8:9
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <- docx_contents[1:2] %>%
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <- docx_contents[3] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[3] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: table title", "table subtitle")
  )

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c("", "My Column Span L2", "","","","",
      "My 1st Column Span L1", "","", "My 2nd Column Span L1",
      "num", "char","fctr", "date", "time","datetime", "currency", "row","group")
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

test_that("tables with summaries can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- gt::exibble %>%
    dplyr::select(-c(fctr, date, time, datetime)) %>%
    gt::gt(rowname_col = "row", groupname_col = "group") %>%
    gt::summary_rows(
      groups = TRUE,
      columns = "num",
      fns = list(
        avg = ~mean(., na.rm = TRUE),
        total = ~sum(., na.rm = TRUE),
        s.d. = ~sd(., na.rm = TRUE)
      ),
      formatter = fmt_number
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  ## "" at beginning for stubheader
  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c( "", "num", "char", "currency")
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text()),
    list(
      "grp_a",
      c("row_1", "1.111e-01", "apricot", "49.950"),
      c("row_2", "2.222e+00", "banana", "17.950"),
      c("row_3", "3.333e+01", "coconut", "1.390"),
      c("row_4", "4.444e+02", "durian", "65100.000"),
      c("avg", "120.02", "—", "—"),
      c("total", "480.06", "—", "—"),
      c("s.d.", "216.79", "—", "—"),
      "grp_b",
      c("row_5", "5.550e+03", "NA", "1325.810"),
      c("row_6", "NA", "fig", "13.255"),
      c("row_7", "7.770e+05", "grapefruit", "NA"),
      c("row_8", "8.880e+06", "honeydew", "0.440"),
      c("avg", "3,220,850.00", "—", "—"),
      c("total", "9,662,550.00", "—", "—"),
      c("s.d.", "4,916,123.25", "—", "—")
    )
  )
})

test_that("tables with footnotes can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <-
    gt::exibble[1:2,] %>%
    gt::gt() %>%
    gt::tab_footnote(
      footnote = md("this is a footer example"),
      locations = cells_column_labels(columns = num )
    ) %>%
    gt::tab_footnote(
      footnote = md("this is a second footer example"),
      locations = cells_column_labels(columns = char )
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
  print(word_doc, target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <-
    docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  ## superscripts will display as "true#false" due to
  ## xml being:
  ## <w:vertAlign w:val="superscript"/><w:i>true</w:i><w:t xml:space="default">1</w:t><w:i>false</w:i>,
  ## and being converted to TRUE due to italic being true, then the superscript, then turning off italics
  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c(
      "numtrue1false", "chartrue2false", "fctr", "date", "time",
      "datetime", "currency", "row", "group"
    )
  )

  ## superscripts will display as "true##" due to
  ## xml being:
  ## <w:vertAlign w:val="superscript"/><w:i>true</w:i><w:t xml:space="default">1</w:t>,
  ## and being converted to TRUE due to italic being true, then the superscript,
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
      ),
      c("true1this is a footer example"),
      c("true2this is a second footer example")
    )
  )
})

test_that("tables with source notes can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- gt::exibble[1:2, ] %>%
    gt::gt() %>%
    gt::tab_source_note(source_note = "this is a source note example")

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
    body_add_gt(gt_exibble_min,
                align = "center")

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc, target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c(
      "num",
      "char",
      "fctr",
      "date",
      "time",
      "datetime",
      "currency",
      "row",
      "group"
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
      ),
      c("this is a source note example")
    )
  )
})

test_that("long tables can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_letters <- tibble::tibble(
    upper_case = c(LETTERS,LETTERS),
    lower_case = c(letters,letters)
  ) %>%
    gt::gt() %>%
    gt::tab_header(
      title = "LETTERS"
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
    body_add_gt(
      gt_letters,
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <- docx_contents[1] %>%
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <- docx_contents[2] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[2] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: LETTERS")
  )

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c("upper_case", "lower_case")
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text()),
    lapply(c(1:26,1:26),function(i)c(LETTERS[i], letters[i]))
  )
})

test_that("long tables with spans can be added to a word doc", {

  check_suggests_xml()

  ## simple table
  gt_letters <- tibble::tibble(
    upper_case = c(LETTERS,LETTERS),
    lower_case = c(letters,letters)
  ) %>%
    gt::gt() %>%
    gt::tab_header(
      title = "LETTERS"
    ) %>%
    gt::tab_spanner(
      "LETTERS",
      columns = 1:2
    )

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
    body_add_gt(
      gt_letters,
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table caption
  docx_table_caption_text <- docx_contents[1] %>%
    xml2::xml_text()

  ## extract table contents
  docx_table_body_header <- docx_contents[2] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[2] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  expect_equal(
    docx_table_caption_text,
    c("Table  SEQ Table \\* ARABIC 1: LETTERS")
  )

  expect_equal(
    docx_table_body_header %>%
      xml2::xml_find_all(".//w:p") %>%
      xml2::xml_text(),
    c("LETTERS", "upper_case", "lower_case")
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text()),
    lapply(c(1:26,1:26),function(i)c(LETTERS[i], letters[i]))
  )
})

test_that("tables with cell & text coloring can be added to a word doc - no spanner", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- gt::exibble[1:4,] %>%
    gt::gt(rowname_col = "char") %>%
    gt::tab_row_group("My Row Group 1",c(1:2)) %>%
    gt::tab_row_group("My Row Group 2",c(3:4)) %>%
    gt::tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_body(
        columns = c(num,fctr,time,currency, group)
      )
    ) %>%
    gt::tab_style(
      style = cell_text(
        color = "green",
        font = "Biome"
      ),
      locations = cells_stub()
    ) %>%
    gt::tab_style(
      style = cell_text(size = 25, v_align = "middle"),
      locations = cells_body(
        columns = c(num,fctr,time,currency, group)
      )
    ) %>%
    tab_style(
      style = cell_text(
        color = "blue",
        stretch = "extra-expanded"
      ),
      locations = cells_row_groups()
    ) %>%
    tab_style(
      style = cell_text(color = "teal"),
      locations = cells_column_labels()
    ) %>%
    tab_style(
      style = cell_fill(color = "green"),
      locations = cells_column_labels()
    ) %>%
    tab_style(
      style = cell_fill(color = "pink"),
      locations = cells_stubhead()
    )

  if (!testthat::is_testing() & interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  ## header
  expect_equal(
    docx_table_body_header %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text(),
    c("","num", "fctr","date", "time","datetime","currency",  "row", "group")
  )
  expect_equal(
    lapply(docx_table_body_header, function(x) x %>% xml2::xml_find_all(".//w:shd") %>% xml2::xml_attr(attr = "fill")),
    list(c("FFC0CB","00FF00","00FF00","00FF00","00FF00","00FF00","00FF00","00FF00","00FF00"))
  )
  expect_equal(
    lapply(docx_table_body_header, function(x) x %>% xml2::xml_find_all(".//w:color") %>% xml2::xml_attr(attr = "val")),
    list(c("008080","008080","008080","008080","008080","008080","008080","008080"))
  )

  ## cell background styling
  expect_equal(
    lapply(docx_table_body_contents, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% lapply(function(y) {
        y %>% xml2::xml_find_all(".//w:shd") %>% xml2::xml_attr(attr = "fill")
      })}),
    list(
      list(character()),
      list(character(),"FFA500","FFA500",character(),"FFA500",character(),"FFA500",character(),"FFA500"),
      list(character(),"FFA500","FFA500",character(),"FFA500",character(),"FFA500",character(),"FFA500"),
      list(character()),
      list(character(),"FFA500","FFA500",character(),"FFA500",character(),"FFA500",character(),"FFA500"),
      list(character(),"FFA500","FFA500",character(),"FFA500",character(),"FFA500",character(),"FFA500")
    )
  )

  ## cell text styling
  expect_equal(
    lapply(docx_table_body_contents, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% lapply(function(y) {
        y %>% xml2::xml_find_all(".//w:color") %>% xml2::xml_attr(attr = "val")
      })}),
    list(
      list("0000FF"),
      list("00FF00",character(),character(),character(),character(),character(),character(),character(),character()),
      list("00FF00",character(),character(),character(),character(),character(),character(),character(),character()),
      list("0000FF"),
      list("00FF00",character(),character(),character(),character(),character(),character(),character(),character()),
      list("00FF00",character(),character(),character(),character(),character(),character(),character(),character())
    )
  )

  expect_equal(
    lapply(docx_table_body_contents, function(x)
      x %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text()),
    list(
      "My Row Group 2",
      c(
        "coconut",
        "33.3300",
        "three",
        "2015-03-15",
        "15:45",
        "2018-03-03 03:44",
        "1.39",
        "row_3",
        "grp_a"
      ),
      c(
        "durian",
        "444.4000",
        "four",
        "2015-04-15",
        "16:50",
        "2018-04-04 15:55",
        "65100.00",
        "row_4",
        "grp_a"
      ),
      "My Row Group 1",
      c(
        "apricot",
        "0.1111",
        "one",
        "2015-01-15",
        "13:35",
        "2018-01-01 02:22",
        "49.95",
        "row_1",
        "grp_a"
      ),
      c(
        "banana",
        "2.2220",
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

test_that("tables with cell & text coloring can be added to a word doc - with spanners", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- exibble[1:4,] %>%
    gt(rowname_col = "char") %>%
    tab_row_group("My Row Group 1",c(1:2)) %>%
    tab_row_group("My Row Group 2",c(3:4)) %>%
    tab_spanner("My Span Label", columns = 1:5) %>%
    tab_spanner("My Span Label top", columns = 2:4, level = 2) %>%
    tab_style(
      style = cell_text(color = "purple"),
      locations = cells_column_labels()
    ) %>%
    tab_style(
      style = cell_fill(color = "green"),
      locations = cells_column_labels()
    ) %>%
    tab_style(
      style = cell_fill(color = "orange"),
      locations = cells_column_spanners("My Span Label")
    ) %>%
    tab_style(
      style = cell_fill(color = "red"),
      locations = cells_column_spanners("My Span Label top")
    ) %>%
    tab_style(
      style = cell_fill(color = "pink"),
      locations = cells_stubhead()
    )

  if (!testthat::is_testing() & interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  ## extract table contents
  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  ## header
  expect_equal(
    docx_table_body_header %>% xml2::xml_find_all(".//w:p") %>% xml2::xml_text(),
    c("", "", "My Span Label top", "", "", "", "", "",
      "", "My Span Label", "", "", "", "",
      "", "num", "fctr", "date", "time", "datetime", "currency", "row", "group")
  )
  expect_equal(
    lapply(docx_table_body_header, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% lapply(function(y) {
        y %>% xml2::xml_find_all(".//w:shd") %>% xml2::xml_attr(attr = "fill")
      })}),
    list(
      list("FFC0CB", character(0), "FF0000", character(0), character(0), character(0), character(0), character(0)),
      list(character(0), "FFA500", character(0), character(0), character(0), character(0)),
      list(character(0), "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00", "00FF00")
    )
  )
  expect_equal(
    lapply(docx_table_body_header, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% lapply(function(y) {
        y %>% xml2::xml_find_all(".//w:color") %>% xml2::xml_attr(attr = "val")
      })}),
    list(
      list(character(0), character(0), character(0), character(0),character(0), character(0), character(0), character(0)),
      list(character(0), character(0), character(0), character(0),character(0), character(0)),
      list(character(0), "A020F0","A020F0", "A020F0", "A020F0", "A020F0", "A020F0", "A020F0","A020F0")
    )
  )
})

test_that("tables with cell & text coloring can be added to a word doc - with source_notes and footnotes", {

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- exibble[1:2,] %>%
    gt() %>%
    tab_source_note("My Source Note") %>%
    tab_footnote("My Footnote") %>%
    tab_footnote("My Footnote 2", locations = cells_column_labels(1)) %>%
    tab_style(
      style = cell_text(color = "orange"),
      locations = cells_source_notes()
    ) %>%
    tab_style(
      style = cell_text(color = "purple"),
      locations = cells_footnotes()
    )

  if (!testthat::is_testing() & interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_meta_info <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header) %>%
    tail(3)

  ## header
  expect_equal(
    docx_table_meta_info %>% lapply(function(x) x %>% xml2::xml_find_all(".//w:t") %>% xml2::xml_text()),
    list(
      c("", "My Footnote"),
      c("1", "My Footnote 2"),
      c("My Source Note")
    )
  )

  expect_equal(
    lapply(docx_table_meta_info, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% lapply(function(y) {
        y %>% xml2::xml_find_all(".//w:color") %>% xml2::xml_attr(attr = "val")
      })}),
    list(
      list(c("A020F0", "A020F0")),
      list(c("A020F0", "A020F0")),
      list("FFA500")
    )
  )
})

test_that("tables with cell & text coloring can be added to a word doc - with summaries (grand/group)", {

  skip("Seems to be issues with gt v0.8.0, and is fixed in devel version point")

  check_suggests_xml()

  ## simple table
  gt_exibble_min <- exibble %>%
    dplyr::select(-c(fctr, date, time, datetime)) %>%
    gt(rowname_col = "row", groupname_col = "group") %>%
    summary_rows(
      columns = num,
      fns = list(
        avg = ~mean(., na.rm = TRUE),
        total = ~sum(., na.rm = TRUE),
        s.d. = ~sd(., na.rm = TRUE)
      ),
      formatter = fmt_number
    ) %>%
    grand_summary_rows(
      columns = num,
      fns = list(
        avg = ~mean(., na.rm = TRUE),
        total = ~sum(., na.rm = TRUE),
        s.d. = ~sd(., na.rm = TRUE)
      ),
      formatter = fmt_number
    ) %>%
    tab_style(
      style = cell_text(color = "orange"),
      locations = cells_summary(groups = "grp_a", columns = char)
    ) %>%
    tab_style(
      style = cell_text(color = "green"),
      locations = cells_stub_summary()
    ) %>%
    tab_style(
      style = cell_text(color = "purple"),
      locations = cells_grand_summary(columns = num, rows = 3)
    ) %>%
    tab_style(
      style = cell_fill(color = "yellow"),
      locations = cells_stub_grand_summary()
    )

  if (!testthat::is_testing() & interactive()) {
    print(gt_exibble_min)
  }

  ## Add table to empty word document
  word_doc <- officer::read_docx() %>%
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
  docx_contents <- docx$doc_obj$get() %>%
    xml2::xml_children() %>%
    xml2::xml_children()

  docx_table_body_header <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tblHeader/ancestor::w:tr")

  docx_table_body_contents <- docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr") %>%
    setdiff(docx_table_body_header)

  ## body text
  expect_equal(
    docx_table_body_contents %>% lapply(function(x) x %>% xml2::xml_find_all(".//w:t") %>% xml2::xml_text()),
    list(
      "grp_a",
      c("row_1", "1.111e-01", "apricot", "49.950"),
      c("row_2","2.222e+00", "banana", "17.950"),
      c("row_3", "3.333e+01", "coconut","1.390"),
      c("row_4", "4.444e+02", "durian", "65100.000"),
      c("avg","120.02", "—", "—"),
      c("total", "480.06", "—", "—"),
      c("s.d.","216.79", "—", "—"),
      "grp_b",
      c("row_5", "5.550e+03", "NA", "1325.810"),
      c("row_6", "NA", "fig", "13.255"),
      c("row_7", "7.770e+05","grapefruit", "NA"),
      c("row_8", "8.880e+06", "honeydew", "0.440"),
      c("avg", "3,220,850.00", "—", "—"),
      c("total", "9,662,550.00","—", "—"),
      c("s.d.", "4,916,123.25", "—", "—"),
      c("avg", "1,380,432.87","—", "—"),
      c("total", "9,663,030.06", "—", "—"),
      c("s.d.", "3,319,613.32","—", "—")
    )
  )

  ## the summaries for group a and b are green,
  ## the 2nd column of the group a summary is orange,
  ## the 1st col, 3rd value in the grand total is purple
  expect_equal(
    lapply(docx_table_body_contents, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% sapply(function(y) {
        val <- y %>% xml2::xml_find_all(".//w:color") %>% xml2::xml_attr(attr = "val")
        if (identical(val, character())) {
          ""
        }else{
          val
        }
      })}),
    list("",
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "",""),
         c("", "", "", ""),
         c("00FF00", "", "FFA500", ""),
         c("00FF00", "", "FFA500", ""),
         c("00FF00", "", "FFA500", ""),
         "",
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("00FF00", "", "", ""),
         c("00FF00", "", "", ""),
         c("00FF00", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "A020F0", "", "")
    )
  )

  ## the grand total row names fill is is yellow
  expect_equal(
    lapply(docx_table_body_contents, function(x) {
      x %>% xml2::xml_find_all(".//w:tc") %>% sapply(function(y) {
        val <- y %>% xml2::xml_find_all(".//w:shd") %>% xml2::xml_attr(attr = "fill")
        if (identical(val, character())) {
          ""
        }else{
          val
        }
      })}),
    list("",
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         "",
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("", "", "", ""),
         c("FFFF00", "", "", ""),
         c("FFFF00", "", "", ""),
         c("FFFF00", "", "", "")
    )
  )
})

test_that("tables preserves spaces in text & can be added to a word doc", {

  skip("Only done in devel gt at this point")

  check_suggests_xml()

  ## simple table
  gt_exibble <-
    exibble[1,1] %>%
    dplyr::mutate(
      `5 Spaces Before` = "     Preserve",
      `5 Spaces After` = "Preserve     ",
      `5 Spaces Before - preserve` = "     Preserve",
      `5 Spaces After - preserve` = "Preserve     ") %>%
    gt() %>%
    tab_style(
      style = cell_text(whitespace = "pre"),
      location = cells_body(columns = contains("preserve"))
    )

  ## Add table to empty word document
  word_doc_normal <-
    officer::read_docx() %>%
    body_add_gt(
      gt_exibble,
      align = "center"
    )

  ## save word doc to temporary file
  temp_word_file <- tempfile(fileext = ".docx")
  print(word_doc_normal,target = temp_word_file)

  ## Manual Review
  if (!testthat::is_testing() & interactive()) {
    shell.exec(temp_word_file)
  }

  ## Programmatic Review
  docx <- officer::read_docx(temp_word_file)

  ## get docx table contents
  docx_contents <- xml2::xml_children(xml2::xml_children(docx$doc_obj$get()))

  ## extract table contents
  docx_table_body_contents <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr")

  ## text is preserved
  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c("num","5 Spaces Before","5 Spaces After","5 Spaces Before - preserve","5 Spaces After - preserve"),
      c("0.1111","     Preserve","Preserve     ","     Preserve","Preserve     ")
    )
  )

  ## text "space" is set to preserve only for last 2 body cols
  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_attr((xml2::xml_find_all(x, ".//w:t")),"space")
    ),
    list(
      c("default", "default", "default", "default", "default"),
      c("default", "default", "default", "preserve","preserve")
    )
  )

})

test_that("tables respects column and cell alignment and can be added to a word doc", {

  skip("Only done in devel gt at this point")

  check_suggests_xml()

  ## simple table
  gt_exibble <-
    exibble[1:2,1:4] %>%
    `colnames<-`(c(
      "wide column number 1",
      "wide column number 2",
      "wide column number 3",
      "tcn4" #thin column number 4
    )) %>%
    gt() %>%
    cols_align(
      "right", columns = `wide column number 1`
    ) %>%
    cols_align(
      "left", columns = c(`wide column number 2`, `wide column number 3`)
    ) %>%
    tab_style(
      style = cell_text(align = "right"),
      location = cells_body(columns = c(`wide column number 2`, `wide column number 3`), rows = 2)
    ) %>%
    tab_style(
      style = cell_text(align = "left"),
      location = cells_body(columns = c(`wide column number 1`), rows = 2)
    ) %>%
    tab_style(
      cell_text(align = "left"),
      location = cells_column_labels(columns = c(tcn4))
    )

  ## Add table to empty word document
  word_doc <-
    officer::read_docx() %>%
    body_add_gt(
      gt_exibble,
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
  docx_table_body_contents <-
    docx_contents[1] %>%
    xml2::xml_find_all(".//w:tr")

  ## text is preserved
  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x) xml2::xml_text(xml2::xml_find_all(x, ".//w:p"))
    ),
    list(
      c(
        "wide column number 1",
        "wide column number 2",
        "wide column number 3",
        "tcn4"
      ),
      c("0.1111", "apricot", "one","2015-01-15"),
      c("2.2220", "banana", "two","2015-02-15")
    )
  )

  ## text "space" is set to preserve only for last 2 body cols
  expect_equal(
    lapply(
      docx_table_body_contents,
      FUN = function(x)
        x %>%
        xml2::xml_find_all(".//w:pPr") %>%
        lapply(FUN = function(y) xml2::xml_attr(xml2::xml_find_all(y,".//w:jc"),"val"))
    ),
    list(
      ## styling only on 4th column of header
      list(character(0), character(0), character(0), "start"),

      ## styling as applied or as default from gt
      list("end", "start", "start", "end"),
      list("start", "end", "end", "end"))
  )

})
