
# gto <img src="man/figures/logo.png" align="right" height="120" />

<!-- badges: start -->
<!-- badges: end -->

The goal of gto is to provide the tools to allow users to insert gt tables into officeverse. Right now the only supported output is word.

## Installation

You can install the development version of gto like so:

``` r
remotes::install_github("GSK-Biostatistics/gto")
```

## Example

``` r

#load officer and gt
library(officer)
library(gt)

## create simple gt table
gt_tbl <- gt(head(exibble))

## Create docx and add gt table
doc <- read_docx()
doc <- body_add_gt(doc, value = gt_tbl)

## Save docx
fileout <- tempfile(fileext = ".docx")
print(doc, target = fileout)
```

