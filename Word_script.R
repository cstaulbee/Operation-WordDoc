library(dplyr)
library(tidyr)
library(docxtractr)


filenames <- list.files(".", pattern="*.docx", full.names=TRUE)
docx.files <- lapply(filenames, function(file) read_docx(file))

idx <- 1
docx.tables <- lapply(docx.files, function(file) {
  
  ifelse(dir.exists("Contents"), {
    unlink("Contents", recursive=T, force=T)
    dir.create("Contents")
  }, {
    dir.create("Contents")
  })
  
  filename <- filenames[idx]
  idx <- idx + 1
  
  tbl <- docx_extract_tbl(file, 1)
  file.copy(filename, "Contents\\word.zip", overwrite=T)
  unzip("Contents\\word.zip", exdir='Contents')
  x <- xml2::read_xml("Contents\\word\\document.xml")
  nodes <- xml2::xml_find_all(x, "w:body/w:p/w:r/w:t")
  data.date <- paste(xml2::xml_text(nodes, trim=T), collapse="::")
  word_df <- strsplit(gsub("[:]{1,}", ":", txt), ":")
  return(
    list(
      date=data.date
    )
  )
})

word_df <- strsplit(gsub("[:]{1,}", ":", docx.tables), ":")