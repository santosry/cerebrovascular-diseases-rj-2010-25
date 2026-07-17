source("scripts/_bootstrap.R", encoding = "UTF-8")
record_session_info()
if (!requireNamespace("rmarkdown", quietly = TRUE)) stop("Pacote rmarkdown ausente.")
rmarkdown::render(
  project_path("analysis", "artigo.Rmd"),
  output_format = "word_document",
  output_file = "manuscrito_RECORD.docx",
  output_dir = project_path("results"),
  quiet = FALSE, envir = new.env(parent = globalenv())
)
rmarkdown::render(
  project_path("analysis", "artigo.Rmd"),
  output_format = "html_document",
  output_file = "manuscrito_RECORD.html",
  output_dir = project_path("results"),
  quiet = FALSE, envir = new.env(parent = globalenv())
)
rmarkdown::render(project_path("analysis", "analises_robustas.Rmd"),
  output_dir = project_path("results"),
  quiet = FALSE, envir = new.env(parent = globalenv())
)
