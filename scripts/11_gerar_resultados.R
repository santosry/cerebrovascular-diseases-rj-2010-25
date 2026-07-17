source("scripts/_bootstrap.R", encoding = "UTF-8")
record_session_info()
if (!requireNamespace("rmarkdown", quietly = TRUE)) stop("Pacote rmarkdown ausente.")
# O manuscrito (artigo_morbimortalidade_avc_rbn_2026.docx) passou a ser mantido
# como DOCX fora do repositorio publico; analysis/artigo.Rmd foi descontinuado
# em 2026-07-17 e seu conteudo incorporado ao manuscrito. Ver CHANGELOG.md.
rmarkdown::render(project_path("analysis", "analises_robustas.Rmd"),
  output_dir = project_path("results"),
  quiet = FALSE, envir = new.env(parent = globalenv())
)
