#bin/bash

pandoc main.tex -o main.docx \
  --filter pandoc-crossref \
  --reference-doc=../my_temp.docx \
  --number-sections \
  -M autoEqnLabels -M tableEqns \
  -M reference-section-title=Reference \
  --bibliography=ref.bib \
  --citeproc --csl ../ieee.csl