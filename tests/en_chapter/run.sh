# 0. based 
pandoc \
    ./main_modified.tex \
    -o ./out.docx \
    --filter pandoc-crossref \
    --number-sections \
    -M autoEqnLabels \
    -M tableEqns \
    -M reference-section-title=Reference \
    --bibliography=../ref.bib \
    --citeproc \
    -t docx+native_numbering

# 1. chapter=false
pandoc \
    ./main_modified.tex \
    -o ./out1.docx \
    --filter pandoc-crossref \
    --number-sections \
    -M autoEqnLabels \
    -M tableEqns \
    -M reference-section-title=Reference \
    -M chapters=false \
    --top-level-division=section \
    --bibliography=../ref.bib \
    --citeproc \
    -t docx+native_numbering

# 2. chaptersDepth=0
pandoc \
    ./main_modified.tex \
    -o ./out2.docx \
    --filter pandoc-crossref \
    --number-sections \
    -M autoEqnLabels \
    -M tableEqns \
    -M reference-section-title=Reference \
    -M chaptersDepth=0 \
    --bibliography=../ref.bib \
    --citeproc \
    -t docx+native_numbering

# 3. sectionsDepth=-1
pandoc \
    ./main_modified.tex \
    -o ./out3.docx \
    --filter pandoc-crossref \
    --number-sections \
    -M autoEqnLabels \
    -M tableEqns \
    -M reference-section-title=Reference \
    -M sectionsDepth=-1 \
    --bibliography=../ref.bib \
    --citeproc \
    -t docx+native_numbering