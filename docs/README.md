# Document compilation script
This repository contains a Python script (`texCompiler.py`) that automates the compilation

## Requirements
- [Python](https://www.python.org/downloads/) 3.8 or superior
- [TexWorks](https://tug.org/texworks/) or any $\LaTeX$ editor
- [Biber](https://www.ctan.org/pkg/biber)

## Usage
```bash
: '
python texCompiler.py [options] [folder] [main.tex] [output.pdf]

Options:
  -F <folder>      Project folder (alternative to positional)
  -o <file.pdf>    Output PDF path / name
  -I bib           Ignore bibliography (skip biber)
  -j <N>           Parallel worker threads  (default: CPU count)
  -f / --force     Force recompilation even if cache is fresh
  --no-inject      Skip APA preamble injection
  --no-preprocess  Skip R-style citation preprocessing
'

# To compile and open the PDF output:
cd docs
python texCompiler.py -F documento main.tex -o main.pdf
cd output
.\main.pdf

# To compile without bibliography (faster, for drafts):
cd docs
python texCompiler.py -F documento main.tex -o main.pdf -I bib
```