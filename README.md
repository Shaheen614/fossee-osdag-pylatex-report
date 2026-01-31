Creation of Custom Latex Template using PyLatex

📑 Overview
This project demonstrates an automated workflow for generating a professional **Beam Analysis Report**.  
It reads input data from an Excel file, processes it with Python, and compiles the results into a LaTeX PDF with tables and diagrams.  
The workflow ensures **reproducibility, clarity, and reviewer‑friendly presentation**.


⚙️ Requirements
- Python 3.9+
- Pandas
- PyLaTeX
- MiKTeX (LaTeX distribution with pgfplots/tikz)

Install dependencies:
bash
pip install pandas pylatex

▶️ How to Run
Clone the repository:

bash
git clone https://github.com/<your-username>/fossee-osdag-pylatex-report.git
cd fossee-osdag-pylatex-report/src


Run the script:
bash
python main.py

Output:
beam_report.tex (LaTeX source)
beam_report.pdf (final report)

📑 Report Contents
Title page & Table of Contents
Introduction
Methodology
Load Data (Excel table)
Analysis (Shear Force & Bending Moment diagrams)
Discussion
Conclusion



