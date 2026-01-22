# PyLaTeX Beam Report Generator

This project generates a professional engineering PDF report using **PyLaTeX**, based on beam loading data provided in an Excel file.

## Features
- Reads force/load data from Excel using **pandas**.
- Computes reactions, shear force, and bending moment for a simply supported beam.
- Recreates the Excel table in LaTeX (selectable text, not image).
- Embeds a beam image in the introduction.
- Generates **Shear Force Diagram (SFD)** and **Bending Moment Diagram (BMD)** using **TikZ/pgfplots**.
- Produces a structured PDF report with title page, table of contents, and analysis sections.

## Requirements
- Python 3.9+
- LaTeX distribution (TeX Live / MiKTeX)
- Python packages:
  ```bash
  pip install -r requirements.txt

Usage
Clone the repository:

bash
git clone https://github.com/<your-username>/beam-report-generator.git
cd beam-report-generator
Run the script:

bash
python src/main.py --excel examples/loads.xlsx --beam_img assets/beam.png --length 6 --out beam_report
The generated PDF (beam_report.pdf) will appear in the project folder.

Project Structure
Code
beam-report-generator/
├── src/              # Source code
├── assets/           # Beam image
├── examples/        
├── README.md
├── requirements.txt
└── LICENSE
License
This work is licensed under Creative Commons Zero v1.0 Universal (CC0)
