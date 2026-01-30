import pandas as pd
from pylatex import Document, Section, Subsection, Command, Figure, Tabular, NoEscape
from pylatex.utils import bold

# === 1. Read Excel data ===
# Use the Force Table Excel file from assets
try:
    df = pd.read_excel("assets/examples/Force Table.xlsx")
except FileNotFoundError:
    # Create sample data for demonstration
    df = pd.DataFrame({
        'Position': [0, 3, 6, 9],
        'Type': ['Support', 'Load', 'Load', 'Support'],
        'Value': [0, 500, 300, 0]
    })

# Creates LaTeX document 
doc = Document("beam_report")

# Add required packages for pgfplots and tikz to preamble
doc.preamble.append(NoEscape(r'\usepackage{pgfplots}'))
doc.preamble.append(NoEscape(r'\usepackage{tikz}'))
doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.16}'))

# Title page
doc.preamble.append(Command('title', 'Beam Analysis Report'))
doc.preamble.append(Command('author', 'PyLaTeX Report Generator'))
doc.preamble.append(Command('date', NoEscape(r'\today')))
doc.append(NoEscape(r'\maketitle'))
doc.append(NoEscape(r'\tableofcontents'))
doc.append(NoEscape(r'\newpage'))

# === 3. Introduction with beam image ===
with doc.create(Section('Introduction')):
    doc.append("This report presents a comprehensive analysis of beam loading and structural behavior.")
    with doc.create(Figure(position='h!')) as beam_fig:
        beam_fig.add_image("assets/beam.png", width=NoEscape(r'0.8\textwidth'))
        beam_fig.add_caption("Simply supported beam under loading")

# === 4. Input data table from Excel ===
doc.append(NoEscape(r'\newpage'))

with doc.create(Section('Load Data')):
    doc.append("Beam structural analysis data read from Excel file:")
    doc.append(NoEscape(r'\begin{center}'))
    with doc.create(Tabular('|c|c|c|')) as table:
        table.add_hline()
        table.add_row((bold("Position (x)"), bold("Shear Force"), bold("Bending Moment")))
        table.add_hline()
        for _, row in df.iterrows():
            table.add_row((row['x'], row['Shear force'], row['Bending Moment']))
            table.add_hline()
    doc.append(NoEscape(r'\end{center}'))

# === 5. Analysis: TikZ/pgfplots diagrams ===
with doc.create(Section('Analysis')):
    doc.append("Shear Force Diagram (SFD) and Bending Moment Diagram (BMD) analysis:")
    
    # Use data from Excel file
    x_data = df['x'].tolist()
    shear_data = df['Shear force'].tolist()
    moment_data = df['Bending Moment'].tolist()

    # Create Shear Force Diagram with pgfplots
    sfd_plot = NoEscape(r"""
\begin{center}
\begin{tikzpicture}
  \begin{axis}[
    title={Shear Force Diagram (SFD)},
    xlabel=Position x (m),
    ylabel=Shear Force (N),
    grid=both,
    width=10cm,
    height=6cm,
    legend pos=north east
  ]
  \addplot[color=blue, mark=*, line width=2pt] coordinates {
    """ + " ".join(f"({x:.1f},{s:.1f})" for x, s in zip(x_data, shear_data)) + r"""
  };
  \addlegendentry{Shear Force}
  \end{axis}
\end{tikzpicture}
\end{center}
    """)
    doc.append(sfd_plot)

    doc.append("")  # Add spacing
    
    # Create Bending Moment Diagram with pgfplots
    bmd_plot = NoEscape(r"""
\begin{center}
\begin{tikzpicture}
  \begin{axis}[
    title={Bending Moment Diagram (BMD)},
    xlabel=Position x (m),
    ylabel=Bending Moment (N·m),
    grid=both,
    width=10cm,
    height=6cm,
    legend pos=north east
  ]
  \addplot[color=red, mark=square*, line width=2pt] coordinates {
    """ + " ".join(f"({x:.1f},{m:.2f})" for x, m in zip(x_data, moment_data)) + r"""
  };
  \addlegendentry{Bending Moment}
  \end{axis}
\end{tikzpicture}
\end{center}
    """)
    doc.append(bmd_plot)

# === 6. Generate LaTeX and optionally PDF ===
# Generate LaTeX file
doc.generate_tex('beam_report')
print("✓ LaTeX file generated: beam_report.tex")

# Trying to generate PDF using pdflatex (ran twice for TOC)
try:
    import subprocess
    # First run to generate aux file
    result1 = subprocess.run(['pdflatex', '-interaction=nonstopmode', 'beam_report.tex'], 
                           capture_output=True, timeout=30)
    # Second run to populate TOC
    result2 = subprocess.run(['pdflatex', '-interaction=nonstopmode', 'beam_report.tex'], 
                           capture_output=True, timeout=30)
    
    import os
    if os.path.exists('beam_report.pdf'):
        print("✓ PDF generated: beam_report.pdf")
    else:
        print("⚠ PDF generation failed")
except Exception as e:
    print(f"⚠ PDF generation error: {str(e)}")
