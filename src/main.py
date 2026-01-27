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

# === 2. Create LaTeX document ===
doc = Document("beam_report")

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
        beam_fig.add_image("src/assets/examples/beam.png", width=NoEscape(r'0.8\textwidth'))
        beam_fig.add_caption("Simply supported beam under loading")

# === 4. Input data table from Excel ===
with doc.create(Section('Load Data')):
    doc.append("Beam structural analysis data read from Excel file:")
    with doc.create(Tabular('|c|c|c|')) as table:
        table.add_hline()
        table.add_row((bold("Position (x)"), bold("Shear Force"), bold("Bending Moment")))
        table.add_hline()
        for _, row in df.iterrows():
            table.add_row((row['x'], row['Shear force'], row['Bending Moment']))
            table.add_hline()

# === 5. Analysis: TikZ/pgfplots diagrams ===
with doc.create(Section('Analysis')):
    doc.append("Shear Force Diagram (SFD) and Bending Moment Diagram (BMD) are plotted below using TikZ/pgfplots.")

    # Example shear force array (replace with computed values)
    x = [0, 2, 4, 6]
    shear = [0, -200, -200, 0]
    moment = [0, -400, -200, 0]

    # TikZ code for SFD
    tikz_sfd = r"""
    \begin{tikzpicture}
    \begin{axis}[xlabel={x}, ylabel={Shear Force}, grid=major]
    \addplot coordinates {
    """ + " ".join(f"({xi},{yi})" for xi, yi in zip(x, shear)) + r"""
    };
    \end{axis}
    \end{tikzpicture}
    """

    # TikZ code for BMD
    tikz_bmd = r"""
    \begin{tikzpicture}
    \begin{axis}[xlabel={x}, ylabel={Bending Moment}, grid=major]
    \addplot coordinates {
    """ + " ".join(f"({xi},{yi})" for xi, yi in zip(x, moment)) + r"""
    };
    \end{axis}
    \end{tikzpicture}
    """

    doc.append(NoEscape(tikz_sfd))
    doc.append(NoEscape(tikz_bmd))

# === 6. Generate LaTeX and optionally PDF ===
# Generate LaTeX file
doc.generate_tex('beam_report')
print("✓ LaTeX file generated: beam_report.tex")

# Try to generate PDF using pdflatex (doesn't require Perl)
try:
    import subprocess
    result = subprocess.run(['pdflatex', '-interaction=nonstopmode', 'beam_report.tex'], 
                          capture_output=True, timeout=30)
    # pdflatex returns non-zero even on success with warnings, so just check if PDF exists
    import os
    if os.path.exists('beam_report.pdf'):
        print("✓ PDF generated: beam_report.pdf")
    else:
        print("⚠ PDF generation failed")
except Exception as e:
    print(f"⚠ PDF generation error: {str(e)}")
