import pandas as pd
from pylatex import Document, Section, Command, Figure, Tabular, NoEscape
from pylatex.utils import bold
import subprocess, os

# === 1. Absolute paths to files ===
excel_path = r"C:\Users\DELL\Osdag\fossee-osdag-pylatex-report\src\assets\examples\Force Table.xlsx"
image_path = r"C:\Users\DELL\Osdag\fossee-osdag-pylatex-report\src\assets\sample_beam.png"

# === 2. Read Excel data ===
if os.path.exists(excel_path):
    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip()  # clean headers
    print("Excel headers:", df.columns.tolist())
else:
    print("Excel file not found at:", excel_path)
    df = pd.DataFrame({
        'x': [0.0, 1.5, 3.0],
        'Shear force': [45.0, 36.0, 27.0],
        'Bending Moment': [0.0, 60.75, 108.0]
    })

# === 3. Create LaTeX document ===
doc = Document("beam_report")
doc.preamble.append(NoEscape(r'\usepackage{pgfplots}'))
doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.18}'))

# Title page
doc.preamble.append(Command('title', 'Beam Analysis Report'))
doc.preamble.append(Command('author', 'PyLaTeX Report Generator'))
doc.preamble.append(Command('date', NoEscape(r'\today')))
doc.append(NoEscape(r'\maketitle'))
doc.append(NoEscape(r'\tableofcontents'))
doc.append(NoEscape(r'\newpage'))

# === 4. Introduction with beam image ===
with doc.create(Section('Introduction')):
    doc.append("This report presents a comprehensive analysis of beam loading and structural behavior.")
    if os.path.exists(image_path):
        with doc.create(Figure(position='h!')) as beam_fig:
            beam_fig.add_image(image_path, width=NoEscape(r'0.8\textwidth'))
            beam_fig.add_caption("Simply supported beam under loading")
    else:
        doc.append("Beam image not found at: " + image_path)

# === 5. Input data table ===
with doc.create(Section('Load Data')):
    with doc.create(Tabular('|c|c|c|')) as table:
        table.add_hline()
        table.add_row((bold("x (m)"), bold("Shear Force (N)"), bold("Bending Moment (Nm)")))
        table.add_hline()
        for _, row in df.iterrows():
            table.add_row((row['x'], row['Shear force'], row['Bending Moment']))
            table.add_hline()

# === 6. Analysis diagrams ===
with doc.create(Section('Analysis')):
    x_data = df['x'].tolist()
    shear_data = df['Shear force'].tolist()
    moment_data = df['Bending Moment'].tolist()

    sfd = r"""
\begin{center}
\begin{tikzpicture}
  \begin{axis}[title={Shear Force Diagram},xlabel=Position (m),ylabel=Shear Force (N),grid=both]
    \addplot[color=blue,mark=*] coordinates {
    """ + " ".join(f"({x},{s})" for x, s in zip(x_data, shear_data)) + r"""
    };
  \end{axis}
\end{tikzpicture}
\end{center}
"""
    doc.append(NoEscape(sfd))

    bmd = r"""
\begin{center}
\begin{tikzpicture}
  \begin{axis}[title={Bending Moment Diagram},xlabel=Position (m),ylabel=Bending Moment (Nm),grid=both]
    \addplot[color=red,mark=square*] coordinates {
    """ + " ".join(f"({x},{m})" for x, m in zip(x_data, moment_data)) + r"""
    };
  \end{axis}
\end{tikzpicture}
\end{center}
"""
    doc.append(NoEscape(bmd))

# === 7. Generate LaTeX and PDF ===
doc.generate_tex('beam_report')
print("✓ LaTeX file generated: beam_report.tex")

try:
    subprocess.run(['pdflatex', '-interaction=nonstopmode', 'beam_report.tex'], check=True)
    subprocess.run(['pdflatex', '-interaction=nonstopmode', 'beam_report.tex'], check=True)
    if os.path.exists('beam_report.pdf'):
        print("✓ PDF generated: beam_report.pdf")
    else:
        print("PDF generation failed")
except Exception as e:
    print("PDF generation error:", e)

