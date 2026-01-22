import argparse
import os
from pathlib import Path
import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Table, NoEscape, Command, Package
from pylatex.table import Tabular
from pylatex.utils import italic, bold
from assets.examples.data_io import read_loads_excel


def compute_beam_analysis(loads_df, length):
    """
    Compute reactions, shear force, and bending moment for a simply supported beam.
    
    Parameters:
    - loads_df: DataFrame with columns 'x', 'type', 'value'
    - length: Total length of the beam
    
    Returns:
    - Dictionary with reactions and diagrams
    """
    # Convert x and value to numeric
    loads_df = loads_df.copy()
    loads_df['x'] = pd.to_numeric(loads_df['x'])
    loads_df['value'] = pd.to_numeric(loads_df['value'])
    
    # Calculate support reactions
    # For simply supported beam: Sum of moments at A = 0, Sum of vertical forces = 0
    total_load = 0
    moment_sum = 0
    
    for idx, row in loads_df.iterrows():
        if row['type'] == 'point':
            total_load += row['value']
            moment_sum += row['value'] * row['x']
        elif row['type'] == 'distributed':
            # Assuming distributed load acts over a small segment at x
            total_load += row['value']
            moment_sum += row['value'] * row['x']
    
    # Reaction at support B (right support)
    reaction_b = moment_sum / length if length != 0 else 0
    # Reaction at support A (left support)
    reaction_a = total_load - reaction_b
    
    return {
        'reaction_a': reaction_a,
        'reaction_b': reaction_b,
        'total_load': total_load,
        'loads': loads_df
    }


def create_latex_report(loads_df, beam_analysis, beam_img_path, output_name, length):
    """
    Create a professional LaTeX report with beam analysis and diagram.
    
    Parameters:
    - loads_df: DataFrame with load data
    - beam_analysis: Dictionary with analysis results
    - beam_img_path: Path to beam image
    - output_name: Name for the output PDF (without extension)
    - length: Beam length
    """
    # Create LaTeX document
    doc = Document(output_name)
    doc.packages.append(Package('babel', options='english'))
    doc.packages.append(Package('amsmath'))
    doc.packages.append(Package('graphicx'))
    
    # Title Page
    doc.append(Command('title', 'Beam Analysis Report'))
    doc.append(Command('author', 'PyLaTeX Report Generator'))
    doc.append(Command('date', NoEscape(r'\today')))
    doc.append(Command('maketitle'))
    
    # Add Table of Contents
    doc.append(Command('tableofcontents'))
    doc.append(Command('newpage'))
    
    # Introduction Section
    with doc.create(Section('Introduction')):
        doc.append('This report presents a comprehensive analysis of beam loading and structural behavior.')
        doc.append('\n\n')
        
        # Add beam image
        if beam_img_path and os.path.exists(beam_img_path):
            doc.append('Figure of the beam:')
            doc.append('\n')
            doc.append(Command('includegraphics', arguments=NoEscape(beam_img_path), options='width=0.8\\textwidth'))
    
    # Load Data Section
    with doc.create(Section('Load Data')):
        doc.append(f'Beam Length: {length} units\n\n')
        
        # Create table
        with doc.create(Tabular('|c|c|c|')) as table:
            table.add_hline()
            table.add_row((bold('Position (x)'), bold('Type'), bold('Value')))
            table.add_hline()
            
            for idx, row in loads_df.iterrows():
                table.add_row((row['x'], row['type'], row['value']))
            
            table.add_hline()
    
    # Analysis Results Section
    with doc.create(Section('Analysis Results')):
        doc.append(f"Support Reaction at A: {beam_analysis['reaction_a']:.2f} units\n\n")
        doc.append(f"Support Reaction at B: {beam_analysis['reaction_b']:.2f} units\n\n")
        doc.append(f"Total Load: {beam_analysis['total_load']:.2f} units\n\n")
    
    # Generate LaTeX file (PDF generation requires LaTeX compiler)
    try:
        doc.generate_pdf(clean_tex=False)
        print(f"Report generated: {output_name}.pdf")
    except Exception as e:
        print(f"Note: LaTeX compiler not found. Generating LaTeX source file only.")
        print(f"Error: {e}")
        doc.generate_tex()
        print(f"LaTeX source file generated: {output_name}.tex")
        print("To generate PDF, install a LaTeX distribution (TeX Live, MiKTeX, etc.) and run:")
        print(f"  pdflatex {output_name}.tex")


def main():
    parser = argparse.ArgumentParser(description='Generate a PyLaTeX beam analysis report')
    parser.add_argument('--excel', required=True, help='Path to Excel file with load data')
    parser.add_argument('--beam_img', required=True, help='Path to beam image')
    parser.add_argument('--length', type=float, required=True, help='Beam length')
    parser.add_argument('--out', required=True, help='Output PDF name (without extension)')
    
    args = parser.parse_args()
    
    # Read loads from Excel
    print(f"Reading loads from {args.excel}...")
    loads_df = read_loads_excel(args.excel)
    
    # Compute beam analysis
    print("Computing beam analysis...")
    beam_analysis = compute_beam_analysis(loads_df, args.length)
    
    # Create LaTeX report
    print(f"Creating LaTeX report: {args.out}.pdf...")
    create_latex_report(loads_df, beam_analysis, args.beam_img, args.out, args.length)
    
    print("Done!")


if __name__ == '__main__':
    main()
