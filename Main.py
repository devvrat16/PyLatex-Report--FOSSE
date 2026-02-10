import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Figure, Tabular, Command, Package
from pylatex.utils import NoEscape, italic
import os

def generate_beam_report():
    print("📊 Reading beam data from Excel...")
    df = pd.read_excel('beam_data.xlsx')
    
    # Standard article class with 12pt for better readability
    doc = Document(documentclass='article', document_options=['a4paper', '12pt'])
    
    # --- Professional Packages ---
    doc.packages.append(Package('xcolor'))
    doc.packages.append(Package('charter'))
    doc.packages.append(Package('tikz'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(Package('geometry', options=['margin=1in']))
    doc.packages.append(Package('fancyhdr'))

    # --- Custom Colors & Formatting ---
    doc.preamble.append(NoEscape(r'\definecolor{midnight}{HTML}{2C3E50}'))
    doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.18}'))

    # --- Page 1: Cover Page ---
    doc.append(NoEscape(r'\begin{titlepage}'))
    doc.append(NoEscape(r'\centering'))
    doc.append(NoEscape(r'\vspace*{3cm}'))
    doc.append(NoEscape(r'{\Huge \textbf{\color{midnight}Structural Analysis Report} \par}'))
    doc.append(NoEscape(r'\vspace{1cm}'))
    doc.append(NoEscape(r'{\Large \textbf{Project:} Simply Supported Beam Simulation \par}'))
    doc.append(NoEscape(r'\vspace{3cm}'))
    doc.append(NoEscape(r'{\large \textbf{Prepared By:} \par}'))
    doc.append(NoEscape(r'{\Large Dev vrat \par}'))
    doc.append(NoEscape(r'{\large Registration Number: 23BAI10095 \par}'))
    doc.append(NoEscape(r'\vfill'))
    doc.append(NoEscape(r'{\large \today \par}'))
    doc.append(NoEscape(r'\end{titlepage}'))

    # --- Page 2 Setup ---
    doc.append(NoEscape(r'\newpage'))
    doc.preamble.append(NoEscape(r'\pagestyle{fancy}'))
    doc.preamble.append(NoEscape(r'\fancyhf{}'))
    doc.preamble.append(NoEscape(r'\lhead{\small \color{gray} Dev vrat | 23BAI10095}'))
    doc.preamble.append(NoEscape(r'\rhead{\small \color{gray} Page \thepage}'))

    # Section 1: Introduction & Theory
    with doc.create(Section('Structural Overview')):
        doc.append('This report analyzes the internal stress distribution of a 12-meter structural element. The assessment focuses on Shear Force (V) and Bending Moment (M) envelopes.')
        
        with doc.create(Subsection('Analysis Assumptions')):
            doc.append('1. Material behaves in a linear-elastic manner.')
            doc.append(NoEscape(r'\\'))
            doc.append('2. Plane sections remain plane after bending.')
            doc.append(NoEscape(r'\\'))
            doc.append('3. Self-weight is included in the provided loading data.')

        with doc.create(Figure(position='h!')) as beam_fig:
            if os.path.exists('beam.png'):
                beam_fig.add_image('beam.png', width=NoEscape(r'0.6\textwidth'))
            beam_fig.add_caption('Schematic Loading Diagram')

    # Section 2: Numerical Results
    with doc.create(Section('Tabulated Data Log')):
        doc.append('Extracted values at discrete intervals across the span:')
        doc.append(NoEscape(r'\vspace{0.4cm}'))
        
        with doc.create(Tabular('|c|c|c|', row_height=1.3)) as table:
            table.add_hline()
            table.add_row(('Point (m)', 'Shear (kN)', 'Moment (kNm)'))
            table.add_hline()
            for _, row in df.iloc[::2].iterrows():
                table.add_row((f"{row['X']:.1f}", f"{row['Shear force']:.1f}", f"{row['Bending Moment']:.1f}"))
            table.add_hline()

    # Section 3: Visual Distribution
    with doc.create(Section('Internal Force Envelopes')):
        doc.append('The following diagrams represent the calculated envelopes for the beam span.')
        
        with doc.create(Subsection('Shear Force Diagram (SFD)')):
            doc.append(NoEscape(generate_sfd_plot(df)))

        with doc.create(Subsection('Bending Moment Diagram (BMD)')):
            doc.append(NoEscape(generate_bmd_plot(df)))

    # Generate PDF
    print("📄 Generating 2-Page PDF...")
    doc.generate_pdf('Dev_Vrat_Beam_Report', clean_tex=False, compiler='pdflatex')
    print("✅ Success: Dev_Vrat_Beam_Report.pdf")

def generate_sfd_plot(df):
    coords = "".join([f"({r['X']}, {r['Shear force']}) " for _, r in df.iterrows()])
    return r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[width=11cm, height=4.5cm, axis lines=middle, xlabel=$x$, ylabel=$V$, 
        grid=both, title=\textbf{\small Shear Distribution}]
    \addplot[thick, midnight, fill=midnight, fill opacity=0.1] coordinates {""" + coords + r"""} \closedcycle;
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """

def generate_bmd_plot(df):
    coords = "".join([f"({r['X']}, {r['Bending Moment']}) " for _, r in df.iterrows()])
    return r"""
    \begin{center}
    \begin{tikzpicture}
    \begin{axis}[width=11cm, height=4.5cm, axis lines=middle, xlabel=$x$, ylabel=$M$, 
        grid=both, title=\textbf{\small Flexural Distribution}]
    \addplot[thick, red!70!black, fill=red, fill opacity=0.05] coordinates {""" + coords + r"""} \closedcycle;
    \end{axis}
    \end{tikzpicture}
    \end{center}
    """

if __name__ == "__main__":
    generate_beam_report()