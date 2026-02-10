# PyLatex-Report--FOSSE

# 📊 Structural Beam Analysis Report Generator

## 📖 Project Overview

The **Structural Beam Analysis Report Generator** is a Python-based automation tool designed to generate professional engineering reports for **simply supported beams**.  
It bridges the gap between numerical structural analysis and high-quality technical documentation by reading load data from Excel files and producing well-formatted PDF reports using **PyLaTeX**.

The generated reports follow standard structural engineering practices and include **vector-based diagrams** that remain sharp and clear at any zoom level.

---

## ✨ Features

- ✅ **Automated Data Ingestion**  
  Reads beam loading positions and magnitudes directly from Excel spreadsheets using `pandas`.

- ✅ **LaTeX-Native Tables**  
  Converts load and force data into selectable and searchable LaTeX tabulars (no image-based tables).

- ✅ **Vector-Based Diagrams**  
  Generates **Shear Force Diagrams (SFD)** and **Bending Moment Diagrams (BMD)** using TikZ/pgfplots.

- ✅ **Professional Report Structure**  
  Automatically creates:
  - Title Page  
  - Table of Contents  
  - Introduction  
  - Analysis Section  
  - Summary  

- ✅ **Visual Integration**  
  Embeds reference beam diagrams directly into the report using PyLaTeX `Figure` environments.

---

## 🛠️ Technical Stack

- **Programming Language:** Python 3.10+  
- **Data Processing:** pandas, openpyxl  
- **PDF Generation:** PyLaTeX  
- **Graphics:** TikZ / pgfplots (LaTeX-native vector graphics)  
- **LaTeX Backend:** TeX Live / MiKTeX / MacTeX  

---

## 🚀 Quick Start

### 1️⃣ Prerequisites

Make sure a LaTeX distribution is installed on your system:

- **Windows:** MiKTeX  
- **Linux:** TeX Live  
- **macOS:** MacTeX  

This is required to compile the `.tex` file into a PDF report.

---

### 2️⃣ Installation

Clone the repository and install the required dependencies:

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME
pip install -r requirements.txt
