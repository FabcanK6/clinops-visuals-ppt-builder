# ClinOps Visuals PowerPoint Deck Builder

## Overview
This Python script automates the creation of a standardized PowerPoint deck from Power BI-exported PDFs. It converts each PDF page into an image, crops the top 10% to remove slicers/bookmarks, and organizes slides by study in a fixed order. The output file is dynamically named with the current date (e.g., `ClinOps_Visuals_PPT_Deck_09Jan2026.pptx`).

## Features
- Auto-renames raw downloaded files (e.g., `ClinOps_Visuals (1).pdf`) to expected names.
- Converts PDF pages to PNG at **200 DPI** for clarity.
- Crops **10% from the top** of each image to remove slicers/bookmarks.
- Creates a cover slide titled **Power BI ClinOps_Visuals Export**.
- Inserts `FullReport.pdf` pages first, then study sections in fixed order:
  - GD-2503 → MG → CIDP → SjD → RA → CLE
- Adds divider slides for each study.
- Saves the deck in `~/Downloads` with a date suffix.

## Requirements
- Python 3.x
- Packages:
  - [PyMuPDF](https://pypi.org/project/PyMuPDF/) (`fitz`)
  - [python-pptx](https://pypi.org/project/python-pptx/)
  - [Pillow](https://pypi.org/project/Pillow/)
