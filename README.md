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

## Automation with Power Automate
This script can be integrated with Microsoft Power Automate to enable scheduled or event-driven execution without manual intervention. Using Power Automate, you can:

Schedule Runs: Trigger the Python script at specific times (e.g., weekly or monthly) to generate the ClinOps Visuals deck automatically.
Cloud Execution: Use a Power Automate Desktop flow or a cloud flow with a connector to run the script on a designated machine.
File Handling:

Automatically download Power BI-exported PDFs to the target folder (e.g., ~/Downloads or a shared directory).
Move the generated PPTX to a SharePoint site or Teams channel for easy access.


Notifications: Send an email or Teams message when the deck is ready, including a link to the PPTX file.

Example Flow Steps

Trigger: Scheduled recurrence (e.g., every Friday at 2 PM).
Download PDFs: Use Power BI export actions or browser automation to save files.
Run Script: Execute build_clinops_deck.py on your local machine or a VM using Power Automate Desktop.
Post-Processing:

Rename or archive PDFs.
Upload the generated PPTX to SharePoint or Teams.

Notify Stakeholders: Send an email or Teams message with the PPTX link.

## Power Automate Integration Diagram

Below is an example of how the Power Automate flow integrates with this script:

![Power Automate Flow Diagram](assets/power_automate_flow.png)

### Flow Overview:
1. **Trigger**: Scheduled recurrence (e.g., every Friday at 2 PM).
2. **Download PDFs**: Export from Power BI and save to target folder.
3. **Run Script**: Execute `build_clinops_deck.py` on a local machine or VM.
4. **Upload Output**: Save generated PPTX to SharePoint or Teams.
5. **Notify Stakeholders**: Send email or Teams message with PPTX link.

### Benefits

Eliminates manual steps for building the deck.

Ensures consistency and timeliness for executive reporting.

Integrates seamlessly with existing Microsoft ecosystem tools.


