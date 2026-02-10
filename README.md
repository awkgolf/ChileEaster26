⛏️ Geological Field Journal Automation (Chile 2026)
This project is a specialized Node.js pipeline designed to transform structured geological field observations into a professionally formatted, academic-style Word document.

📋 Table of Contents
Overview

System Architecture

Setup & Installation

Data Entry Guide

Field Troubleshooting

🔍 Overview
The system automates the tedious parts of field reporting—formatting, image placement, and indexing—so you can focus on the geology. It features a Photo Audit system that verifies your stratigraphic evidence is present before building the final report.

🏗 System Architecture
The project is built to be lightweight and portable:

Logic: Node.js using the docx library.

Data Strata: travelData.json (The single source of truth).

Assets: Organized /photos directory with support for multiple images per day.

Outputs: Professional .docx with automated Table of Contents and a Stratigraphic Index.

🚀 Setup & Installation
Windows Environment
Install Node.js.

Clone this repository.

Run npm install to build your local node_modules.

Run build_journal.bat or press F5 in VS Code to generate your journal.

Android / Samsung Tablet (Termux)
Install Termux and the Acode editor.

Navigate to the project folder.

Run termux-setup-storage.

Execute ./setup-tablet.sh to install the environment.

Build using node index.js.

📝 Data Entry Guide
Update your observations in travelData.json. The schema supports technical field notes and multiple annotated images:

JSON
{
  "day": "Day 12",
  "title": "El Tatio Geyser Field",
  "images": [
    { "url": "tatio_1.jpg", "caption": "Fig 12.1: Sinter terraces." }
  ],
  "geoNote": {
    "title": "Geothermal Precipitation",
    "text": "High silica content noted in active discharge zones."
  }
}
⚠️ Field Troubleshooting
JSON Syntax: Use the "Cheat Sheet" in the project folder if the script crashes due to a missing comma.

Photo Audit: If the terminal warns of a missing image, check that the filename in the JSON matches the file in /photos exactly (including case).

File Locks: Ensure the generated Word document is closed before running a new build.

📚 Technical Context
This project documents the tectonic and volcanic evolution of the South American Plate, focusing on the Andean orogeny and the unique volcanic history of Rapa Nui.# legendary-chile-memory
