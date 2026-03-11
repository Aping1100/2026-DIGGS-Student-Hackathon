# Dalah San Beta Submission

## Team Members
- Irene Peng

## Challenge Theme(s)
Direct Design/Interpretation: Develop tools that use DIGGS data for engineering analysis, design recommendations, or automated interpretations.

## Project Description
Dalah San is a geotechnical engineering web application that integrates DIGGS XML data into practical engineering workflows. The project focuses on transforming DIGGS-structured subsurface data into engineering calculations, design checks, and interpretation-ready outputs for field and design decisions.

Current completed modules used for submission:
- Liquefaction analysis
- Shallow foundation analysis
- Deep excavation safety checks (uplift and sand boil)

The application includes DIGGS import/processing, engineering parameter handling, calculation pipelines, and result visualization to support applied geotechnical practice.

## Technologies Used
- Python
- Flask
- NumPy
- Pandas
- Matplotlib
- OpenPyXL / XlsxWriter
- HTML/CSS/JavaScript
- DIGGS XML data model

## Setup Instructions


## Demo



## Repository Structure

<pre>
dalah_san/
├── README.md
├── LICENSE
├── docs/              # Documentation
├── demo/              # Demo materials
└── src/               # Application source
    ├── app.py         # Flask entry
    ├── routes/        # API routes (feedback, diggs, excavation, shallow, etc.)
    ├── templates/     # index.html (single-page UI)
    ├── static/        # CSS, JS, images, Leaflet
    ├── utils/         # Helpers
    ├── tools/         # Preprocessing scripts
    └── *.py           # Module logic (liquefaction, shallow_foundation, excavation, diggs_db, …)
</pre>


