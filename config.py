"""Configuration file for the project."""

from datetime import datetime
from pathlib import Path

# Lien vers les données à récupérer
URL = "https://minio.lab.sspcloud.fr/fabienhos/td-reporting-financial/financial_data.parquet"

INDICATORS = [
    # Répartition PP/PM
    {"row": 8, "formule": "COUNTIF", "args": [("B", "PP")]},
    {"row": 9, "formule": "COUNTIF", "args": [("B", "PM")]},
    {"row": 10, "formule": "SUM", "args": "E8:E9"},
    # Scores V/O/R/S
    {"row": 14, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "V")]},
    {"row": 15, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "O")]},
    {"row": 16, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "R")]},
    {"row": 17, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "S")]},
    {"row": 18, "formule": "SUM", "args": "E14:E17"},
    # DRC Complet
    {"row": 22, "formule": "COUNTIFS", "args": [("B", "PP"), ("G", "VRAI")]},
    {"row": 23, "formule": "COUNTIFS", "args": [("B", "PM"), ("G", "VRAI")]},
    {"row": 24, "formule": "SUM", "args": "E22:E23"},
    # Focus V/O
    {"row": 28, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "V")]},
    {"row": 29, "formule": "COUNTIFS", "args": [("B", "PM"), ("D", "V")]},
    {"row": 30, "formule": "SUM", "args": "E28:E29"},
    {"row": 31, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "O")]},
    {"row": 32, "formule": "COUNTIFS", "args": [("B", "PM"), ("D", "O")]},
    {"row": 33, "formule": "SUM", "args": "E31:E32"},
    # Focus V/O avec DRC complet
    {
        "row": 34,
        "formule": "COUNTIFS",
        "args": [("B", "PP"), ("D", "V"), ("G", "VRAI")],
    },
    {
        "row": 35,
        "formule": "COUNTIFS",
        "args": [("B", "PM"), ("D", "V"), ("G", "VRAI")],
    },
    {"row": 36, "formule": "SUM", "args": "E34:E35"},
    {
        "row": 37,
        "formule": "COUNTIFS",
        "args": [("B", "PP"), ("D", "O"), ("G", "VRAI")],
    },
    {
        "row": 38,
        "formule": "COUNTIFS",
        "args": [("B", "PM"), ("D", "O"), ("G", "VRAI")],
    },
    {"row": 39, "formule": "SUM", "args": "E37:E38"},
    # Focus R
    {"row": 43, "formule": "COUNTIFS", "args": [("B", "PP"), ("D", "R")]},
    {"row": 44, "formule": "COUNTIFS", "args": [("B", "PM"), ("D", "R")]},
    {"row": 45, "formule": "SUM", "args": "E43:E44"},
    {"row": 46, "formule": "COUNTIFS", "args": [("D", "R"), ("F", "AUTO")]},
    {"row": 47, "formule": "COUNTIFS", "args": [("D", "R"), ("F", "MANUEL")]},
    {"row": 48, "formule": "SUM", "args": "E46:E47"},
    # R avec DRC complet
    {
        "row": 49,
        "formule": "COUNTIFS",
        "args": [("B", "PP"), ("D", "R"), ("G", "VRAI")],
    },
    {
        "row": 50,
        "formule": "COUNTIFS",
        "args": [("B", "PM"), ("D", "R"), ("G", "VRAI")],
    },
    {"row": 51, "formule": "SUM", "args": "E49:E50"},
    # Nouveaux clients (score_prev = N)
    {
        "row": 52,
        "formule": "COUNTIFS",
        "args": [("B", "PP"), ("E", "N"), ("G", "VRAI")],
    },
    {
        "row": 53,
        "formule": "COUNTIFS",
        "args": [("B", "PM"), ("E", "N"), ("G", "VRAI")],
    },
    {"row": 54, "formule": "SUM", "args": "E52:E53"},
]

SHEET_INDICATORS = "Indicateurs"

# Chemins des fichiers Excel
today = datetime.today().strftime("%Y-%m-%d")
PATH_TEMPLATE = Path("template/template_reporting.xlsx")
PATH_OUTPUT = Path(f"output/Reporting_Financier_{today}.xlsx")
