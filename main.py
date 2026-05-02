# main.py
"""Orchestre la génération du reporting."""

from openpyxl import load_workbook
import pandas as pd

from utils import get_data, clean_data, fill_indicators
from config import URL, PATH_TEMPLATE, INDICATORS, SHEET_INDICATORS, PATH_OUTPUT


def main():
    # 1. Récupérer les données
    df = get_data(URL)

    # 2. Appliquer des traitements
    df = clean_data(df)

    # 3. Renseigner les données dans le `template`
    with pd.ExcelWriter(
        PATH_TEMPLATE, mode="a", engine="openpyxl", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name="DATA", index=False)

    wb = load_workbook(PATH_TEMPLATE)

    # 4. Renseigner les indicateurs
    fill_indicators(wb[SHEET_INDICATORS], indicators=INDICATORS)

    # 5. Sauvegarder et fermer le fichier de sortie
    # Le template sert de modèle (comme son nom l'indique...)
    # on écrit bien dans un fichier différent pour laisser le template vide de toutes données.
    wb.save(PATH_OUTPUT)
    wb.close()


if __name__ == "__main__":
    main()
