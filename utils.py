"""Utility functions for financial reporting."""

import pandas as pd


def get_data(url: str) -> pd.DataFrame:
    """Récupère le jeu de données en format Parquet à partir de l'URL indiqué"""
    return pd.read_parquet(url)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Applique une série de traitement sur le dataframe

    Liste des traitements effectués :
        1. Remplacer les scores NA par "S" afin de d'identifier les clients non-scorés
        2. Remplacer les scores_prev NA par N afin d'identifier les nouveaux clients
        3. Remplacer les matricules des agents par un libellé "MANUEL" pour mieux les distinguer

        Retourne le DataFrame
    """
    df = df.assign(
        score=df["score"].fillna("S"),
        score_prev=df["score_prev"].fillna("N"),
        id_agent=df["id_agent"].where(df["id_agent"] == "AUTO", "MANUEL"),
    )
    return df


def fill_indicators(ws, indicators: list[dict], data_sheet: str = "DATA") -> None:
    """
    Remplit les indicateurs dans une feuille Excel avec des formules.

    Args:
        ws: Feuille openpyxl cible (wb['Indicateurs']).
        indicators: Liste des indicateurs à insérer, définie dans config.py.
        data_sheet: Nom de la feuille de données. Défaut : 'DATA'.
    """

    def formula_countif(col: str, val: str) -> str:
        return f'=COUNTIF({data_sheet}!{col}:{col}, "{val}")'

    def formula_countifs(pairs: list[tuple[str, str]]) -> str:
        conditions = ", ".join(
            f'{data_sheet}!{col}:{col}, "{val}"' for col, val in pairs
        )
        return f"=COUNTIFS({conditions})"

    def formula_sum(range_str: str) -> str:
        return f"=SUM({range_str})"

    for item in indicators:
        cell = f"E{item['row']}"
        if item["formule"] == "COUNTIF":
            ws[cell] = formula_countif(*item["args"][0])
        elif item["formule"] == "COUNTIFS":
            ws[cell] = formula_countifs(item["args"])
        elif item["formule"] == "SUM":
            ws[cell] = formula_sum(item["args"])
        else:
            raise ValueError(f"Formule inconnue : {item['formule']}")
