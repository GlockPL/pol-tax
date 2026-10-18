"""
How statement tables are presented - shared by the window and the Excel workbook export:
position hierarchy (A > A.I > A.I.1) and year-over-year change columns.
"""
from decimal import Decimal, InvalidOperation

import pandas as pd

# Sections with a current and a previous year column, where a year-over-year change makes sense
YEAR_SECTIONS = ("assets", "passives", "gain_and_loses", "change_in_capital", "capital_flow")

CHANGE_COLUMN = "Zmiana r/r"
CHANGE_PERCENT_COLUMN = "Zmiana r/r %"


class Percent(Decimal):
    """
    A percentage, kept apart from amounts so it is formatted differently
    """


def format_amount(value):
    return format(value, ',.2f').replace(',', ' ').replace('.', ',')


def format_percent(value):
    return f"{value:+.1f}%".replace('.', ',')


def short_year_headers(keywords, date_to):
    """
    'Kwota na dzień kończący bieżący rok obrotowy' -> 'Rok bieżący (2021)'
    """
    year = (date_to or "")[:4]
    current = f"Rok bieżący ({year})" if year.isdigit() else "Rok bieżący"
    previous = f"Rok poprzedni ({int(year) - 1})" if year.isdigit() else "Rok poprzedni"
    headers = {}
    if keywords.get("KwotaA"):
        headers[keywords["KwotaA"]] = current
    if keywords.get("KwotaB"):
        headers[keywords["KwotaB"]] = previous
    return headers


def parent_key(key, keys):
    """
    Closest existing ancestor of a position: 'A.II.1' -> 'A.II' -> 'A' -> '' (the total), 'A_RB' -> 'A'.
    Returns None for top level positions.
    """
    for cut in range(len(key) - 1, 0, -1):
        if key[cut] in "._" and key[:cut] in keys:
            return key[:cut]
    if key != "" and "" in keys:
        return ""
    return None


def hierarchy(index):
    """
    For every position of the table: (parent key, depth, has children)
    """
    keys = [str(key) for key in index]
    key_set = set(keys)
    parents = {key: parent_key(key, key_set) for key in keys}

    def depth(key):
        level = 0
        while parents.get(key) is not None:
            key = parents[key]
            level += 1
        return level

    with_children = {parent for parent in parents.values() if parent is not None}
    return {key: (parents[key], depth(key), key in with_children) for key in keys}


def year_change(current, previous):
    if not isinstance(current, Decimal) or not isinstance(previous, Decimal):
        return None, None
    change = current - previous
    if previous == 0:
        return change, None
    try:
        return change, Percent(change / abs(previous) * 100)
    except InvalidOperation:
        return change, None


def with_year_change(dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    Adds the change between the current (2nd column) and previous year (3rd column)
    """
    current_col, previous_col = dataframe.columns[1], dataframe.columns[2]
    changes = [year_change(row[current_col], row[previous_col]) for _, row in dataframe.iterrows()]
    result = dataframe.copy()
    result[CHANGE_COLUMN] = pd.Series([c for c, _ in changes], index=dataframe.index, dtype=object)
    result[CHANGE_PERCENT_COLUMN] = pd.Series([p for _, p in changes], index=dataframe.index, dtype=object)
    return result
