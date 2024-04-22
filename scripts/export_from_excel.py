import os
import sys
import traceback
from pathlib import Path
from typing import Iterator, Dict

import numpy as np
import pandas as pd
import xlwings as xw

EXCEL_NAME = "Cards.xlsx"


def main():
    base_dir = get_base_dir()
    excel_path = base_dir / EXCEL_NAME

    print(f"Exporting card data from '{EXCEL_NAME}'...")
    export_from_excel(excel_path, base_dir)
    print("Done!")


def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        # We're running from .exe
        # .exe is located in the root
        base_dir = Path(os.path.dirname(sys.executable))
    else:
        try:
            base_dir = Path(os.path.realpath(__file__)).parent.parent
            # If this doesn't return NameError, then we're running in non-interactive mode
            # script is located in a directory in the root
        except NameError:
            # We're running in interactive mode
            # script is located in a directory in the root
            base_dir = Path(os.getcwd()).parent

    return base_dir


def export_from_excel(excel_path: Path, base_dir: Path):
    creatures_dir = base_dir / "card_data" / "creatures"
    effects_dir = base_dir / "card_data" / "effects"
    traits_dir = base_dir / "card_data" / "traits"

    with xw.App(visible=False) as app:
        excel_book = app.books.open(str(excel_path))

        export_from_creatures_sheet(excel_book, creatures_dir)
        export_from_effects_sheet(excel_book, effects_dir)
        export_from_traits_sheet(excel_book, traits_dir)


def export_from_creatures_sheet(excel_book: xw.Book, creatures_dir: Path):
    df = export_creatures_sheet_to_df(excel_book.sheets["Creatures"])
    populate_id_row(df)

    traits_text = get_traits_text_for_creature_files(excel_book.sheets["Traits"])
    for _, row in df.iterrows():
        export_creature_to_file(row, creatures_dir, traits_text)


def export_creatures_sheet_to_df(creatures_sheet: xw.Sheet) -> pd.DataFrame:
    # Int64 is used here instead of int because it's nullable
    df_types = {
        "id": str,
        "order": "Int64",
        "name": str,
        "color": str,
        "cost-total": "Int64",
        "cost-color": "Int64",
        "hp": "Int64",
        "atk": "Int64",
        "spd": "Int64",
        "value": "Int64",
        "dev-stage": str,
        "dev-name": str,
        "summary": str,
        "notes": str,
        "id-trait-1": str,
        "id-trait-2": str,
        "id-trait-3": str,
        "id-trait-4": str,
    }

    excel_table_last_cell = creatures_sheet.range("TableCreature").last_cell
    df_raw: pd.DataFrame = creatures_sheet.range("A1", excel_table_last_cell).options(pd.DataFrame, index=False).value
    df_raw_cols_used = \
        df_raw.columns[0:9].append(df_raw.columns[13:22])
    df: pd.DataFrame = df_raw[df_raw_cols_used].copy()
    df = df.set_axis(
        list(df_types),
        axis=1,
    )
    for col, col_type in df_types.items():
        if col_type == str:
            df[col].fillna("", inplace=True)
        if col_type == "Int64":
            df[col].fillna(np.nan, inplace=True)
    df = df.astype(df_types)

    # At this point, the df has the expected types
    # But missing values in int columns are going to be written as "<NA>" in the YAML
    # So we replace them as the empty string
    df = df.astype(object).fillna("")

    return df


def export_creature_to_file(creature_row: pd.Series, creatures_dir: Path, traits_text: Dict[str, str]):
    yaml_path = creatures_dir / f"{creature_row['id']}.yaml"

    traits = ""
    if not creature_row["id-trait-1"] == "":
        id_trait = creature_row["id-trait-1"]
        traits += f"    traits:\n{traits_text[id_trait]}\n"
    if not creature_row["id-trait-2"] == "":
        id_trait = creature_row["id-trait-2"]
        traits += f"{traits_text[id_trait]}\n"
    if not creature_row["id-trait-3"] == "":
        id_trait = creature_row["id-trait-3"]
        traits += f"{traits_text[id_trait]}\n"
    if not creature_row["id-trait-4"] == "":
        id_trait = creature_row["id-trait-4"]
        traits += f"{traits_text[id_trait]}\n"

    notes = creature_row["notes"].strip().replace("\n", "\n      ")

    yaml_content = f"""
creature:
  data:
    name: {creature_row["name"]}
    color: {creature_row["color"]}
    cost-total: {creature_row["cost-total"]}
    cost-color: {creature_row["cost-color"]}
    hp: {creature_row["hp"]}
    atk: {creature_row["atk"]}
    spd: {creature_row["spd"]}
{traits}
  metadata:
    id: {creature_row["id"]}
    value: {creature_row["value"]}
    dev-stage: {creature_row["dev-stage"]}
    dev-name: {creature_row["dev-name"]}
    order: {creature_row["order"]}
    summary: {creature_row["summary"]}
    notes: |
      {notes}"""[1:]

    with open(yaml_path, "w") as f:
        f.write(yaml_content)


def get_traits_text_for_creature_files(traits_sheet: xw.Sheet) -> Dict[str, str]:
    output = {}

    df = export_traits_sheet_to_df(traits_sheet)
    for _, row in df.iterrows():
        output[row["id"]] = f"""
      - name: {row["name"]}
        description: {row["description"]}
        id: {row["id"]}"""[1:]

    return output


def export_from_effects_sheet(excel_book: xw.Book, effects_dir: Path):
    df = export_effects_sheet_to_df(excel_book.sheets["Effects"])
    populate_id_row(df)

    for _, row in df.iterrows():
        export_effect_to_file(row, effects_dir)


def export_effects_sheet_to_df(effects_sheet: xw.Sheet) -> pd.DataFrame:
    # Int64 is used here instead of int because it's nullable
    df_types = {
        "id": str,
        "order": "Int64",
        "name": str,
        "color": str,
        "type": str,
        "cost-total": "Int64",
        "cost-color": "Int64",
        "description": str,
        "dev-stage": str,
        "dev-name": str,
        "summary": str,
        "notes": str,
    }

    excel_table_last_cell = effects_sheet.range("TableEffect").last_cell
    df_raw: pd.DataFrame = effects_sheet.range("A1", excel_table_last_cell).options(pd.DataFrame, index=False).value
    df_raw_cols_used = df_raw.columns[0:-1]
    df: pd.DataFrame = df_raw[df_raw_cols_used].copy()
    df = df.set_axis(
        list(df_types),
        axis=1,
    )
    for col, col_type in df_types.items():
        if col_type == str:
            df[col].fillna("", inplace=True)
        if col_type == "Int64":
            df[col].fillna(np.nan, inplace=True)
    df = df.astype(df_types)

    # At this point, the df has the expected types
    # But missing values in int columns are going to be written as "<NA>" in the YAML
    # So we replace them as the empty string
    df = df.astype(object).fillna("")

    return df


def export_effect_to_file(effect_row: pd.Series, effects_dir: Path):
    yaml_path = effects_dir / f"{effect_row['id']}.yaml"

    notes = effect_row["notes"].strip().replace("\n", "\n      ")

    yaml_content = f"""
effect:
  data:
    name: {effect_row["name"]}
    color: {effect_row["color"]}
    type: {effect_row["type"]}
    cost-total: {effect_row["cost-total"]}
    cost-color: {effect_row["cost-color"]}
    description: |
      {effect_row["description"]}

  metadata:
    id: {effect_row["id"]}
    dev-stage: {effect_row["dev-stage"]}
    dev-name: {effect_row["dev-name"]}
    order: {effect_row["order"]}
    summary: {effect_row["summary"]}
    notes: |
      {notes}"""[1:]

    with open(yaml_path, "w") as f:
        f.write(yaml_content)


def export_from_traits_sheet(excel_book: xw.Book, traits_dir: Path):
    df = export_traits_sheet_to_df(excel_book.sheets["Traits"])
    populate_id_row(df)

    for _, row in df.iterrows():
        export_trait_to_file(row, traits_dir)


def export_traits_sheet_to_df(traits_sheet: xw.Sheet) -> pd.DataFrame:
    # Int64 is used here instead of int because it's nullable
    df_types = {
        "id": str,
        "order": "Int64",
        "name": str,
        "description": str,
        "type": str,
        "value": "Int64",
        "dev-stage": str,
        "dev-name": str,
        "summary": str,
        "notes": str,
    }

    excel_table_last_cell = traits_sheet.range("TableTrait").last_cell
    df_raw: pd.DataFrame = traits_sheet.range("A1", excel_table_last_cell).options(pd.DataFrame, index=False).value
    df_raw_cols_used = df_raw.columns[0:-1]
    df: pd.DataFrame = df_raw[df_raw_cols_used].copy()
    df = df.set_axis(
        list(df_types),
        axis=1,
    )
    for col, col_type in df_types.items():
        if col_type == str:
            df[col].fillna("", inplace=True)
        if col_type == "Int64":
            df[col].fillna(np.nan, inplace=True)
    df = df.astype(df_types)

    # At this point, the df has the expected types
    # But missing values in int columns are going to be written as "<NA>" in the YAML
    # So we replace them as the empty string
    df = df.astype(object).fillna("")

    return df


def export_trait_to_file(trait_row: pd.Series, traits_dir: Path):
    yaml_path = traits_dir / f"{trait_row['id']}.yaml"

    notes = trait_row["notes"].strip().replace("\n", "\n      ")

    yaml_content = f"""
trait:
  data:
    name: {trait_row["name"]}
    description: |
      {trait_row["description"]}

  metadata:
    id: {trait_row["id"]}
    type: {trait_row["type"]}
    value: {trait_row["value"]}
    dev-stage: {trait_row["dev-stage"]}
    dev-name: {trait_row["dev-name"]}
    order: {trait_row["order"]}
    summary: {trait_row["summary"]}
    notes: |
      {notes}"""[1:]

    with open(yaml_path, "w") as f:
        f.write(yaml_content)


def populate_id_row(df: pd.DataFrame):
    id_generator = new_id_generator(df["id"])

    number_missing_ids = len(df.loc[df["id"] == ""])
    new_ids = [x for x, _ in zip(id_generator, range(number_missing_ids))]
    df.loc[df["id"] == "", "id"] = new_ids


def new_id_generator(id_row: pd.Series) -> Iterator[str]:
    """
    Yields IDs that don't clash with the ones in the given ID row
    """

    first_id = id_row[id_row != ""].iloc[0]
    id_initial_letter: str = first_id[0]

    for i in range(1, 1000):
        new_id = f"{id_initial_letter}{i:03d}"
        new_id_exists = id_row.isin([new_id]).any()
        if not new_id_exists:
            yield new_id

    raise ValueError("All IDs are used")


if __name__ == "__main__":
    # noinspection PyBroadException
    try:
        main()
    except BaseException:
        e_data = sys.exc_info()
        traceback.print_exception(*e_data)
        input("\nPress enter to exit...")
