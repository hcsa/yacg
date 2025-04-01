"""
Run this script to move card data from Excel to YAML files
"""

import sys
import traceback
from typing import Iterator

import numpy as np
import pandas as pd
import xlwings as xw

import src.cards as cards
from src.utils import EXCEL_PATH


def main():
    print(f"Exporting card data from '{EXCEL_PATH.name}' to YAML files...")
    import_from_excel()
    cards.export_all_data()
    print("Done!")


def import_from_excel():
    with xw.App(visible=False) as app:
        excel_book = app.books.open(str(EXCEL_PATH))

        import_from_traits_sheet(excel_book.sheets["Traits"], excel_book.sheets["Colors Overview"])
        import_from_creatures_sheet(excel_book.sheets["Creatures"])
        import_from_effects_sheet(excel_book.sheets["Effects"])
        import_from_colors_overview_sheet(excel_book.sheets["Colors Overview"])


def import_from_traits_sheet(traits_sheet: xw.Sheet, colors_overview_sheet: xw.Sheet):
    df = import_traits_sheet_to_df(traits_sheet)
    populate_id_row(df, cards.Trait.id_prefix)

    df_colors_overview = import_colors_overview_sheet_to_df(colors_overview_sheet)
    df_colors = pd.melt(
        df_colors_overview,
        id_vars=["id"],
        value_vars=[
            cards.Color.ORANGE.name,
            cards.Color.GREEN.name,
            cards.Color.BLUE.name,
            cards.Color.WHITE.name,
            cards.Color.YELLOW.name,
            cards.Color.PURPLE.name,
            cards.Color.PINK.name,
            cards.Color.BLACK.name,
            cards.Color.CYAN.name,
        ],
        var_name="color",
        value_name="value"
    )
    df_colors = df_colors[df_colors["value"] != ""]

    for _, row in df.iterrows():
        df_row_colors = df_colors[df_colors["id"] == row["id"].strip()]

        trait_data = cards.TraitData(
            name=row["name"].strip(),
            description=row["description"].strip(),
        )
        trait_colors = cards.TraitColors(
            primary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "XXX"]["color"]
            ],
            secondary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "XX"]["color"]
            ],
            tertiary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "X"]["color"]
            ]
        )
        trait_metadata = cards.TraitMetadata(
            id=row["id"].strip(),
            type=cards.TraitType(row["type"].strip()),
            colors=trait_colors,
            value=(
                int(row["value"])
                if not pd.isna(row["value"])
                else None
            ),
            dev_stage=cards.DevStage(row["dev-stage"].strip()),
            dev_name=row["dev-name"].strip(),
            order=(
                int(row["order"])
                if not pd.isna(row["order"])
                else None
            ),
            summary=row["summary"].strip(),
            notes=row["notes"].strip(),
        )
        _ = cards.Trait(
            data=trait_data,
            metadata=trait_metadata,
        )


def import_traits_sheet_to_df(traits_sheet: xw.Sheet) -> pd.DataFrame:
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


def import_from_creatures_sheet(creatures_sheet: xw.Sheet):
    df = import_creatures_sheet_to_df(creatures_sheet)
    populate_id_row(df, cards.Creature.id_prefix)

    for _, row in df.iterrows():
        traits = [
            cards.Trait.get_trait(row[f"id-trait-{i}"].strip())
            for i in range(1, 5)
            if not row[f"id-trait-{i}"].strip() == ""
        ]
        creature_data = cards.CreatureData(
            name=row["name"],
            color=(
                cards.Color(row["color"].strip())
                if not row["color"].strip() == ""
                else None
            ),
            is_token=row["is-token"],
            cost_total=(
                int(row["cost-total"])
                if not pd.isna(row["cost-total"])
                else None
            ),
            cost_color=(
                int(row["cost-color"])
                if not pd.isna(row["cost-color"])
                else None
            ),
            hp=(
                int(row["hp"])
                if not pd.isna(row["hp"])
                else None
            ),
            atk=(
                int(row["atk"])
                if not pd.isna(row["atk"])
                else None
            ),
            spe=(
                int(row["spe"])
                if not pd.isna(row["spe"])
                else None
            ),
            traits=traits,
            flavor_text=row["flavor-text"].strip()
        )
        creature_metadata = cards.CreatureMetadata(
            id=row["id"].strip(),
            value=(
                int(row["value"])
                if not pd.isna(row["value"])
                else None
            ),
            dev_stage=cards.DevStage(row["dev-stage"].strip()),
            dev_name=row["dev-name"].strip(),
            order=(
                int(row["order"])
                if not pd.isna(row["order"])
                else None
            ),
            summary=row["summary"].strip(),
            notes=row["notes"].strip(),
        )
        _ = cards.Creature(
            data=creature_data,
            metadata=creature_metadata,
        )


def import_creatures_sheet_to_df(creatures_sheet: xw.Sheet) -> pd.DataFrame:
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
        "spe": "Int64",
        "value": "Int64",
        "is-token": bool,
        "flavor-text": str,
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
        df_raw.columns[0:9].append(df_raw.columns[13:24])
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

    return df


def import_from_effects_sheet(effects_sheet: xw.Sheet):
    df = import_effects_sheet_to_df(effects_sheet)
    populate_id_row(df, cards.Effect.id_prefix)

    for _, row in df.iterrows():
        effect_data = cards.EffectData(
            name=row["name"].strip(),
            color=(
                cards.Color(row["color"].strip())
                if not row["color"].strip() == ""
                else None
            ),
            type=cards.EffectType(row["type"].strip()),
            cost_total=(
                int(row["cost-total"])
                if not pd.isna(row["cost-total"])
                else None
            ),
            cost_color=(
                int(row["cost-color"])
                if not pd.isna(row["cost-color"])
                else None
            ),
            description=row["description"].strip(),
            flavor_text=row["flavor-text"].strip()
        )
        effect_metadata = cards.EffectMetadata(
            id=row["id"].strip(),
            dev_stage=cards.DevStage(row["dev-stage"].strip()),
            dev_name=row["dev-name"].strip(),
            order=(
                int(row["order"])
                if not pd.isna(row["order"])
                else None
            ),
            summary=row["summary"].strip(),
            notes=row["notes"].strip(),
        )
        _ = cards.Effect(
            data=effect_data,
            metadata=effect_metadata,
        )


def import_effects_sheet_to_df(effects_sheet: xw.Sheet) -> pd.DataFrame:
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
        "flavor-text": str,
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

    return df


def import_from_colors_overview_sheet(colors_overview_sheet: xw.Sheet):
    df = import_colors_overview_sheet_to_df(colors_overview_sheet)

    # This df has mechanics and traits
    # Assume rows without id are mechanics (traits are added through their sheet)
    populate_id_row(df, cards.Mechanic.id_prefix)

    df_colors = pd.melt(
        df,
        id_vars=["id"],
        value_vars=[
            cards.Color.ORANGE.name,
            cards.Color.GREEN.name,
            cards.Color.BLUE.name,
            cards.Color.WHITE.name,
            cards.Color.YELLOW.name,
            cards.Color.PURPLE.name,
            cards.Color.PINK.name,
            cards.Color.BLACK.name,
            cards.Color.CYAN.name,
        ],
        var_name="color",
        value_name="value"
    )
    df_colors = df_colors[df_colors["value"] != ""]

    for _, row in df.iterrows():
        df_row_colors = df_colors[df_colors["id"] == row["id"].strip()]

        mechanic_colors = cards.MechanicColors(
            primary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "XXX"]["color"]
            ],
            secondary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "XX"]["color"]
            ],
            tertiary=[
                cards.Color(color_name)
                for color_name in df_row_colors[df_row_colors["value"] == "X"]["color"]
            ]
        )

        # This may be a trait. If it is, creation will give error, and we'll just ignore it
        try:
            _ = cards.Mechanic(
                name=row["name"].strip(),
                colors=mechanic_colors,
                id=row["id"].strip(),
                dev_stage=cards.DevStage(row["dev-stage"].strip()),
                order=(
                    int(row["order"])
                    if not pd.isna(row["order"])
                    else None
                ),
                notes=row["notes"].strip(),
            )
        except ValueError:
            continue


def import_colors_overview_sheet_to_df(colors_overview: xw.Sheet) -> pd.DataFrame:
    # Int64 is used here instead of int because it's nullable
    df_types = {
        "id": str,
        "order": "Int64",
        "name": str,
        cards.Color.ORANGE.name: str,
        cards.Color.GREEN.name: str,
        cards.Color.BLUE.name: str,
        cards.Color.WHITE.name: str,
        cards.Color.YELLOW.name: str,
        cards.Color.PURPLE.name: str,
        cards.Color.PINK.name: str,
        cards.Color.BLACK.name: str,
        cards.Color.CYAN.name: str,
        "dev-stage": str,
        "notes": str,
    }

    excel_table_last_cell = colors_overview.range("TableMechanic").last_cell
    df_raw: pd.DataFrame = colors_overview.range("A1", excel_table_last_cell).options(pd.DataFrame, index=False).value
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

    return df


def populate_id_row(df: pd.DataFrame, id_prefix: str):
    id_generator = new_id_generator(df["id"], id_prefix)

    number_missing_ids = len(df.loc[df["id"] == ""])
    new_ids = [x for x, _ in zip(id_generator, range(number_missing_ids))]
    df.loc[df["id"] == "", "id"] = new_ids


def new_id_generator(id_row: pd.Series, prefix: str) -> Iterator[str]:
    """
    Yields IDs that don't clash with the ones in the given ID row
    """

    for i in range(1, 1000):
        new_id = f"{prefix}{i:03d}"
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
