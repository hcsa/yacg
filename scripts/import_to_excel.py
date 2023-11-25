import os
import shutil
import sys
import traceback
from pathlib import Path

import pandas as pd
import xlwings as xw
import yaml

EXCEL_NAME = "Cards.xlsx"
EXCEL_TEMPLATE_NAME = "excel_base.xlsx"


def main():
    print(Path(__file__))

    base_dir = get_base_dir()
    excel_path = base_dir / EXCEL_NAME
    excel_template_path = base_dir / EXCEL_TEMPLATE_NAME
    copy_excel_base(excel_template_path, excel_path)

    print(f"Importing card data to '{EXCEL_NAME}'...")
    import_to_excel(excel_path, base_dir)
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


def copy_excel_base(source_path: Path, destination_path: Path):
    if destination_path.is_file():
        backup_path = destination_path.parent / "Cards (backup).xlsx"
        shutil.copy2(destination_path, backup_path)
        os.remove(destination_path)
        print(f"This will override '{destination_path.name}', created backup file")
    shutil.copy2(source_path, destination_path)


def import_to_excel(excel_path: Path, base_dir: Path):
    creatures_dir = base_dir / "cards" / "creatures"
    effects_dir = base_dir / "cards" / "effects"
    traits_dir = base_dir / "cards" / "traits"
    values_path = base_dir / "game_data" / "values.yaml"

    with xw.App(visible=False) as app:
        excel_book = app.books.open(str(excel_path))

        import_to_creatures_sheet(excel_book.sheets["Creatures"], creatures_dir)
        import_to_effects_sheet(excel_book.sheets["Effects"], effects_dir)
        import_to_traits_sheet(excel_book.sheets["Traits"], traits_dir)
        import_to_creature_values_sheet(excel_book.sheets["Creatures - Value"], values_path)

        excel_book.save()


def import_to_creatures_sheet(creatures_sheet: xw.Sheet, creatures_dir: Path):
    df = import_creatures_to_df(creatures_dir)

    # Copy formatting from template row
    for _ in range(len(df)):
        creatures_sheet.range("3:3").insert(
            shift="down",
            copy_origin="format_from_left_or_above"
        )

    creatures_sheet["A3"].options(index=False, header=False).value = df["id"]
    creatures_sheet["B3"].options(index=False, header=False).value = df["order"]
    creatures_sheet["C3"].options(index=False, header=False).value = df["name"]
    creatures_sheet["D3"].options(index=False, header=False).value = df["color"]
    creatures_sheet["E3"].options(index=False, header=False).value = df["cost-total"]
    creatures_sheet["F3"].options(index=False, header=False).value = df["cost-color"]
    creatures_sheet["G3"].options(index=False, header=False).value = df["hp"]
    creatures_sheet["H3"].options(index=False, header=False).value = df["atk"]
    creatures_sheet["I3"].options(index=False, header=False).value = df["spd"]
    creatures_sheet["O3"].options(index=False, header=False).value = df["dev-stage"]
    creatures_sheet["P3"].options(index=False, header=False).value = df["dev-name"]
    creatures_sheet["Q3"].options(index=False, header=False).value = df["summary"]
    creatures_sheet["R3"].options(index=False, header=False).value = df["notes"]
    creatures_sheet["S3"].options(index=False, header=False).value = df["id-trait-1"]
    creatures_sheet["T3"].options(index=False, header=False).value = df["id-trait-2"]
    creatures_sheet["U3"].options(index=False, header=False).value = df["id-trait-3"]
    creatures_sheet["V3"].options(index=False, header=False).value = df["id-trait-4"]

    # Delete template row
    creatures_sheet.range("2:2").delete(shift="up")

    creatures_sheet["A1"].expand("table").api.Sort(
        Key1=creatures_sheet.range("AF:AF").api,
        Order1=2,
        Key2=creatures_sheet.range("B:B").api,
        Header=1,
        Orientation=1,
    )


def import_creatures_to_df(creatures_dir: Path) -> pd.DataFrame:
    df_data = []

    for yaml_path in creatures_dir.iterdir():
        with open(yaml_path, "r") as f:
            creature_data = yaml.safe_load(f)["creature"]
        df_row = {
            "id": creature_data["metadata"]["id"],
            "order": creature_data["metadata"]["order"],
            "name": creature_data["data"]["name"],
            "color": creature_data["data"]["color"],
            "cost-total": creature_data["data"]["cost-total"],
            "cost-color": creature_data["data"]["cost-color"],
            "hp": creature_data["data"]["hp"],
            "atk": creature_data["data"]["atk"],
            "spd": creature_data["data"]["spd"],
            "dev-stage": creature_data["metadata"]["dev-stage"],
            "dev-name": creature_data["metadata"]["dev-name"],
            "summary": creature_data["metadata"]["summary"],
            "notes": creature_data["metadata"]["notes"].replace("\n      ", "\n").strip(),
        }

        if "traits" in creature_data["data"]:
            for i, trait in enumerate(creature_data["data"]["traits"], start=1):
                df_row[f"id-trait-{i}"] = trait["id"]

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        "color",
        "cost-total",
        "cost-color",
        "hp",
        "atk",
        "spd",
        "dev-stage",
        "dev-name",
        "summary",
        "notes",
        "id-trait-1",
        "id-trait-2",
        "id-trait-3",
        "id-trait-4"
    ])
    return df


def import_to_effects_sheet(effects_sheet: xw.Sheet, effects_dir: Path):
    df = import_effects_to_df(effects_dir)

    # Copy formatting from template row
    for _ in range(len(df)):
        effects_sheet.range("3:3").insert(
            shift="down",
            copy_origin="format_from_left_or_above"
        )

    effects_sheet["A3"].options(index=False, header=False).value = df["id"]
    effects_sheet["B3"].options(index=False, header=False).value = df["order"]
    effects_sheet["C3"].options(index=False, header=False).value = df["name"]
    effects_sheet["D3"].options(index=False, header=False).value = df["color"]
    effects_sheet["E3"].options(index=False, header=False).value = df["type"]
    effects_sheet["F3"].options(index=False, header=False).value = df["cost-total"]
    effects_sheet["G3"].options(index=False, header=False).value = df["cost-color"]
    effects_sheet["H3"].options(index=False, header=False).value = df["description"]
    effects_sheet["I3"].options(index=False, header=False).value = df["dev-stage"]
    effects_sheet["J3"].options(index=False, header=False).value = df["dev-name"]
    effects_sheet["K3"].options(index=False, header=False).value = df["summary"]
    effects_sheet["L3"].options(index=False, header=False).value = df["notes"]

    # Delete template row
    effects_sheet.range("2:2").delete(shift="up")

    effects_sheet["A1"].expand("table").api.Sort(
        Key1=effects_sheet.range("M:M").api,
        Order1=2,
        Key2=effects_sheet.range("B:B").api,
        Header=1,
        Orientation=1,
    )


def import_effects_to_df(effects_dir: Path) -> pd.DataFrame:
    df_data = []

    for yaml_path in effects_dir.iterdir():
        with open(yaml_path, "r") as f:
            effect_data = yaml.safe_load(f)["effect"]
        df_row = {
            "id": effect_data["metadata"]["id"],
            "order": effect_data["metadata"]["order"],
            "name": effect_data["data"]["name"],
            "color": effect_data["data"]["color"],
            "type": effect_data["data"]["type"],
            "cost-total": effect_data["data"]["cost-total"],
            "cost-color": effect_data["data"]["cost-color"],
            "description": effect_data["data"]["description"].strip(),
            "dev-stage": effect_data["metadata"]["dev-stage"],
            "dev-name": effect_data["metadata"]["dev-name"],
            "summary": effect_data["metadata"]["summary"],
            "notes": effect_data["metadata"]["notes"].replace("\n      ", "\n").strip(),
        }

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        "color",
        "type",
        "cost-total",
        "cost-color",
        "description",
        "dev-stage",
        "dev-name",
        "summary",
        "notes"
    ])
    return df


def import_to_traits_sheet(traits_sheet: xw.Sheet, traits_dir: Path):
    df = import_traits_to_df(traits_dir)

    # Copy formatting from template row
    for _ in range(len(df)):
        traits_sheet.range("3:3").insert(
            shift="down",
            copy_origin="format_from_left_or_above"
        )

    traits_sheet["A3"].options(index=False, header=False).value = df["id"]
    traits_sheet["B3"].options(index=False, header=False).value = df["order"]
    traits_sheet["C3"].options(index=False, header=False).value = df["name"]
    traits_sheet["D3"].options(index=False, header=False).value = df["description"]
    traits_sheet["E3"].options(index=False, header=False).value = df["type"]
    traits_sheet["F3"].options(index=False, header=False).value = df["value"]
    traits_sheet["G3"].options(index=False, header=False).value = df["dev-stage"]
    traits_sheet["H3"].options(index=False, header=False).value = df["dev-name"]
    traits_sheet["I3"].options(index=False, header=False).value = df["summary"]
    traits_sheet["J3"].options(index=False, header=False).value = df["notes"]

    # Delete template row
    traits_sheet.range("2:2").delete(shift="up")

    traits_sheet["A1"].expand("table").api.Sort(
        Key1=traits_sheet.range("K:K").api,
        Order1=2,
        Key2=traits_sheet.range("B:B").api,
        Header=1,
        Orientation=1,
    )


def import_traits_to_df(traits_dir: Path) -> pd.DataFrame:
    df_data = []

    for yaml_path in traits_dir.iterdir():
        with open(yaml_path, "r") as f:
            trait_data = yaml.safe_load(f)["trait"]
        df_row = {
            "id": trait_data["metadata"]["id"],
            "order": trait_data["metadata"]["order"],
            "name": trait_data["data"]["name"],
            "description": trait_data["data"]["description"].strip(),
            "type": trait_data["metadata"]["type"],
            "value": trait_data["metadata"]["value"],
            "dev-stage": trait_data["metadata"]["dev-stage"],
            "dev-name": trait_data["metadata"]["dev-name"],
            "summary": trait_data["metadata"]["summary"],
            "notes": trait_data["metadata"]["notes"].replace("\n      ", "\n").strip(),
        }

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        "description",
        "type",
        "value",
        "dev-stage",
        "dev-name",
        "summary",
        "notes"
    ])
    return df


def import_to_creature_values_sheet(creature_values_sheet: xw.Sheet, values_path: Path):
    with open(values_path, "r") as f:
        values_data = yaml.safe_load(f)["values"]

    for row_index, (cost_total, value) in enumerate(values_data["cost-total"].items(), start=2):
        creature_values_sheet[f"A{row_index}"].value = cost_total
        creature_values_sheet[f"B{row_index}"].value = value

    for row_index, (color, value_data) in enumerate(values_data["color"].items(), start=2):
        creature_values_sheet[f"D{row_index}"].value = color
        creature_values_sheet[f"E{row_index}"].value = value_data["cost-total-low"]
        creature_values_sheet[f"F{row_index}"].value = value_data["cost-total-mid"]
        creature_values_sheet[f"G{row_index}"].value = value_data["cost-total-high"]

    for row_index, (hp, value) in enumerate(values_data["hp"].items(), start=2):
        creature_values_sheet[f"I{row_index}"].value = hp
        creature_values_sheet[f"J{row_index}"].value = value

    for row_index, (atk, value) in enumerate(values_data["atk"].items(), start=2):
        creature_values_sheet[f"L{row_index}"].value = atk
        creature_values_sheet[f"M{row_index}"].value = value

    for row_index, (spd, value) in enumerate(values_data["spd"].items(), start=2):
        creature_values_sheet[f"O{row_index}"].value = spd
        creature_values_sheet[f"P{row_index}"].value = value


if __name__ == "__main__":
    # noinspection PyBroadException
    try:
        main()
    except BaseException:
        e_data = sys.exc_info()
        traceback.print_exception(*e_data)
        input("\nPress any key to exit...")
