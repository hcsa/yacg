"""
Run this script to move card data from YAML files to Excel
"""

import os
import shutil
import sys
import traceback

import pandas as pd
import xlwings as xw
import yaml

import src.cards as cards
from src.utils import EXCEL_PATH, EXCEL_BACKUP_PATH, EXCEL_TEMPLATE_PATH, VALUES_DATA_PATH


def main():
    print(f"Importing card data from YAML files to '{EXCEL_PATH.name}'...")
    cards.import_all_data()
    copy_excel_template()
    export_to_excel()
    print("Done!")


def copy_excel_template():
    if EXCEL_PATH.is_file():
        shutil.copy2(EXCEL_PATH, EXCEL_BACKUP_PATH)
        os.remove(EXCEL_PATH)
        print(f"Created backup file '{EXCEL_BACKUP_PATH.name}' of old '{EXCEL_PATH.name}'")
    shutil.copy2(EXCEL_TEMPLATE_PATH, EXCEL_PATH)


def export_to_excel():
    with xw.App(visible=False) as app:
        excel_book = app.books.open(str(EXCEL_PATH))

        export_to_traits_sheet(excel_book.sheets["Traits"])
        export_to_creatures_sheet(excel_book.sheets["Creatures"])
        export_to_effects_sheet(excel_book.sheets["Effects"])
        export_to_attacks_sheet(excel_book.sheets["Attacks"])
        export_to_colors_overview_sheet(excel_book.sheets["Colors Overview"])
        export_to_creature_values_sheet(excel_book.sheets["Creatures - Value"])

        excel_book.save()


def export_to_traits_sheet(traits_sheet: xw.Sheet):
    df = get_traits_df()

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


def get_traits_df() -> pd.DataFrame:
    df_data = []

    for trait in cards.Trait.get_trait_dict().values():
        df_row = {
            "id": trait.metadata.id,
            "order": trait.metadata.order,
            "name": trait.data.name,
            "description": trait.data.description,
            "type": trait.metadata.type.name,
            "value": trait.metadata.value,
            "dev-stage": trait.metadata.dev_stage.name,
            "dev-name": trait.metadata.dev_name,
            "summary": trait.metadata.summary,
            "notes": trait.metadata.notes,
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


def export_to_creatures_sheet(creatures_sheet: xw.Sheet):
    df = get_creatures_df()

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
    creatures_sheet["H3"].options(index=False, header=False).value = df["atk-strong"]
    creatures_sheet["I3"].options(index=False, header=False).value = df["atk-technical"]
    creatures_sheet["J3"].options(index=False, header=False).value = df["spe"]
    creatures_sheet["P3"].options(index=False, header=False).value = df["atk-strong-effect-variable"]
    creatures_sheet["R3"].options(index=False, header=False).value = df["atk-technical-effect-variable"]
    creatures_sheet["S3"].options(index=False, header=False).value = df["is-token"]
    creatures_sheet["U3"].options(index=False, header=False).value = df["flavor-text"]
    creatures_sheet["V3"].options(index=False, header=False).value = df["dev-stage"]
    creatures_sheet["W3"].options(index=False, header=False).value = df["dev-name"]
    creatures_sheet["X3"].options(index=False, header=False).value = df["summary"]
    creatures_sheet["Y3"].options(index=False, header=False).value = df["notes"]
    creatures_sheet["Z3"].options(index=False, header=False).value = df["id-trait-1"]
    creatures_sheet["AA3"].options(index=False, header=False).value = df["id-trait-2"]
    creatures_sheet["AB3"].options(index=False, header=False).value = df["id-trait-3"]
    creatures_sheet["AC3"].options(index=False, header=False).value = df["id-trait-4"]
    creatures_sheet["AD3"].options(index=False, header=False).value = df["atk-strong-effect-id"]
    creatures_sheet["AE3"].options(index=False, header=False).value = df["atk-technical-effect-id"]

    # Delete template row
    creatures_sheet.range("2:2").delete(shift="up")


def get_creatures_df() -> pd.DataFrame:
    df_data = []

    creature_list = [c for c in cards.get_all_cards() if isinstance(c, cards.Creature)]
    for creature in creature_list:
        df_row = {
            "id": creature.metadata.id,
            "order": creature.metadata.order,
            "name": creature.data.name,
            "color": (creature.data.color.name if creature.data.color is not None else None),
            "is-token": creature.data.is_token,
            "flavor-text": creature.data.flavor_text,
            "cost-total": creature.data.cost_total,
            "cost-color": creature.data.cost_color,
            "hp": creature.data.hp,
            "atk-strong": creature.data.atk_strong,
            "atk-strong-effect-id": (
                creature.data.atk_strong_effect.get_id()
                if creature.data.atk_strong_effect is not None
                else None
            ),
            "atk-strong-effect-variable": (
                creature.data.atk_strong_effect_variable
                if creature.data.atk_strong_effect_variable is not None
                else None
            ),
            "atk-technical": creature.data.atk_technical,
            "atk-technical-effect-id": (
                creature.data.atk_technical_effect.get_id()
                if creature.data.atk_technical_effect is not None
                else None
            ),
            "atk-technical-effect-variable": (
                creature.data.atk_technical_effect_variable
                if creature.data.atk_technical_effect_variable is not None
                else None
            ),
            "spe": creature.data.spe,
            "dev-stage": creature.metadata.dev_stage.name,
            "dev-name": creature.metadata.dev_name,
            "summary": creature.metadata.summary,
            "notes": creature.metadata.notes,
        }

        for i, trait in enumerate(creature.data.traits, start=1):
            df_row[f"id-trait-{i}"] = trait.get_id()

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        "color",
        "flavor-text",
        "is-token",
        "cost-total",
        "cost-color",
        "hp",
        "atk-strong",
        "atk-strong-effect-id",
        "atk-strong-effect-variable",
        "atk-technical",
        "atk-technical-effect-id",
        "atk-technical-effect-variable",
        "spe",
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


def export_to_effects_sheet(effects_sheet: xw.Sheet):
    df = get_effects_df()

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
    effects_sheet["I3"].options(index=False, header=False).value = df["flavor-text"]
    effects_sheet["J3"].options(index=False, header=False).value = df["dev-stage"]
    effects_sheet["K3"].options(index=False, header=False).value = df["dev-name"]
    effects_sheet["L3"].options(index=False, header=False).value = df["summary"]
    effects_sheet["M3"].options(index=False, header=False).value = df["notes"]

    # Delete template row
    effects_sheet.range("2:2").delete(shift="up")


def get_effects_df() -> pd.DataFrame:
    df_data = []

    effect_list = [c for c in cards.get_all_cards() if isinstance(c, cards.Effect)]
    for effect in effect_list:
        df_row = {
            "id": effect.metadata.id,
            "order": effect.metadata.order,
            "name": effect.data.name,
            "color": (effect.data.color.name if effect.data.color is not None else None),
            "type": effect.data.type.name,
            "cost-total": effect.data.cost_total,
            "cost-color": effect.data.cost_color,
            "description": effect.data.description,
            "flavor-text": effect.data.flavor_text,
            "dev-stage": effect.metadata.dev_stage.name,
            "dev-name": effect.metadata.dev_name,
            "summary": effect.metadata.summary,
            "notes": effect.metadata.notes,
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
        "flavor-text",
        "dev-stage",
        "dev-name",
        "summary",
        "notes"
    ])
    return df


def export_to_attacks_sheet(attacks_sheet: xw.Sheet):
    df = get_attacks_df()

    # Copy formatting from template row
    for _ in range(len(df)):
        attacks_sheet.range("3:3").insert(
            shift="down",
            copy_origin="format_from_left_or_above"
        )

    attacks_sheet["A3"].options(index=False, header=False).value = df["id"]
    attacks_sheet["B3"].options(index=False, header=False).value = df["order"]
    attacks_sheet["C3"].options(index=False, header=False).value = df["name"]
    attacks_sheet["D3"].options(index=False, header=False).value = df["description"]
    attacks_sheet["E3"].options(index=False, header=False).value = df["value"]
    attacks_sheet["F3"].options(index=False, header=False).value = df["dev-stage"]
    attacks_sheet["G3"].options(index=False, header=False).value = df["dev-name"]
    attacks_sheet["H3"].options(index=False, header=False).value = df["summary"]
    attacks_sheet["I3"].options(index=False, header=False).value = df["notes"]

    # Delete template row
    attacks_sheet.range("2:2").delete(shift="up")

    attacks_sheet["A1"].expand("table").api.Sort(
        Key1=attacks_sheet.range("J:J").api,
        Order1=2,
        Key2=attacks_sheet.range("B:B").api,
        Header=1,
        Orientation=1,
    )


def get_attacks_df() -> pd.DataFrame:
    df_data = []

    for attack in cards.Attack.get_attack_dict().values():
        df_row = {
            "id": attack.metadata.id,
            "order": attack.metadata.order,
            "name": attack.data.name,
            "description": attack.data.description,
            "value": attack.metadata.value,
            "dev-stage": attack.metadata.dev_stage.name,
            "dev-name": attack.metadata.dev_name,
            "summary": attack.metadata.summary,
            "notes": attack.metadata.notes,
        }

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        "description",
        "value",
        "dev-stage",
        "dev-name",
        "summary",
        "notes"
    ])
    return df


def export_to_colors_overview_sheet(colors_overview_sheet: xw.Sheet):
    df = get_colors_overview_df()

    # Copy formatting from template row
    for _ in range(len(df)):
        colors_overview_sheet.range("3:3").insert(
            shift="down",
            copy_origin="format_from_left_or_above"
        )

    colors_overview_sheet["A3"].options(index=False, header=False).value = df["id"]
    colors_overview_sheet["B3"].options(index=False, header=False).value = df["order"]
    colors_overview_sheet["C3"].options(index=False, header=False).value = df["name"]
    colors_overview_sheet["D3"].options(index=False, header=False).value = df[cards.Color.ORANGE.name]
    colors_overview_sheet["E3"].options(index=False, header=False).value = df[cards.Color.GREEN.name]
    colors_overview_sheet["F3"].options(index=False, header=False).value = df[cards.Color.BLUE.name]
    colors_overview_sheet["G3"].options(index=False, header=False).value = df[cards.Color.WHITE.name]
    colors_overview_sheet["H3"].options(index=False, header=False).value = df[cards.Color.YELLOW.name]
    colors_overview_sheet["I3"].options(index=False, header=False).value = df[cards.Color.PURPLE.name]
    colors_overview_sheet["J3"].options(index=False, header=False).value = df[cards.Color.PINK.name]
    colors_overview_sheet["K3"].options(index=False, header=False).value = df[cards.Color.BLACK.name]
    colors_overview_sheet["L3"].options(index=False, header=False).value = df[cards.Color.CYAN.name]
    colors_overview_sheet["M3"].options(index=False, header=False).value = df["dev-stage"]
    colors_overview_sheet["N3"].options(index=False, header=False).value = df["notes"]

    # Delete template row
    colors_overview_sheet.range("2:2").delete(shift="up")

    colors_overview_sheet["A1"].expand("table").api.Sort(
        Key1=colors_overview_sheet.range("O:O").api,
        Order1=2,
        Key2=colors_overview_sheet.range("B:B").api,
        Header=1,
        Orientation=1,
    )


def get_colors_overview_df() -> pd.DataFrame:
    df_data = []

    mechanic_list = cards.Mechanic.get_mechanic_dict().values()
    for mechanic in mechanic_list:
        df_row = {
            "id": mechanic.id,
            "order": mechanic.order,
            "name": mechanic.name,
            cards.Color.ORANGE.name: None,
            cards.Color.GREEN.name: None,
            cards.Color.BLUE.name: None,
            cards.Color.WHITE.name: None,
            cards.Color.YELLOW.name: None,
            cards.Color.PURPLE.name: None,
            cards.Color.PINK.name: None,
            cards.Color.BLACK.name: None,
            cards.Color.CYAN.name: None,
            "dev-stage": mechanic.dev_stage.name,
            "notes": mechanic.notes,
        }

        for color in mechanic.colors.primary:
            df_row[color.name] = "XXX"
        for color in mechanic.colors.secondary:
            df_row[color.name] = "XX"
        for color in mechanic.colors.tertiary:
            df_row[color.name] = "X"

        df_data.append(df_row)

    trait_list = cards.Trait.get_trait_dict().values()
    for trait in trait_list:
        df_row = {
            "id": trait.metadata.id,
            "order": 10000 + trait.metadata.order,
            "name": f"{trait.data.name}\n{trait.data.description}",
            cards.Color.ORANGE.name: None,
            cards.Color.GREEN.name: None,
            cards.Color.BLUE.name: None,
            cards.Color.WHITE.name: None,
            cards.Color.YELLOW.name: None,
            cards.Color.PURPLE.name: None,
            cards.Color.PINK.name: None,
            cards.Color.BLACK.name: None,
            cards.Color.CYAN.name: None,
            "dev-stage": trait.metadata.dev_stage.name,
            "notes": "[Use 'Traits' sheet]",
        }

        for color in trait.metadata.colors.primary:
            df_row[color.name] = "XXX"
        for color in trait.metadata.colors.secondary:
            df_row[color.name] = "XX"
        for color in trait.metadata.colors.tertiary:
            df_row[color.name] = "X"

        df_data.append(df_row)

    attack_list = cards.Attack.get_attack_dict().values()
    for attack in attack_list:
        df_row = {
            "id": attack.metadata.id,
            "order": 20000 + attack.metadata.order,
            "name": f"{attack.data.name}\n{attack.data.description}",
            cards.Color.ORANGE.name: None,
            cards.Color.GREEN.name: None,
            cards.Color.BLUE.name: None,
            cards.Color.WHITE.name: None,
            cards.Color.YELLOW.name: None,
            cards.Color.PURPLE.name: None,
            cards.Color.PINK.name: None,
            cards.Color.BLACK.name: None,
            cards.Color.CYAN.name: None,
            "dev-stage": attack.metadata.dev_stage.name,
            "notes": "[Use 'Attacks' sheet]",
        }

        for color in attack.metadata.colors.primary:
            df_row[color.name] = "XXX"
        for color in attack.metadata.colors.secondary:
            df_row[color.name] = "XX"
        for color in attack.metadata.colors.tertiary:
            df_row[color.name] = "X"

        df_data.append(df_row)

    # noinspection PyTypeChecker
    df = pd.DataFrame(data=df_data, dtype=object, columns=[
        "id",
        "order",
        "name",
        cards.Color.ORANGE.name,
        cards.Color.GREEN.name,
        cards.Color.BLUE.name,
        cards.Color.WHITE.name,
        cards.Color.YELLOW.name,
        cards.Color.PURPLE.name,
        cards.Color.PINK.name,
        cards.Color.BLACK.name,
        cards.Color.CYAN.name,
        "dev-stage",
        "dev-name",
        "summary",
        "notes"
    ])
    return df


def export_to_creature_values_sheet(creature_values_sheet: xw.Sheet):
    with open(VALUES_DATA_PATH, "r") as f:
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
        input("\nPress enter to exit...")
