import re
import shutil
import tempfile
from pathlib import Path
from typing import List, Dict

import pywintypes
import win32com.client.gencache

import scripts.yacg_python.card_data as card_data
import scripts.yacg_python.illustrator_com as illustrator
from scripts.yacg_python.common_vars import CARD_FRONT_PATH


class IllustratorTemplateError(ValueError):
    """
    Raised when an unexpected element or setting of an Illustrator template is found
    """

    def __init__(self, message):
        super().__init__(message)


def get_illustrator_app() -> illustrator._Application:
    """
    If Illustrator is open, returns the opened application.
    Otherwise, opens a new one and returns it.
    """

    try:
        app: illustrator._Application = win32com.client.GetActiveObject("Illustrator.Application")._dispobj_
        # If this doesn't return com_error, then an Illustrator app was already open
    except pywintypes.com_error:
        # The Illustrator app wasn't open, open it
        app: illustrator._Application = win32com.client.Dispatch("Illustrator.Application")
    return app


def create_card(card: card_data.Card, output_path: Path) -> None:
    shutil.copy2(CARD_FRONT_PATH, output_path)

    app = get_illustrator_app()
    document: illustrator.Document = app.Open(output_path)

    if not document.Layers.Count == 5:
        raise IllustratorTemplateError(f"Expected 5 layers, found {document.Layers.Count} instead")

    non_creature_layer = document.Layers.Item(1)
    if not non_creature_layer.Name == "NonCreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 1st layer's name to be 'NonCreatureLayer', found '{non_creature_layer.Name}' instead")
    create_card_non_creature_layer(card, non_creature_layer)

    creature_layer = document.Layers.Item(2)
    if not creature_layer.Name == "CreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 2nd layer's name to be 'CreatureLayer', found '{creature_layer.Name}' instead")
    create_card_creature_layer(card, creature_layer)

    base_layer = document.Layers.Item(3)
    if not base_layer.Name == "BaseLayer":
        raise IllustratorTemplateError(
            f"Expected 3rd layer's name to be 'BaseLayer', found '{creature_layer.Name}' instead")
    create_card_base_layer(card, base_layer)

    background_layer = document.Layers.Item(4)
    if not background_layer.Name == "BackgroundLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundLayer', found '{background_layer.Name}' instead")
    create_card_background_layer(card, background_layer)

    aux_layer = document.Layers.Item(5)
    if not aux_layer.Name == "AuxLayer":
        raise IllustratorTemplateError(
            f"Expected 5th layer's name to be 'AuxLayer', found '{aux_layer.Name}' instead")
    aux_layer.Visible = False


def create_card_non_creature_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, card_data.Creature):
        # This layer isn't used by creature cards
        layer.Visible = False
        return

    # If we're here, card is an effect
    card: card_data.Effect
    layer.Visible = True

    # Validate and select card icon - START
    if not layer.GroupItems.Count == 3:
        raise IllustratorTemplateError(
            f"Expected non-creature layer to have 3 group items, found {layer.GroupItems.Count} instead")
    group_item_names_found = {
        "AuraIcon": False,
        "ActionIcon": False,
        "FieldIcon": False,
    }
    group_item_name_to_effect_type = {
        "AuraIcon": card_data.EffectType.AURA,
        "ActionIcon": card_data.EffectType.ACTION,
        "FieldIcon": card_data.EffectType.FIELD,
    }
    for i in range(1, 4):
        group_item = layer.GroupItems.Item(i)
        group_item_name = group_item.Name
        if group_item_name not in group_item_names_found.keys():
            raise IllustratorTemplateError(
                f"Expected non-creature layer's group icon {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In non-creature layer, found two group icons with the name '{group_item_name}'")
        group_item_names_found[group_item_name] = True
        effect_type = group_item_name_to_effect_type[group_item_name]

        if effect_type == card.data.type:
            group_item.Hidden = False
        else:
            group_item.Hidden = True
    # Validate and select card icon - END

    # Validate and fill in card description - START
    if not layer.TextFrames.Count == 1:
        raise IllustratorTemplateError(
            f"Expected non-creature layer to have 1 text frame, found {layer.TextFrames.Count} instead"
        )
    description_text_frame = layer.TextFrames.Item(1)
    description_text_frame.Contents = card.data.description
    # Validate and fill in card description - END


def create_card_creature_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, card_data.Effect):
        # This layer isn't used by effect cards
        layer.Visible = False
        return

    # If we're here, card is a creature
    card: card_data.Creature
    layer.Visible = True

    # Validate and fill in card's stats - START
    if not layer.GroupItems.Count == 4:
        raise IllustratorTemplateError(
            f"Expected 'CreatureLayer' to have 4 group items, found {layer.GroupItems.Count} instead")
    group_item_names_found = {
        "CreatureIcon": False,
        "Health": False,
        "Attack": False,
        "Speed": False,
    }
    for i in range(1, 5):
        group_item = layer.GroupItems.Item(i)
        group_item_name = group_item.Name
        if group_item_name not in group_item_names_found.keys():
            raise IllustratorTemplateError(
                f"Expected creature layer's group icon {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In creature layer, found two group icons with the name '{group_item_name}'")
        group_item_names_found[group_item_name] = True

        if group_item_name == "CreatureIcon":
            continue

        # If we're here, we're looking at one of the 3 creature's stats
        if not group_item.TextFrames.Count == 1:
            raise IllustratorTemplateError(
                f"Expected non-creature layer's '{group_item_name}' to have 1 text frame, found "
                f"{group_item.TextFrames.Count} instead")
        stat_value_text_frame = group_item.TextFrames.Item(1)
        if group_item_name == "Health":
            stat_value_text_frame.Contents = str(card.data.hp)
        elif group_item_name == "Attack":
            stat_value_text_frame.Contents = str(card.data.atk)
        else:
            stat_value_text_frame.Contents = str(card.data.spe)
    # Validate and fill in card's stats - END

    # Validate and fill in card's description - START
    if not layer.TextFrames.Count == 1:
        raise IllustratorTemplateError(
            f"Expected non-creature layer to have 1 text frame, found {layer.TextFrames.Count} instead"
        )
    description_text_frame = layer.TextFrames.Item(1)

    # Load required fonts
    trait_name_style_name = "Creature Trait Name"
    trait_description_style_name = "Creature Trait Description"
    icon_style_name = "Icons"
    auxiliary_style_name = "Auxiliary"
    try:
        trait_name_style = layer.Parent.CharacterStyles.Item(trait_name_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{trait_name_style_name}'")
    try:
        trait_description_style = layer.Parent.CharacterStyles.Item(trait_description_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{trait_description_style_name}'")
    try:
        icon_style = layer.Parent.CharacterStyles.Item(icon_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{icon_style_name}'")
    try:
        auxiliary_style = layer.Parent.CharacterStyles.Item(auxiliary_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{auxiliary_style_name}'")

    # If we change the contents of text that has multiple styles, Illustrator's styles get confused and bugs happen.
    # To avoid this, we first set all contents to a single style, then change the content's text, then apply the
    # different styles.

    # Some of the text will be icons, which are special Unicode characters.
    # Almost all fonts don't have these characters.
    # If a character is set to a style whose font doesn't have that character, bugs happen.
    # That's why we need to start by setting the whole text to the auxiliary style: its font has all Basic Latin
    # characters, plus the special characters for icons.
    auxiliary_style.ApplyTo(description_text_frame.TextRange, True)

    # Stores trait names. Will be used later to format the trait names
    trait_names: List[str] = []

    # Fill in contents
    description_text_frame.Contents = ""
    for trait_index, trait in enumerate(card.data.traits, start=1):
        if trait_index > 1:
            # This isn't the first trait added
            # So we must add a new line to separate from the previous trait
            description_text_frame.Contents = description_text_frame.Contents + "\r"

        # Add trait name and description
        # (If name is empty, use dev name instead)
        trait_name = trait.data.name
        if trait_name == "":
            trait_name = f"({trait.metadata.dev_name})"
        trait_text = f"{trait_name} {trait.data.description}"
        description_text_frame.Contents = description_text_frame.Contents + trait_text

        trait_names.append(trait_name)

    # Format context by paragraph
    for paragraph_index, trait_name in enumerate(trait_names, start=1):
        paragraph = description_text_frame.Paragraphs.Item(paragraph_index)

        icons_indexes = replace_keywords_with_icons(paragraph)
        for i in range(1, len(trait_name) + 2):  # Includes the space after trait name
            trait_name_style.ApplyTo(paragraph.Characters.Item(i), True)

        for i in range(len(trait_name) + 2, len(paragraph.Contents) + 1):
            if i in icons_indexes:
                icon_style.ApplyTo(paragraph.Characters.Item(i), True)
            else:
                trait_description_style.ApplyTo(paragraph.Characters.Item(i), True)
    # Validate and fill in card's description - END


def replace_keywords_with_icons(text_frame: illustrator.TextFrame) -> List[int]:
    """
    Replaces all keywords in the text frame with the respective icons
    Returns the index of the icons (1-indexed)
    """

    keyword_to_character_dict: Dict[str, str] = {
        "[CREATURE]": "\uE100",
        "[ACTION]": "\uE101",
        "[AURA]": "\uE102",
        "[FIELD]": "\uE103",
        "[HP]": "\uE200",
        "[ATK]": "\uE201",
        "[SPE]": "\uE202",
        "[NOCOLOR]": "\uE300",
        "[BLACK]": "\uE301",
        "[BLUE]": "\uE302",
        "[CYAN]": "\uE303",
        "[GREEN]": "\uE304",
        "[ORANGE]": "\uE305",
        "[PINK]": "\uE306",
        "[PURPLE]": "\uE307",
        "[WHITE]": "\uE308",
        "[YELLOW]": "\uE309",
    }

    substitute_pattern_string = "|".join(
        re.escape(key) for key in keyword_to_character_dict  # Escape the values because they have special characters
    )
    substitute_pattern = re.compile(substitute_pattern_string, re.NOFLAG)
    text_frame.Contents = substitute_pattern.sub(
        lambda match: keyword_to_character_dict[match.group(0)],
        text_frame.Contents
    )

    find_icons_pattern_string = "|".join(keyword_to_character_dict.values())
    find_icons_pattern = re.compile(find_icons_pattern_string, re.NOFLAG)
    icons_indexes = [
        match.start() + 1
        for match in find_icons_pattern.finditer(text_frame.Contents)
    ]

    return icons_indexes


def create_card_base_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
    version_tag = "TEST"

    if not layer.TextFrames.Count == 5:
        raise IllustratorTemplateError(
            f"Expected base layer to have 5 group items, found {layer.TextFrames.Count} instead")
    text_frames_names_found = {
        "Title": False,
        "CostTotalText": False,
        "CostColorText": False,
        "CostNonColorText": False,
        "Identifier": False,
    }
    for i in range(1, 6):
        text_frame = layer.TextFrames.Item(i)
        text_frame_name = text_frame.Name
        if text_frame_name not in text_frames_names_found.keys():
            raise IllustratorTemplateError(
                f"Expected base layer's text frame {i} to have a name in the list "
                f"{list(text_frames_names_found.keys())}, found name '{text_frame_name}' instead"
            )
        if text_frames_names_found[text_frame_name]:
            raise IllustratorTemplateError(
                f"In base layer, found two text frames with the name '{text_frame_name}'")
        text_frames_names_found[text_frame_name] = True

        contents = ""
        if text_frame_name == "Title":
            contents = card.data.name
            if contents == "":
                contents = f"({card.metadata.dev_name})"
        elif text_frame_name == "CostTotalText":
            contents = str(card.data.cost_total)
        elif text_frame_name == "CostColorText":
            contents = str(card.data.cost_color)
        elif text_frame_name == "CostNonColorText":
            contents = str(card.data.cost_total - card.data.cost_color)
        elif text_frame_name == "Identifier":
            contents = f"{version_tag} | {card.metadata.id}"
        text_frame.Contents = contents


def create_card_background_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
    group_item_names_to_color = {
        "BackgroundNone": card_data.Color.NONE,
        "BackgroundBlack": card_data.Color.BLACK,
        "BackgroundBlue": card_data.Color.BLUE,
        "BackgroundCyan": card_data.Color.CYAN,
        "BackgroundGreen": card_data.Color.GREEN,
        "BackgroundOrange": card_data.Color.ORANGE,
        "BackgroundPink": card_data.Color.PINK,
        "BackgroundPurple": card_data.Color.PURPLE,
        "BackgroundWhite": card_data.Color.WHITE,
        "BackgroundYellow": card_data.Color.YELLOW,
    }
    color: card_data.Color = card.data.color

    if not layer.GroupItems.Count == 10:
        raise IllustratorTemplateError(
            f"Expected 'BackgroundLayer' to have 10 group items, found {layer.GroupItems.Count} instead")
    group_item_names_found = {
        "BackgroundNone": False,
        "BackgroundBlack": False,
        "BackgroundBlue": False,
        "BackgroundCyan": False,
        "BackgroundGreen": False,
        "BackgroundOrange": False,
        "BackgroundPink": False,
        "BackgroundPurple": False,
        "BackgroundWhite": False,
        "BackgroundYellow": False,
    }

    for i in range(1, 11):
        group_item = layer.GroupItems.Item(i)
        group_item_name = group_item.Name
        if group_item_name not in group_item_names_found.keys():
            raise IllustratorTemplateError(
                f"Expected background layer's group icon {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In background layer, found two group icons with the name '{group_item_name}'")
        group_item_names_found[group_item_name] = True

        group_item.Hidden = not (group_item_names_to_color[group_item_name] == color)


card_data.import_all_data()
creature = card_data.Creature.get_creature_dict()["C046"]
effect = card_data.Effect.get_effect_dict()["E001"]

with tempfile.TemporaryDirectory() as temp_dir:
    output_path = Path(temp_dir) / "card_front.temp"
    create_card(creature, output_path)

    print("HERE")
