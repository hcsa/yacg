import re
import shutil
import tempfile
from pathlib import Path
from typing import List, Dict

import pywintypes
import win32com.client.gencache

import scripts.yacg_python.card_data as card_data
import scripts.yacg_python.illustrator_com as illustrator
from scripts.yacg_python.common_vars import CARD_FRONT_PATH, GIT_TAG_NAME


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
    app = get_illustrator_app()

    card_front_path = output_path / f"{card.metadata.id}_front.temp"
    shutil.copy2(CARD_FRONT_PATH, card_front_path)
    card_front_document: illustrator.Document = app.Open(card_front_path)
    create_card_front(card, card_front_document)

    card_back_path = output_path / f"{card.metadata.id}_back.temp"
    shutil.copy2(CARD_FRONT_PATH, card_back_path)
    card_back_document: illustrator.Document = app.Open(card_back_path)
    create_card_back(card, card_back_document)


def create_card_front(card: card_data.Card, document: illustrator.Document) -> None:
    if not document.Layers.Count == 5:
        raise IllustratorTemplateError(f"Expected 5 layers, found {document.Layers.Count} instead")

    non_creature_layer = document.Layers.Item(1)
    if not non_creature_layer.Name == "NonCreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 1st layer's name to be 'NonCreatureLayer', found '{non_creature_layer.Name}' instead")
    create_card_front_non_creature_layer(card, non_creature_layer)

    creature_layer = document.Layers.Item(2)
    if not creature_layer.Name == "CreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 2nd layer's name to be 'CreatureLayer', found '{creature_layer.Name}' instead")
    create_card_front_creature_layer(card, creature_layer)

    base_layer = document.Layers.Item(3)
    if not base_layer.Name == "BaseLayer":
        raise IllustratorTemplateError(
            f"Expected 3rd layer's name to be 'BaseLayer', found '{creature_layer.Name}' instead")
    create_card_front_base_layer(card, base_layer)

    background_color_layer = document.Layers.Item(4)
    if not background_color_layer.Name == "BackgroundColorLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundColorLayer', found '{background_color_layer.Name}' instead")
    create_card_front_background_color_layer(card, background_color_layer)

    aux_layer = document.Layers.Item(5)
    if not aux_layer.Name == "AuxLayer":
        raise IllustratorTemplateError(
            f"Expected 5th layer's name to be 'AuxLayer', found '{aux_layer.Name}' instead")
    aux_layer.Visible = False


def create_card_front_non_creature_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
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
                f"Expected non-creature layer's group item {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In non-creature layer, found two group items with the name '{group_item_name}'")
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

    # Load required fonts
    description_style_name = "Description"
    icon_style_name = "Icons"
    auxiliary_style_name = "Auxiliary"
    try:
        trait_description_style = layer.Parent.CharacterStyles.Item(description_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{description_style_name}'")
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

    description_text_frame.Contents = card.data.description
    icons_indexes = replace_keywords_with_icons(description_text_frame)
    for i in range(1, len(description_text_frame.Contents) + 1):
        if i in icons_indexes:
            icon_style.ApplyTo(description_text_frame.Characters.Item(i), True)
        else:
            trait_description_style.ApplyTo(description_text_frame.Characters.Item(i), True)
    # Validate and fill in card description - END


def create_card_front_creature_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
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
                f"Expected creature layer's group item {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In creature layer, found two group items with the name '{group_item_name}'")
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
    description_style_name = "Description"
    trait_name_style_name = "Trait Name"
    icon_style_name = "Icons"
    auxiliary_style_name = "Auxiliary"
    try:
        trait_name_style = layer.Parent.CharacterStyles.Item(trait_name_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{trait_name_style_name}'")
    try:
        trait_description_style = layer.Parent.CharacterStyles.Item(description_style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{description_style_name}'")
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
        "(CREATURE)": "\uE100",
        "(ACTION)": "\uE101",
        "(AURA)": "\uE102",
        "(FIELD)": "\uE103",
        "(HP)": "\uE200",
        "(ATK)": "\uE201",
        "(SPE)": "\uE202",
        "(NOCOLOR)": "\uE300",
        "(BLACK)": "\uE301",
        "(BLUE)": "\uE302",
        "(CYAN)": "\uE303",
        "(GREEN)": "\uE304",
        "(ORANGE)": "\uE305",
        "(PINK)": "\uE306",
        "(PURPLE)": "\uE307",
        "(WHITE)": "\uE308",
        "(YELLOW)": "\uE309",
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


def create_card_front_base_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
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
            if not card.data.name == "":
                contents = card.data.name
            elif not card.metadata.dev_name == "":
                contents = f"({card.metadata.dev_name})"
        elif text_frame_name == "CostTotalText":
            contents = str(card.data.cost_total)
        elif text_frame_name == "CostColorText":
            contents = str(card.data.cost_color)
        elif text_frame_name == "CostNonColorText":
            contents = str(card.data.cost_total - card.data.cost_color)
        elif text_frame_name == "Identifier":
            if GIT_TAG_NAME is None:
                contents = f"TEST | {card.metadata.id}"
            else:
                contents = f"{GIT_TAG_NAME} | {card.metadata.id}"
        text_frame.Contents = contents


def create_card_front_background_color_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
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
            f"Expected 'BackgroundColorLayer' to have 10 group items, found {layer.GroupItems.Count} instead")
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
                f"Expected background color layer's group item {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In background color layer, found two group items with the name '{group_item_name}'")
        group_item_names_found[group_item_name] = True

        group_item.Hidden = not (group_item_names_to_color[group_item_name] == color)


def create_card_back(card: card_data.Card, document: illustrator.Document) -> None:
    if not document.Layers.Count == 5:
        raise IllustratorTemplateError(f"Expected 5 layers, found {document.Layers.Count} instead")

    non_creature_layer = document.Layers.Item(1)
    if not non_creature_layer.Name == "NonCreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 1st layer's name to be 'NonCreatureLayer', found '{non_creature_layer.Name}' instead")
    non_creature_layer.Visible = False

    creature_layer = document.Layers.Item(2)
    if not creature_layer.Name == "CreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 2nd layer's name to be 'CreatureLayer', found '{creature_layer.Name}' instead")
    creature_layer.Visible = False

    base_layer = document.Layers.Item(3)
    if not base_layer.Name == "BaseLayer":
        raise IllustratorTemplateError(
            f"Expected 3rd layer's name to be 'BaseLayer', found '{creature_layer.Name}' instead")
    base_layer.Visible = False

    background_color_layer = document.Layers.Item(4)
    if not background_color_layer.Name == "BackgroundColorLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundColorLayer', found '{background_color_layer.Name}' instead")
    create_card_back_background_color_layer(card, background_color_layer)

    aux_layer = document.Layers.Item(5)
    if not aux_layer.Name == "AuxLayer":
        raise IllustratorTemplateError(
            f"Expected 5th layer's name to be 'AuxLayer', found '{aux_layer.Name}' instead")
    aux_layer.Visible = False


def create_card_back_background_color_layer(card: card_data.Card, layer: illustrator.Layer) -> None:
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
            f"Expected 'BackgroundColorLayer' to have 10 group items, found {layer.GroupItems.Count} instead"
        )
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
                f"Expected background color layer's group item {i} to have a name in the list "
                f"{list(group_item_names_found.keys())}, found name '{group_item_name}' instead"
            )
        if group_item_names_found[group_item_name]:
            raise IllustratorTemplateError(
                f"In background color layer, found two group items with the name '{group_item_name}'")
        group_item_names_found[group_item_name] = True

        if not group_item_names_to_color[group_item_name] == color:
            group_item.Hidden = True
            continue

        # If we're here, we've found the group whose color matches the card
        group_item.Hidden = False

        # Get background icon
        if not group_item.GroupItems.Count == 1:
            raise IllustratorTemplateError(
                f"Expected group '{group_item_name}' in 'BackgroundColorLayer' to have 1 group item, found "
                f"{group_item.GroupItems.Count} instead"
            )
        if not group_item.GroupItems.Item(1).Name == "Icon":
            raise IllustratorTemplateError(
                f"Expected group '{group_item_name}' in 'BackgroundColorLayer' to have a subgroup named 'Icon', found "
                f"name {group_item.GroupItems.Item(1).Name} instead"
            )
        icon = group_item.GroupItems.Item(1)

        # Move icon to center
        (icon_position_x, icon_position_y) = icon.Position
        icon_width = icon.Width
        icon_height = icon.Height
        document_width = icon.Application.ActiveDocument.Width
        document_height = icon.Application.ActiveDocument.Height
        icon.Translate(
            - icon_position_x + document_width / 2 - icon_width / 2,
            - icon_position_y + document_height / 2 + icon_height / 2
        )

        # Rescale icon
        icon.Resize(200, 200, ScaleAbout=illustrator.constants.aiTransformCenter)

        # Validate path items - START
        if not group_item.PathItems.Count == 3:
            raise IllustratorTemplateError(
                f"Expected group '{group_item_name}' in 'BackgroundColorLayer' to have 3 path items, found "
                f"{group_item.PathItems.Count} instead"
            )
        path_item_names_found = {
            "CostColorBackground": False,
            "Background": False,
            "Border": False,
        }
        for j in range(1, 4):
            path_item = group_item.PathItems.Item(j)
            path_item_name = path_item.Name
            if path_item_name not in path_item_names_found.keys():
                raise IllustratorTemplateError(
                    f"Expected background color layer, group '{group_item_name}', path item {j} to have a name in the "
                    f"list {list(path_item_names_found.keys())}, found name '{path_item_name}' instead"
                )
            if path_item_names_found[path_item_name]:
                raise IllustratorTemplateError(
                    f"In background color layer, group '{group_item_name}', found two subgroups with the name "
                    f"'{path_item_name}'")
            path_item_names_found[path_item_name] = True

            if path_item_name == "CostColorBackground":
                path_item.Hidden = True
            if path_item_name == "Background":
                path_item.Hidden = False
            if path_item_name == "Border":
                path_item.Hidden = False

                # Change icon opacity and color to be the same as border
                icon.Opacity = path_item.Opacity
                for k in range(1, icon.PathItems.Count):
                    icon.PathItems.Item(k).FillColor.Red = path_item.FillColor.Red
                    icon.PathItems.Item(k).FillColor.Green = path_item.FillColor.Green
                    icon.PathItems.Item(k).FillColor.Blue = path_item.FillColor.Blue

                # Change border to gray
                path_item.Opacity = 100
                path_item.FillColor.Red = 120
                path_item.FillColor.Green = 120
                path_item.FillColor.Blue = 120


card_data.import_all_data()

with tempfile.TemporaryDirectory() as temp_dir:
    # card_list = ["E002", "C020", "C049", "C095", "C069", "E032", "E077", "E053", "E061", "E069"]
    card_list = ["E053"]
    for card_id in card_list:
        if card_id[0] == "E":
            card = card_data.Effect.get_effect_dict()[card_id]
        else:
            card = card_data.Creature.get_creature_dict()[card_id]
        create_card(card, Path(temp_dir))

    print("HERE")
