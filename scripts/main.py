import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import List, Dict, Any

import pywintypes
import win32com.client.gencache

import scripts.yacg_python.cards as cards
import scripts.yacg_python.illustrator_com as illustrator
from scripts.yacg_python.common_vars import CARD_TEMPLATE_PATH, GIT_TAG_NAME, CART_ART_DIR


class CardGenerationError(ValueError):
    """
    Raised when something unexpected happened while generating a card
    """

    def __init__(self, message):
        super().__init__(message)


class IllustratorTemplateError(CardGenerationError):
    """
    Raised when an unexpected element or setting of an Illustrator template is found
    """

    def __init__(self, message):
        super().__init__(message)


def get_illustrator_app() -> illustrator.Application:
    """
    If Illustrator is open, returns the opened application.
    Otherwise, opens a new one and returns it.
    """

    try:
        app: illustrator.Application = win32com.client.GetActiveObject("Illustrator.Application")._dispobj_
        # If this doesn't return com_error, then an Illustrator app was already open
    except pywintypes.com_error:
        # The Illustrator app wasn't open, open it
        app: illustrator.Application = win32com.client.Dispatch("Illustrator.Application")
    return app


def export_to_tiff(document: illustrator.Document, file_path: Path) -> None:
    export_options = illustrator.ExportOptionsTIFF()
    export_options.AntiAliasing = illustrator.constants.aiArtOptimized
    export_options.ByteOrder = illustrator.constants.aiIBMPC
    export_options.ImageColorSpace = illustrator.constants.aiImageCMYK
    export_options.LZWCompression = False
    export_options.Resolution = 400
    export_options.SaveMultipleArtboards = False

    # We use .with_suffix("") to remove extension, because it's added automatically
    document.Export(file_path.with_suffix(""), illustrator.constants.aiTIFF, export_options)


def get_style_description(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Description"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def get_style_trait_name(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Trait Name"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def get_style_icons(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Icons"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def prepare_text_frame(text_frame: illustrator.TextFrame) -> None:
    """
    Sets the character style of the text frame to the auxiliary character style.

    This function should be called when either:

    - The text frame's contents have multiple styles. Changing the contents of text that has multiple styles makes
      Illustrator confused, so it's needed to have a single style before changing contents.

    - The text frame will have icons. The icons are encoded as special Unicode characters that almost all fonts don't
      implement. If a character is set to a style whose font doesn't have that character, bugs happen. The auxiliary
      character style implements these special characters.
    """

    style_name = "Auxiliary"
    try:
        style: illustrator.CharacterStyle = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    style.ApplyTo(text_frame.TextRange, True)


def get_all_page_items_by_name(parent: Any, page_item_names: List[str]) -> Dict[str, Any]:
    """
    Returns a dict with the page items name as keys and the page items as values.

    - Parent must have the "PageItems" property.
    - The list of page item names must be exhaustive.
    """

    parent_name = parent.Name

    if not parent.PageItems.Count == len(page_item_names):
        raise IllustratorTemplateError(
            f"Object '{parent_name}': expected {len(page_item_names)} page items, found {parent.PageItems.Count} "
            f"instead"
        )

    page_items_dict: Dict[str, Any] = {}
    for i in range(1, len(page_item_names) + 1):
        page_item = parent.PageItems.Item(i)
        page_item_name = page_item.Name
        if page_item_name not in page_item_names:
            raise IllustratorTemplateError(
                f"Object '{parent_name}', page item with index {i}: expected to have a name in the list "
                f"{page_item_names}, found name '{page_item_name}' instead"
            )
        if page_item_name in page_items_dict:
            raise IllustratorTemplateError(
                f"bject '{parent_name}': found two page items with the name '{page_item_name}'"
            )
        page_items_dict[page_item_name] = page_item

    return page_items_dict


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


def replace_card_refs(text_frame: illustrator.TextFrame) -> None:
    """
    Replaces all references in card with the corresponding name
    """

    substitute_pattern_string = r"\(REF\:(?P<card_data_id>.+?)\)"
    substitute_pattern = re.compile(substitute_pattern_string, re.NOFLAG)

    def replacement(match):
        card_data_id = match["card_data_id"]
        card_data = cards.get_card_data(card_data_id)
        name = card_data.get_name()
        if name == "":
            if isinstance(card_data, cards.Card):
                return f"(Card {card_data_id})"
            else:
                return f"(Trait {card_data_id})"
        return name

    text_frame.Contents = substitute_pattern.sub(
        replacement,
        text_frame.Contents
    )


def create_card(card: cards.Card, output_dir: Path) -> None:
    app = get_illustrator_app()

    # Card front
    card_front_path = output_dir / f"{card.get_id()}_front"  # Extension is added on export
    card_front_document_path = output_dir / f"{card.get_id()}_front.temp"

    shutil.copy2(CARD_TEMPLATE_PATH, card_front_document_path)
    card_front_document = app.Open(card_front_document_path)
    create_card_front(card, card_front_document)
    export_to_tiff(card_front_document, card_front_path)
    card_front_document.Close(illustrator.constants.aiDoNotSaveChanges)
    os.remove(card_front_document_path)

    # Card back
    card_back_path = output_dir / f"{card.get_id()}_back"  # Extension is added on export
    card_back_document_path = output_dir / f"{card.get_id()}_back.temp"

    shutil.copy2(CARD_TEMPLATE_PATH, card_back_document_path)
    card_back_document = app.Open(card_back_document_path)
    create_card_back(card, card_back_document)
    export_to_tiff(card_back_document, card_back_path)
    card_back_document.Close(illustrator.constants.aiDoNotSaveChanges)
    os.remove(card_back_document_path)


def create_card_front(card: cards.Card, document: illustrator.Document) -> None:
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


def create_card_front_non_creature_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, cards.Creature):
        # This layer isn't used by creature cards
        layer.Visible = False
        return

    # If we're here, card is an effect
    card: cards.Effect
    layer.Visible = True

    page_items = get_all_page_items_by_name(
        layer,
        ["AuraIcon", "ActionIcon", "FieldIcon", "Description"]
    )

    page_items["AuraIcon"].Hidden = (not card.data.type == cards.EffectType.AURA)
    page_items["ActionIcon"].Hidden = (not card.data.type == cards.EffectType.ACTION)
    page_items["FieldIcon"].Hidden = (not card.data.type == cards.EffectType.FIELD)
    create_card_front_non_creature_layer_description(card, page_items["Description"])


def create_card_front_non_creature_layer_description(card: cards.Effect, description: illustrator.TextFrame) -> None:
    description_style = get_style_description(description)
    icons_style = get_style_icons(description)

    prepare_text_frame(description)

    description.Contents = card.data.description
    replace_card_refs(description)
    icons_indexes = replace_keywords_with_icons(description)
    for i in range(1, len(description.Contents) + 1):
        if i in icons_indexes:
            icons_style.ApplyTo(description.Characters.Item(i), True)
        else:
            description_style.ApplyTo(description.Characters.Item(i), True)


def create_card_front_creature_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, cards.Effect):
        # This layer isn't used by effect cards
        layer.Visible = False
        return

    # If we're here, card is a creature
    card: cards.Creature
    layer.Visible = True

    page_items = get_all_page_items_by_name(
        layer,
        ["CreatureIcon", "Health", "Attack", "Speed", "Description"]
    )

    page_items["CreatureIcon"].Hidden = False
    create_card_front_creature_layer_stat_group(card, page_items["Health"])
    create_card_front_creature_layer_stat_group(card, page_items["Attack"])
    create_card_front_creature_layer_stat_group(card, page_items["Speed"])
    create_card_front_creature_layer_description(card, page_items["Description"])


def create_card_front_creature_layer_stat_group(card: cards.Creature, stat_group: illustrator.GroupItem) -> None:
    stat_group_name = stat_group.Name

    if not stat_group.TextFrames.Count == 1:
        raise IllustratorTemplateError(
            f"Layer 'CreatureLayer', page item '{stat_group_name}': expected 1 text frame, found "
            f"{stat_group.TextFrames.Count} instead"
        )

    stat_text_frame = stat_group.TextFrames.Item(1)
    if stat_group_name == "Health":
        stat_text_frame.Contents = str(card.data.hp)
    elif stat_group_name == "Attack":
        stat_text_frame.Contents = str(card.data.atk)
    elif stat_group_name == "Speed":
        stat_text_frame.Contents = str(card.data.spe)


def create_card_front_creature_layer_description(card: cards.Creature, description: illustrator.TextFrame) -> None:
    description_style = get_style_description(description)
    trait_name_style = get_style_trait_name(description)
    icons_style = get_style_icons(description)

    prepare_text_frame(description)
    description.Contents = ""

    # Stores trait names. Will be used later to format the trait names
    trait_names: List[str] = []

    # Fill in contents
    for trait_index, trait in enumerate(card.data.traits, start=1):
        if trait_index > 1:
            # This isn't the first trait added
            # So we must add a new line to separate from the previous trait
            description.Contents = description.Contents + "\r"

        # Add trait name and description
        trait_name = trait.get_name()
        description.Contents = description.Contents + f"{trait_name} {trait.data.description}"

        trait_names.append(trait_name)

    # Format contents by paragraph
    for paragraph_index, trait_name in enumerate(trait_names, start=1):
        paragraph = description.Paragraphs.Item(paragraph_index)

        replace_card_refs(paragraph)
        icons_indexes = replace_keywords_with_icons(paragraph)

        for i in range(1, len(trait_name) + 2):  # Includes the space after trait name
            trait_name_style.ApplyTo(paragraph.Characters.Item(i), True)

        for i in range(len(trait_name) + 2, len(paragraph.Contents) + 1):
            if i in icons_indexes:
                # Sometimes needs to be applied twice to take effect
                icons_style.ApplyTo(paragraph.Characters.Item(i), True)
                icons_style.ApplyTo(paragraph.Characters.Item(i), True)
            else:
                description_style.ApplyTo(paragraph.Characters.Item(i), True)


def create_card_front_base_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    layer.Visible = True

    page_items = get_all_page_items_by_name(
        layer,
        [
            "Title",
            "CostTotalText",
            "CostColorText",
            "CostNonColorText",
            "CostNonColorBackground",
            "ArtClipGroup",
            "Identifier",
            "OuterBorderLine",
            "InnerBorderLine"
        ]
    )

    page_items["Title"].Hidden = False
    page_items["Title"].Contents = card.get_name()

    page_items["CostTotalText"].Hidden = False
    page_items["CostTotalText"].Contents = str(card.get_cost_total())

    if card.get_color() == cards.Color.NONE:
        page_items["CostColorText"].Hidden = True
    else:
        page_items["CostColorText"].Hidden = False
        page_items["CostColorText"].Contents = str(card.get_cost_color())

    page_items["CostNonColorText"].Hidden = False
    page_items["CostNonColorText"].Contents = str(card.get_cost_total() - card.get_cost_color())

    page_items["Identifier"].Hidden = False
    if GIT_TAG_NAME is None:
        page_items["Identifier"].Contents = f"TEST | {card.get_id()}"
    else:
        page_items["Identifier"].Contents = f"{GIT_TAG_NAME} | {card.get_id()}"

    create_card_front_base_layer_add_art(card, page_items["ArtClipGroup"])

    page_items["CostNonColorBackground"].Hidden = False
    page_items["OuterBorderLine"].Hidden = True
    page_items["InnerBorderLine"].Hidden = True


def create_card_front_base_layer_add_art(card: cards.Card, art_clip_group: illustrator.GroupItem) -> None:
    art_clip_group.Hidden = False

    page_items = get_all_page_items_by_name(
        art_clip_group,
        [
            "ArtBorder",
            "ArtLinkedFile"
        ]
    )
    page_items["ArtBorder"].Hidden = False
    art_linked_file = page_items["ArtLinkedFile"]

    art_files = list(CART_ART_DIR.glob(f"{card.get_id()}.*"))
    if len(art_files) == 0:
        # No art file found, just hide the default art
        art_linked_file.Hidden = True
        return
    if len(art_files) > 1:
        raise CardGenerationError(f"Found multiple arts for card {card.get_id()}: {[file.name for file in art_files]}")
    art_file = art_files[0]
    art_linked_file.File = str(art_file)

    # Adjust picture settings
    art_linked_file.Position = art_clip_group.Position
    art_linked_file.Height = art_clip_group.Height
    art_linked_file.Width = art_clip_group.Width


def create_card_front_background_color_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    color = card.get_color()

    page_items = get_all_page_items_by_name(
        layer,
        [
            "BackgroundNone",
            "BackgroundBlack",
            "BackgroundBlue",
            "BackgroundCyan",
            "BackgroundGreen",
            "BackgroundOrange",
            "BackgroundPink",
            "BackgroundPurple",
            "BackgroundWhite",
            "BackgroundYellow",
        ]
    )
    page_items_to_color = {
        "BackgroundNone": cards.Color.NONE,
        "BackgroundBlack": cards.Color.BLACK,
        "BackgroundBlue": cards.Color.BLUE,
        "BackgroundCyan": cards.Color.CYAN,
        "BackgroundGreen": cards.Color.GREEN,
        "BackgroundOrange": cards.Color.ORANGE,
        "BackgroundPink": cards.Color.PINK,
        "BackgroundPurple": cards.Color.PURPLE,
        "BackgroundWhite": cards.Color.WHITE,
        "BackgroundYellow": cards.Color.YELLOW,
    }

    for page_item_name, page_item in page_items.items():
        if not page_items_to_color[page_item_name] == color:
            page_item.Hidden = True
            continue

        # If we're here, the page item color matches the card's color
        page_item.Hidden = False
        for i in range(1, page_item.PageItems.Count + 1):
            page_item.PageItems.Item(i).Hidden = False


def create_card_back(card: cards.Card, document: illustrator.Document) -> None:
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


def create_card_back_background_color_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    layer.Visible = True

    color = card.get_color()

    page_items = get_all_page_items_by_name(
        layer,
        [
            "BackgroundNone",
            "BackgroundBlack",
            "BackgroundBlue",
            "BackgroundCyan",
            "BackgroundGreen",
            "BackgroundOrange",
            "BackgroundPink",
            "BackgroundPurple",
            "BackgroundWhite",
            "BackgroundYellow",
        ]
    )
    page_items_to_color = {
        "BackgroundNone": cards.Color.NONE,
        "BackgroundBlack": cards.Color.BLACK,
        "BackgroundBlue": cards.Color.BLUE,
        "BackgroundCyan": cards.Color.CYAN,
        "BackgroundGreen": cards.Color.GREEN,
        "BackgroundOrange": cards.Color.ORANGE,
        "BackgroundPink": cards.Color.PINK,
        "BackgroundPurple": cards.Color.PURPLE,
        "BackgroundWhite": cards.Color.WHITE,
        "BackgroundYellow": cards.Color.YELLOW,
    }

    for page_item_name, page_item in page_items.items():
        if not page_items_to_color[page_item_name] == color:
            page_item.Hidden = True
        else:
            create_card_back_background_color_layer_color_group(card, page_item)


def create_card_back_background_color_layer_color_group(card: cards.Card, group_item: illustrator.GroupItem) -> None:
    group_item.Hidden = False

    page_item_names = [
        "Icon",
        "Background",
        "Border"
    ]
    if not card.get_color() == cards.Color.NONE:
        page_item_names.append("CostColorBackground")

    page_items = get_all_page_items_by_name(group_item, page_item_names)
    icon = page_items["Icon"]
    border = page_items["Border"]

    page_items["Background"].Hidden = False
    if not card.get_color() == cards.Color.NONE:
        page_items["CostColorBackground"].Hidden = True
    icon.Hidden = False
    border.Hidden = False

    # Move icon to center
    (icon_position_x, icon_position_y) = icon.Position
    icon_width = icon.Width
    icon_height = icon.Height
    document_width = icon.Application.ActiveDocument.Width
    document_height = icon.Application.ActiveDocument.Height
    icon.Translate(
        - icon_position_x + document_width / 2 - icon_width / 2,
        - icon_position_y - document_height / 2 + icon_height / 2
    )

    # Rescale icon
    icon.Resize(200, 200, ScaleAbout=illustrator.constants.aiTransformCenter)

    # Change icon opacity and color to be the same as border
    icon.Opacity = page_items["Border"].Opacity
    for k in range(1, icon.PageItems.Count + 1):
        icon.PathItems.Item(k).FillColor = border.FillColor

    # Change border to gray
    border.FillColor = illustrator.GrayColor()
    border.FillColor.Gray = 70


cards.import_all_data()

with tempfile.TemporaryDirectory() as temp_dir:
    print(temp_dir)
    card_list = ["E002", "C020", "C049", "C095", "C069", "E032", "E077", "E050", "E061", "E069"]
    # card_list = ["C049"]
    for card_id in card_list:
        c = cards.get_card(card_id)
        create_card(c, Path(temp_dir))

    print("HERE")
