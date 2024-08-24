import os
import re
import shutil
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple, Callable

import pywintypes
import win32com.client.gencache

import scripts.yacg_python.cards as cards
import scripts.yacg_python.illustrator_com as illustrator
from scripts.yacg_python.common_vars import CARD_TEMPLATE_PATH, GIT_TAG_NAME, CARD_ARTS_DIR

# For tokens, a "fake trait" is added describing the token rules
# These configure the "fake trait" name and description
_TOKEN_TRAIT_NAME = "Token Creature"
_TOKEN_TRAIT_DESCRIPTION = \
    (
        "This (CREATURE) can't be in the deck at the start of the game. If it leaves the field, remove it from the "
        "game."
    )


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


def get_all_cards_to_print(
        print_criteria: Callable[[cards.Card], bool] = lambda c: c.is_playable()
) -> List[cards.Card]:
    """
    Returns all cards, sorted canonically and filtered according to the given criteria
    """

    cards_list = cards.get_all_cards()
    cards_list = [c for c in cards_list if print_criteria(c)]
    return cards_list


def get_card_printing_number(
        card: cards.Card,
        print_criteria: Callable[[cards.Card], bool] = lambda c: c.is_playable()
) -> int:
    cards_list = get_all_cards_to_print(print_criteria)

    # Find position of card in the list
    for i, c in enumerate(cards_list, start=1):
        if c.get_id() == card.get_id():
            return i
    raise CardGenerationError(f"Card '{card.get_id()}' doesn't verify the printing criteria")


# AUXILIARY FUNCTIONS - START


def _get_illustrator_app() -> illustrator.Application:
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


def _export_to_tiff(document: illustrator.Document, file_path: Path) -> None:
    export_options = illustrator.ExportOptionsTIFF()
    export_options.AntiAliasing = illustrator.constants.aiArtOptimized
    export_options.ByteOrder = illustrator.constants.aiIBMPC
    export_options.ImageColorSpace = illustrator.constants.aiImageCMYK
    export_options.LZWCompression = False
    export_options.Resolution = 300
    export_options.SaveMultipleArtboards = False

    # We use .with_suffix("") to remove extension, because it's added automatically
    document.Export(file_path.with_suffix(""), illustrator.constants.aiTIFF, export_options)


def _get_style_description(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Description"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def _get_style_trait_title(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Trait Title"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def _get_style_reference_name(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Reference Name"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def _get_style_icons(text_frame: illustrator.TextFrame) -> illustrator.CharacterStyle:
    style_name = "Icons"
    try:
        style = text_frame.Layer.Parent.CharacterStyles.Item(style_name)
    except pywintypes.com_error:
        raise IllustratorTemplateError(f"Failed to find a character style named '{style_name}'")
    return style


def _prepare_text_frame(text_frame: illustrator.TextFrame) -> None:
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


def _get_all_page_items_by_name(parent: Any, page_item_names: List[str]) -> Dict[str, Any]:
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


def _replace_placeholders(text_frame: illustrator.TextFrame) -> Tuple[List[int], List[int]]:
    """
    Replaces all keywords in the text frame with the respective icons, and references with content

    Returns 2 lists:
      - Indices of the character icons (1-indexed)
      - Indices of the characters belonging to references' names (1-indexed)
    """

    # While replacing the placeholders, we'll need to store extra info to return the outputs correctly.
    # Because the placeholders are replaced with text with variable lengths, the starting position of the replacements
    # may be different from the starting position of the placeholders.
    # We'll have to account for this, after doing any replacements.
    # The _replace_placeholders_update_start_indices_ routine makes this adjustment

    # We replace references twice (because references' descriptions may contain references themselves) and then
    # replace keywords. Keywords are done at the end because references may contain keywords

    replacements_data = []
    for _ in range(2):
        replacements_done = _replace_placeholders_refs(text_frame)
        replacements_data = _replace_placeholders_update_start_indices(
            replacements_data + replacements_done,
            replacements_done
        )
    replacements_done = _replace_placeholders_keywords(text_frame)
    replacements_data = _replace_placeholders_update_start_indices(
        replacements_data + replacements_done,
        replacements_done
    )

    icons_characters_indices = []
    reference_names_characters_indices = []
    for replacement_data in sorted(replacements_data, key=lambda x: x.start_index):
        if replacement_data.match_type == "icon":
            icons_characters_indices += list(range(
                replacement_data.start_index,
                replacement_data.start_index + replacement_data.match_after_length
            ))
        elif replacement_data.match_type == "reference_name":
            reference_names_characters_indices += list(range(
                replacement_data.start_index,
                replacement_data.start_index + replacement_data.match_after_length
            ))

    return icons_characters_indices, reference_names_characters_indices


class _ReplacementData:
    """
    Auxiliary class to replace_placeholders.
    Holds data regarding replacements done:
      - What replacement was done (an icon, a name referencing something else, or a description of something else)
        we need to store data
      - The start index (1-indexed)
      - The length of the matched text
      - The length of the name (that is replacing the matched text)
    """

    def __init__(self, match_type: str, start_index: int, match_before_length: int, match_after_length: int):
        if match_type not in ["icon", "reference_name", "reference_description"]:
            raise ValueError(f"Unexpected match type '{match_type}'")
        self.match_type = match_type
        self.start_index = start_index
        self.match_before_length = match_before_length
        self.match_after_length = match_after_length


def _replace_placeholders_keywords(text_frame: illustrator.TextFrame) -> List[_ReplacementData]:
    """
    Aux function to replace_placeholders
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

    match_data = []

    def replacement(match):
        character = keyword_to_character_dict[match.group(0)]

        nonlocal match_data
        match_data.append(_ReplacementData(
            match_type="icon",
            start_index=match.start() + 1,
            match_before_length=len(match.group(0)),
            match_after_length=len(character),
        ))

        return character

    text_frame.Contents = substitute_pattern.sub(
        replacement,
        text_frame.Contents
    )

    return match_data


def _replace_placeholders_refs(text_frame: illustrator.TextFrame) -> List[_ReplacementData]:
    """
    Aux function to replace_placeholders
    """

    substitute_pattern_string = r"\(REF\:(?P<card_data_id>.+?)\.(?P<field_name>.+?)\)"
    substitute_pattern = re.compile(substitute_pattern_string, re.NOFLAG)

    match_data = []

    def replacement(match):
        card_data_id = str(match["card_data_id"])
        card_data = cards.get_card_data(card_data_id)

        field_name = str(match["field_name"]).lower()
        if field_name == "name":
            replacement_text = card_data.get_name()
            if replacement_text == "":
                if isinstance(card_data, cards.Card):
                    replacement_text = f"(Card {card_data_id})"
                else:
                    replacement_text = f"(Trait {card_data_id})"
        elif field_name == "description":
            if not isinstance(card_data, cards.Trait):
                raise CardGenerationError(
                    f"Found reference '{match.group(0)}', but '{card_data_id}' isn't a trait ID"
                )
            replacement_text = card_data.data.description
            # Remove ending full stop, if it has
            if replacement_text[-1] == ".":
                replacement_text = replacement_text[:-1]
        else:
            raise CardGenerationError(
                f"Found reference '{match.group(0)}', but '{field_name}' isn't a known field name"
            )

        nonlocal match_data
        match_data.append(_ReplacementData(
            match_type=("reference_name" if field_name == "name" else "reference_description"),
            start_index=match.start() + 1,
            match_before_length=len(match.group(0)),
            match_after_length=len(replacement_text)
        ))

        return replacement_text

    text_frame.Contents = substitute_pattern.sub(
        replacement,
        text_frame.Contents
    )

    return match_data


def _replace_placeholders_update_start_indices(
        replacements_to_adjust: List[_ReplacementData],
        replacements_done: List[_ReplacementData],
) -> List[_ReplacementData]:
    replacements_to_adjust_ordered = sorted(replacements_to_adjust, key=lambda x: x.start_index)
    replacements_done_ordered = sorted(replacements_done, key=lambda x: x.start_index)

    replacements_adjusted = []
    replacements_done_index = 0
    start_index_offset = 0
    for replacement_data in replacements_to_adjust_ordered:
        # Go over replacements done, until we go over the replacement we're looking at
        while replacements_done_index < len(replacements_done):
            replacement_done = replacements_done_ordered[replacements_done_index]
            if replacement_done.start_index >= replacement_data.start_index:
                break
            # Update offset
            start_index_offset += replacement_done.match_after_length - replacement_done.match_before_length
            replacements_done_index += 1

        start_index_with_offset = replacement_data.start_index + start_index_offset
        replacements_adjusted.append(_ReplacementData(
            match_type=replacement_data.match_type,
            start_index=start_index_with_offset,
            match_before_length=replacement_data.match_before_length,
            match_after_length=replacement_data.match_after_length
        ))

    return replacements_adjusted


# AUXILIARY FUNCTIONS - END

def print_card(
        card: cards.Card,
        output_dir: Path,
        skip_front: bool = False,
        skip_back: bool = False,
        front_file_name: Optional[str] = None,
        back_file_name: Optional[str] = None,
        print_criteria: Callable[[cards.Card], bool] = lambda c: c.is_playable()
) -> None:
    app = _get_illustrator_app()

    # Card front
    if not skip_front:
        card_front_file_name = front_file_name
        if card_front_file_name is None:
            card_front_file_name = f"{card.get_id()}_front"  # Extension is added on export

        card_front_path = output_dir / card_front_file_name
        card_front_document_path = output_dir / f"{card_front_file_name}.temp"

        shutil.copy2(CARD_TEMPLATE_PATH, card_front_document_path)
        card_front_document = app.Open(card_front_document_path)
        _print_card_front(card, card_front_document, print_criteria)
        _export_to_tiff(card_front_document, card_front_path)
        card_front_document.Close(illustrator.constants.aiDoNotSaveChanges)
        os.remove(card_front_document_path)

    # Card back
    if not skip_back:
        card_back_file_name = back_file_name
        if card_back_file_name is None:
            card_back_file_name = f"{card.get_id()}_back"  # Extension is added on export

        card_back_path = output_dir / card_back_file_name  # Extension is added on export
        card_back_document_path = output_dir / f"{card_back_file_name}.temp"

        shutil.copy2(CARD_TEMPLATE_PATH, card_back_document_path)
        card_back_document = app.Open(card_back_document_path)
        _print_card_back(card, card_back_document)
        _export_to_tiff(card_back_document, card_back_path)
        card_back_document.Close(illustrator.constants.aiDoNotSaveChanges)
        os.remove(card_back_document_path)


def _print_card_front(
        card: cards.Card,
        document: illustrator.Document,
        print_criteria: Callable[[cards.Card], bool]
) -> None:
    if not document.Layers.Count == 5:
        raise IllustratorTemplateError(f"Expected 5 layers, found {document.Layers.Count} instead")

    non_creature_layer = document.Layers.Item(1)
    if not non_creature_layer.Name == "NonCreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 1st layer's name to be 'NonCreatureLayer', found '{non_creature_layer.Name}' instead")
    _print_card_front_non_creature_layer(card, non_creature_layer)

    creature_layer = document.Layers.Item(2)
    if not creature_layer.Name == "CreatureLayer":
        raise IllustratorTemplateError(
            f"Expected 2nd layer's name to be 'CreatureLayer', found '{creature_layer.Name}' instead")
    _print_card_front_creature_layer(card, creature_layer)

    base_layer = document.Layers.Item(3)
    if not base_layer.Name == "BaseLayer":
        raise IllustratorTemplateError(
            f"Expected 3rd layer's name to be 'BaseLayer', found '{base_layer.Name}' instead")
    _print_card_front_base_layer(card, base_layer, print_criteria)

    background_color_layer = document.Layers.Item(4)
    if not background_color_layer.Name == "BackgroundColorLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundColorLayer', found '{background_color_layer.Name}' instead")
    _print_card_front_background_color_layer(card, background_color_layer)

    aux_layer = document.Layers.Item(5)
    if not aux_layer.Name == "AuxLayer":
        raise IllustratorTemplateError(
            f"Expected 5th layer's name to be 'AuxLayer', found '{aux_layer.Name}' instead")
    aux_layer.Visible = False


def _print_card_front_non_creature_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, cards.Creature):
        # This layer isn't used by creature cards
        layer.Visible = False
        return

    # If we're here, card is an effect
    card: cards.Effect
    layer.Visible = True

    page_items = _get_all_page_items_by_name(
        layer,
        ["AuraIcon", "ActionIcon", "FieldIcon", "Description"]
    )

    page_items["AuraIcon"].Hidden = (not card.data.type == cards.EffectType.AURA)
    page_items["ActionIcon"].Hidden = (not card.data.type == cards.EffectType.ACTION)
    page_items["FieldIcon"].Hidden = (not card.data.type == cards.EffectType.FIELD)
    _print_card_front_non_creature_layer_description(card, page_items["Description"])


def _print_card_front_non_creature_layer_description(card: cards.Effect, description: illustrator.TextFrame) -> None:
    description_style = _get_style_description(description)
    icons_style = _get_style_icons(description)
    reference_name_style = _get_style_reference_name(description)

    _prepare_text_frame(description)

    description.Contents = card.data.description
    icons_indexes, reference_names_indexes = _replace_placeholders(description)
    for i in range(1, len(description.Contents) + 1):
        if i in icons_indexes:
            icons_style.ApplyTo(description.Characters.Item(i), True)
        elif i in reference_names_indexes:
            reference_name_style.ApplyTo(description.Characters.Item(i), True)
        else:
            description_style.ApplyTo(description.Characters.Item(i), True)


def _print_card_front_creature_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    if isinstance(card, cards.Effect):
        # This layer isn't used by effect cards
        layer.Visible = False
        return

    # If we're here, card is a creature
    card: cards.Creature
    layer.Visible = True

    page_items = _get_all_page_items_by_name(
        layer,
        ["CreatureIcon", "Health", "Attack", "Speed", "Description"]
    )

    page_items["CreatureIcon"].Hidden = False
    _print_card_front_creature_layer_stat_group(card, page_items["Health"])
    _print_card_front_creature_layer_stat_group(card, page_items["Attack"])
    _print_card_front_creature_layer_stat_group(card, page_items["Speed"])
    _print_card_front_creature_layer_description(card, page_items["Description"])


def _print_card_front_creature_layer_stat_group(card: cards.Creature, stat_group: illustrator.GroupItem) -> None:
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


def _print_card_front_creature_layer_description(card: cards.Creature, description: illustrator.TextFrame) -> None:
    description_style = _get_style_description(description)
    trait_name_style = _get_style_trait_title(description)
    reference_name_style = _get_style_reference_name(description)
    icons_style = _get_style_icons(description)

    _prepare_text_frame(description)
    description.Contents = ""

    traits_name = [t.get_name() for t in card.data.traits]
    traits_description = [t.data.description for t in card.data.traits]

    # If the creature is a token, we add a "fake trait"
    if card.data.is_token:
        traits_name.insert(0, _TOKEN_TRAIT_NAME)
        traits_description.insert(0, _TOKEN_TRAIT_DESCRIPTION)

    # Fill in contents
    for trait_index, (trait_name, trait_description) in enumerate(zip(traits_name, traits_description)):
        if trait_index > 0:
            # This isn't the first trait added
            # So we must add a new line to separate from the previous trait
            description.Contents = description.Contents + "\r"

        description.Contents = description.Contents + f"{trait_name} {trait_description}"

    # Format contents by paragraph
    for paragraph_index, trait_name in enumerate(traits_name, start=1):
        paragraph = description.Paragraphs.Item(paragraph_index)

        icons_indexes, reference_names_indexes = _replace_placeholders(paragraph)

        for i in range(1, len(paragraph.Contents) + 1):
            if i < len(trait_name) + 2:
                trait_name_style.ApplyTo(paragraph.Characters.Item(i), True)
            elif i in icons_indexes:
                # Sometimes needs to be applied twice to take effect
                icons_style.ApplyTo(paragraph.Characters.Item(i), True)
                icons_style.ApplyTo(paragraph.Characters.Item(i), True)
            elif i in reference_names_indexes:
                reference_name_style.ApplyTo(paragraph.Characters.Item(i), True)
            else:
                description_style.ApplyTo(paragraph.Characters.Item(i), True)


def _print_card_front_base_layer(
        card: cards.Card,
        layer: illustrator.Layer,
        print_criteria: Callable[[cards.Card], bool]
) -> None:
    layer.Visible = True

    page_items = _get_all_page_items_by_name(
        layer,
        [
            "Title",
            "CostTotalText",
            "CostColorText",
            "CostNonColorText",
            "CostNonColorBackground",
            "ArtClipGroup",
            "Number",
            "Identifier",
            "OuterBorderLine",
            "InnerBorderLine"
        ]
    )

    page_items["Title"].Hidden = False
    page_items["Title"].Contents = card.get_name()

    if isinstance(card, cards.Creature) and card.data.is_token:
        page_items["CostTotalText"].Hidden = True
        page_items["CostColorText"].Hidden = True
        page_items["CostNonColorText"].Hidden = True
        page_items["CostNonColorBackground"].Hidden = True
    else:
        page_items["CostTotalText"].Hidden = False
        page_items["CostTotalText"].Contents = str(card.get_cost_total())

        if card.get_color() == cards.Color.NONE:
            page_items["CostColorText"].Hidden = True
        else:
            page_items["CostColorText"].Hidden = False
            page_items["CostColorText"].Contents = str(card.get_cost_color())

        page_items["CostNonColorText"].Hidden = False
        page_items["CostNonColorText"].Contents = str(card.get_cost_total() - card.get_cost_color())

        page_items["CostNonColorBackground"].Hidden = False

    page_items["Number"].Hidden = False
    number_cards = len(get_all_cards_to_print(print_criteria))
    card_number = get_card_printing_number(card, print_criteria)
    page_items["Number"].Contents = f"{card_number:03d}/{number_cards:03d}"

    page_items["Identifier"].Hidden = False
    git_tag = GIT_TAG_NAME
    if git_tag is None:
        git_tag = "TEST"
    page_items["Identifier"].Contents = f"{git_tag} | {card.get_id()}"

    _print_card_front_base_layer_add_art(card, page_items["ArtClipGroup"])

    page_items["OuterBorderLine"].Hidden = True
    page_items["InnerBorderLine"].Hidden = True


def _print_card_front_base_layer_add_art(card: cards.Card, art_clip_group: illustrator.GroupItem) -> None:
    art_clip_group.Hidden = False

    page_items = _get_all_page_items_by_name(
        art_clip_group,
        [
            "ArtBorder",
            "ArtLinkedFile"
        ]
    )
    art_border = page_items["ArtBorder"]
    art_linked_file = page_items["ArtLinkedFile"]

    art_border.Hidden = False

    art_files = list(CARD_ARTS_DIR.glob(f"{card.get_id()}.*"))
    if len(art_files) == 0:
        # No art file found, just hide the default art
        art_linked_file.Hidden = True
        return
    if len(art_files) > 1:
        raise CardGenerationError(f"Found multiple arts for card {card.get_id()}: {[file.name for file in art_files]}")
    art_file = art_files[0]
    art_linked_file.File = str(art_file)

    # Adjust picture settings
    art_linked_file.Position = art_border.Position
    art_linked_file.Height = art_border.Height
    art_linked_file.Width = art_border.Width


def _print_card_front_background_color_layer(card: cards.Card, layer: illustrator.Layer) -> None:
    layer.Visible = True

    color = card.get_color()

    page_items = _get_all_page_items_by_name(
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
            if (
                    isinstance(card, cards.Creature)
                    and card.data.is_token
                    and page_item.PageItems.Item(i).Name == "CostColorBackground"
            ):
                page_item.PageItems.Item(i).Hidden = True
            else:
                page_item.PageItems.Item(i).Hidden = False


def _print_card_back(color: cards.Color, document: illustrator.Document) -> None:
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
            f"Expected 3rd layer's name to be 'BaseLayer', found '{base_layer.Name}' instead")
    base_layer.Visible = False

    background_color_layer = document.Layers.Item(4)
    if not background_color_layer.Name == "BackgroundColorLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundColorLayer', found '{background_color_layer.Name}' instead")
    _print_card_back_background_color_layer(color, background_color_layer)

    aux_layer = document.Layers.Item(5)
    if not aux_layer.Name == "AuxLayer":
        raise IllustratorTemplateError(
            f"Expected 5th layer's name to be 'AuxLayer', found '{aux_layer.Name}' instead")
    aux_layer.Visible = False


def _print_card_back_background_color_layer(color: cards.Color, layer: illustrator.Layer) -> None:
    layer.Visible = True

    page_items = _get_all_page_items_by_name(
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
            _print_card_back_background_color_layer_color_group(color, page_item)


def _print_card_back_background_color_layer_color_group(color: cards.Color, group_item: illustrator.GroupItem) -> None:
    group_item.Hidden = False

    page_item_names = [
        "Icon",
        "Background",
        "Border"
    ]
    if not color == cards.Color.NONE:
        page_item_names.append("CostColorBackground")

    page_items = _get_all_page_items_by_name(group_item, page_item_names)
    icon = page_items["Icon"]
    border = page_items["Border"]

    page_items["Background"].Hidden = False
    if not color == cards.Color.NONE:
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


def print_blank_card(
        color: cards.Color,
        output_dir: Path,
        skip_front: bool = False,
        skip_back: bool = False,
        front_file_name: Optional[str] = None,
        back_file_name: Optional[str] = None
) -> None:
    app = _get_illustrator_app()

    # Card front
    if not skip_front:
        card_front_file_name = front_file_name
        if card_front_file_name is None:
            card_front_file_name = f"{str(color)}_front"  # Extension is added on export

        card_front_path = output_dir / card_front_file_name
        card_front_document_path = output_dir / f"{card_front_file_name}.temp"

        shutil.copy2(CARD_TEMPLATE_PATH, card_front_document_path)
        card_front_document = app.Open(card_front_document_path)
        _print_blank_card_front(color, card_front_document)
        _export_to_tiff(card_front_document, card_front_path)
        card_front_document.Close(illustrator.constants.aiDoNotSaveChanges)
        os.remove(card_front_document_path)

    # Card back
    if not skip_back:
        card_back_file_name = back_file_name
        if card_back_file_name is None:
            card_back_file_name = f"{str(color)}_back"  # Extension is added on export

        card_back_path = output_dir / card_back_file_name  # Extension is added on export
        card_back_document_path = output_dir / f"{card_back_file_name}.temp"

        shutil.copy2(CARD_TEMPLATE_PATH, card_back_document_path)
        card_back_document = app.Open(card_back_document_path)
        _print_card_back(color, card_back_document)
        _export_to_tiff(card_back_document, card_back_path)
        card_back_document.Close(illustrator.constants.aiDoNotSaveChanges)
        os.remove(card_back_document_path)


def _print_blank_card_front(color: cards.Color, document: illustrator.Document) -> None:
    for i in range(1, document.Layers.Count + 1):
        layer = document.Layers.Item(i)
        layer.Visible = False

    base_layer = document.Layers.Item(3)
    if not base_layer.Name == "BaseLayer":
        raise IllustratorTemplateError(
            f"Expected 3rd layer's name to be 'BaseLayer', found '{base_layer.Name}' instead")
    _print_blank_card_front_base_layer(base_layer)

    background_color_layer = document.Layers.Item(4)
    if not background_color_layer.Name == "BackgroundColorLayer":
        raise IllustratorTemplateError(
            f"Expected 4th layer's name to be 'BackgroundColorLayer', found '{background_color_layer.Name}' instead")
    _print_blank_card_front_background_color_layer(color, background_color_layer)


def _print_blank_card_front_base_layer(layer: illustrator.Layer) -> None:
    layer.Visible = True

    page_items = _get_all_page_items_by_name(
        layer,
        [
            "Title",
            "CostTotalText",
            "CostColorText",
            "CostNonColorText",
            "CostNonColorBackground",
            "ArtClipGroup",
            "Number",
            "Identifier",
            "OuterBorderLine",
            "InnerBorderLine"
        ]
    )

    page_items["Title"].Hidden = True
    page_items["CostTotalText"].Hidden = True
    page_items["CostColorText"].Hidden = True
    page_items["CostNonColorText"].Hidden = True
    page_items["CostNonColorBackground"].Hidden = False
    page_items["Number"].Hidden = True
    page_items["Identifier"].Hidden = True
    page_items["OuterBorderLine"].Hidden = True
    page_items["InnerBorderLine"].Hidden = True

    _print_blank_card_front_base_layer_add_art(page_items["ArtClipGroup"])


def _print_blank_card_front_base_layer_add_art(art_clip_group: illustrator.GroupItem) -> None:
    art_clip_group.Hidden = False

    page_items = _get_all_page_items_by_name(
        art_clip_group,
        [
            "ArtBorder",
            "ArtLinkedFile"
        ]
    )
    page_items["ArtBorder"].Hidden = False
    page_items["ArtLinkedFile"].Hidden = True


def _print_blank_card_front_background_color_layer(color: cards.Color, layer: illustrator.Layer) -> None:
    layer.Visible = True
    page_items = _get_all_page_items_by_name(
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
