from pathlib import Path
from typing import List, Dict, Any

import pywintypes
import win32com.client

import src.cards as cards
import src.print_cards.configs as configs
import src.print_cards.errors as errors
from src.print_cards import illustrator_com as illustrator_com


def get_all_printable_cards() -> List[cards.Card]:
    """
    Returns an ordered list of all printable cards
    """

    return cards.get_all_cards(
        filter_method=cards.FilterMethod.IS_PLAYABLE,
        sort_method=cards.SortMethod.SORT_CANONICAL
    )


def get_card_printing_number(card: cards.Card) -> int:
    """
    Returns the position (1-indexed) of the card in the ordered list of printable cards
    """

    all_cards = get_all_printable_cards()
    for i, c in enumerate(all_cards, start=1):
        if c.get_id() == card.get_id():
            return i
    raise errors.CardPrintError(f"Card '{card.get_id()}' doesn't verify the printing criteria")


def get_illustrator_app() -> illustrator_com.Application:
    """
    If Illustrator is open, returns the opened application.
    Otherwise, opens a new one and returns it.
    """

    try:
        app: illustrator_com.Application = win32com.client.GetActiveObject("Illustrator.Application")._dispobj_
        # If this doesn't return com_error, then an Illustrator app was already open
    except pywintypes.com_error:
        # The Illustrator app wasn't open, open it
        app: illustrator_com.Application = win32com.client.Dispatch("Illustrator.Application")
    return app


def export_to_tiff(document: illustrator_com.Document, file_path: Path) -> None:
    """
    Exports the document to a .tiff file
    """

    # We use .with_suffix("") to remove extension, because it's added automatically
    document.Export(
        file_path.with_suffix(""),
        illustrator_com.constants.aiTIFF,
        configs.get_export_options()
    )


def get_style(text_frame: illustrator_com.TextFrame, style: configs.ILLUSTRATOR_STYLE):
    try:
        output = text_frame.Layer.Parent.CharacterStyles.Item(style)
    except pywintypes.com_error:
        raise errors.IllustratorTemplateError(f"Failed to find a character style named '{style}'")
    return output


def prepare_text_frame(text_frame: illustrator_com.TextFrame) -> None:
    """
    Sets the character style of the text frame to the auxiliary character style.

    After changing the contents of this text frame, YOU MUST SET ITS STYLE.

    -----------------------------------------------------------------------------------------------------------------

    This function should be called before editing text frames where either:

    - The text frame's contents have multiple styles. Changing the contents of text that has multiple styles makes
      Illustrator confused, so it's needed to have a single style before changing contents.
    - The text frame has/will have icons. The icons are encoded as special Unicode characters that almost all fonts
      don't implement. If a character is set to a style whose font doesn't have that character, bugs happen. The
      auxiliary character style implements these special characters.
    """

    style = get_style(text_frame, configs.ILLUSTRATOR_STYLE.AUXILIARY)
    style.ApplyTo(text_frame.TextRange, True)


def get_layer(document: illustrator_com.Document, layer: configs.ILLUSTRATOR_LAYER) -> illustrator_com.Layer:
    number_layers = len(configs.ILLUSTRATOR_LAYER_LIST)
    if not document.Layers.Count == number_layers:
        raise errors.IllustratorTemplateError(
            f"Expected {number_layers} layers, found {document.Layers.Count} instead"
        )

    if layer not in configs.ILLUSTRATOR_LAYER_LIST:
        raise errors.CardPrintError(
            f"Layer '{layer}' is not in the implementation's list of layers"
        )

    index = configs.ILLUSTRATOR_LAYER_LIST.index(layer) + 1
    output = document.Layers.Item(index)
    if not output.Name == layer:
        raise errors.IllustratorTemplateError(
            f"Expected layer in position {index} to have name '{layer}', found '{output.name}' instead"
        )

    return output


def get_all_page_items_by_name(parent: Any, page_item_names: List[str]) -> Dict[str, Any]:
    """
    Returns a dict with the page items name as keys and the page items as values.

    - Parent must have the "PageItems" property.
    - The list of page item names must be exhaustive.
    """

    parent_name = parent.Name

    if not parent.PageItems.Count == len(page_item_names):
        raise errors.IllustratorTemplateError(
            f"Object '{parent_name}': expected {len(page_item_names)} page items, found {parent.PageItems.Count} "
            f"instead"
        )

    page_items_dict: Dict[str, Any] = {}
    for i in range(1, len(page_item_names) + 1):
        page_item = parent.PageItems.Item(i)
        page_item_name = page_item.Name
        if page_item_name not in page_item_names:
            raise errors.IllustratorTemplateError(
                f"Object '{parent_name}', page item with index {i}: expected to have a name in the list "
                f"{page_item_names}, found name '{page_item_name}' instead"
            )
        if page_item_name in page_items_dict:
            raise errors.IllustratorTemplateError(
                f"Object '{parent_name}': found two page items with the name '{page_item_name}'"
            )
        page_items_dict[page_item_name] = page_item

    return page_items_dict
