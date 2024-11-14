import src.cards as cards
import src.print_cards.helpers as helpers
import src.print_cards.illustrator_com as illustrator_com
from src.print_cards.configs import ILLUSTRATOR_LAYER


def generate_blank_front(color: cards.Color, document: illustrator_com.Document) -> None:
    non_creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.NON_CREATURE)
    non_creature_layer.Visible = False

    creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.CREATURE)
    creature_layer.Visible = False

    base_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BASE)
    _generate_base_layer(base_layer)

    background_color_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BACKGROUND_COLOR)
    _generate_background_color_layer(color, background_color_layer)

    aux_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.AUXILIARY)
    aux_layer.Visible = False


def _generate_base_layer(layer: illustrator_com.Layer) -> None:
    layer.Visible = True

    page_items = helpers.get_all_page_items_by_name(
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

    _generate_base_layer_add_art(page_items["ArtClipGroup"])


def _generate_base_layer_add_art(art_clip_group: illustrator_com.GroupItem) -> None:
    art_clip_group.Hidden = False

    page_items = helpers.get_all_page_items_by_name(
        art_clip_group,
        [
            "ArtBorder",
            "ArtLinkedFile"
        ]
    )
    page_items["ArtBorder"].Hidden = False
    page_items["ArtLinkedFile"].Hidden = True


def _generate_background_color_layer(color: cards.Color, layer: illustrator_com.Layer) -> None:
    layer.Visible = True
    page_items = helpers.get_all_page_items_by_name(
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
