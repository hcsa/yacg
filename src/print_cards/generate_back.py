import src.cards as cards
import src.print_cards.helpers as helpers
import src.print_cards.illustrator_com as illustrator_com
from src.print_cards.configs import ILLUSTRATOR_LAYER


def generate_back(color: cards.Color, document: illustrator_com.Document) -> None:
    non_creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.NON_CREATURE)
    non_creature_layer.Visible = False

    creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.CREATURE)
    creature_layer.Visible = False

    base_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BASE)
    base_layer.Visible = False

    background_color_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BACKGROUND_COLOR)
    _generate_background_color_layer(color, background_color_layer)

    aux_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.AUXILIARY)
    aux_layer.Visible = False


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
        else:
            _generate_background_color_layer_color_group(color, page_item)


def _generate_background_color_layer_color_group(color: cards.Color,
                                                 group_item: illustrator_com.GroupItem) -> None:
    group_item.Hidden = False

    page_item_names = [
        "Icon",
        "Background",
        "Border"
    ]
    if not color == cards.Color.NONE:
        page_item_names.append("CostColorBackground")

    page_items = helpers.get_all_page_items_by_name(group_item, page_item_names)
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
    icon.Resize(200, 200, ScaleAbout=illustrator_com.constants.aiTransformCenter)

    # Change icon opacity and color to be the same as border
    icon.Opacity = page_items["Border"].Opacity
    for k in range(1, icon.PageItems.Count + 1):
        icon.PathItems.Item(k).FillColor = border.FillColor

    # Change border to gray
    border.FillColor = illustrator_com.GrayColor()
    border.FillColor.Gray = 70
