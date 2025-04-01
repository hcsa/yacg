import src.cards as cards
import src.print_cards.errors as errors
import src.print_cards.helpers as helpers
import src.print_cards.helpers_replacement as helpers_replacement
import src.print_cards.illustrator_com as illustrator_com
from src.print_cards.configs import (ILLUSTRATOR_LAYER, ILLUSTRATOR_STYLE, TOKEN_TRAIT_NAME,
                                     TOKEN_TRAIT_DESCRIPTION)
from src.utils import GIT_TAG_NAME, CARD_ARTS_DIR


def generate_front(card: cards.Card, document: illustrator_com.Document) -> None:
    non_creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.NON_CREATURE)
    _generate_non_creature_layer(card, non_creature_layer)

    creature_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.CREATURE)
    _generate_creature_layer(card, creature_layer)

    base_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BASE)
    _generate_base_layer(card, base_layer)

    background_color_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.BACKGROUND_COLOR)
    _generate_background_color_layer(card, background_color_layer)

    aux_layer = helpers.get_layer(document, ILLUSTRATOR_LAYER.AUXILIARY)
    aux_layer.Visible = False


def _generate_non_creature_layer(card: cards.Card, layer: illustrator_com.Layer) -> None:
    if isinstance(card, cards.Creature):
        # This layer isn't used by creature cards
        layer.Visible = False
        return

    # If we're here, card is an effect
    card: cards.Effect
    layer.Visible = True

    page_items = helpers.get_all_page_items_by_name(
        layer,
        ["AuraIcon", "ActionIcon", "FieldIcon", "Description"]
    )

    page_items["AuraIcon"].Hidden = (not card.data.type == cards.EffectType.AURA)
    page_items["ActionIcon"].Hidden = (not card.data.type == cards.EffectType.ACTION)
    page_items["FieldIcon"].Hidden = (not card.data.type == cards.EffectType.FIELD)
    _generate_non_creature_layer_description(card, page_items["Description"])


def _generate_non_creature_layer_description(card: cards.Effect, description: illustrator_com.TextFrame) -> None:
    description_style = helpers.get_style(description, ILLUSTRATOR_STYLE.DESCRIPTION)
    icons_style = helpers.get_style(description, ILLUSTRATOR_STYLE.ICONS)
    reference_name_style = helpers.get_style(description, ILLUSTRATOR_STYLE.REFERENCE_NAME)

    helpers.prepare_text_frame(description)

    description.Contents = card.data.description
    icons_indexes, reference_names_indexes = helpers_replacement.replace_placeholders(description)
    for i in range(1, len(description.Contents) + 1):
        if i in icons_indexes:
            icons_style.ApplyTo(description.Characters.Item(i), True)
        elif i in reference_names_indexes:
            reference_name_style.ApplyTo(description.Characters.Item(i), True)
        else:
            description_style.ApplyTo(description.Characters.Item(i), True)


def _generate_creature_layer(card: cards.Card, layer: illustrator_com.Layer) -> None:
    if isinstance(card, cards.Effect):
        # This layer isn't used by effect cards
        layer.Visible = False
        return

    # If we're here, card is a creature
    card: cards.Creature
    layer.Visible = True

    page_items = helpers.get_all_page_items_by_name(
        layer,
        ["CreatureIcon", "Health", "Attack", "Speed", "Description"]
    )

    page_items["CreatureIcon"].Hidden = False
    _generate_creature_layer_stat_group(card, page_items["Health"])
    _generate_creature_layer_stat_group(card, page_items["Attack"])
    _generate_creature_layer_stat_group(card, page_items["Speed"])
    _generate_creature_layer_description(card, page_items["Description"])


def _generate_creature_layer_stat_group(card: cards.Creature, stat_group: illustrator_com.GroupItem) -> None:
    stat_group_name = stat_group.Name

    if not stat_group.TextFrames.Count == 1:
        raise errors.IllustratorTemplateError(
            f"Layer '{ILLUSTRATOR_LAYER.CREATURE}', page item '{stat_group_name}': expected 1 text frame, found "
            f"{stat_group.TextFrames.Count} instead"
        )

    stat_text_frame = stat_group.TextFrames.Item(1)
    if stat_group_name == "Health":
        stat_text_frame.Contents = str(card.data.hp)
    elif stat_group_name == "Attack":
        stat_text_frame.Contents = str(card.data.atk_strong)
    elif stat_group_name == "Speed":
        stat_text_frame.Contents = str(card.data.spe)


def _generate_creature_layer_description(card: cards.Creature, description: illustrator_com.TextFrame) -> None:
    description_style = helpers.get_style(description, ILLUSTRATOR_STYLE.DESCRIPTION)
    trait_name_style = helpers.get_style(description, ILLUSTRATOR_STYLE.TRAIT)
    icons_style = helpers.get_style(description, ILLUSTRATOR_STYLE.ICONS)
    reference_name_style = helpers.get_style(description, ILLUSTRATOR_STYLE.REFERENCE_NAME)

    helpers.prepare_text_frame(description)
    description.Contents = ""

    traits_name = [t.get_name() for t in card.data.traits]
    traits_description = [t.data.description for t in card.data.traits]

    # If the creature is a token, we add a "fake trait"
    if card.data.is_token:
        traits_name.insert(0, TOKEN_TRAIT_NAME)
        traits_description.insert(0, TOKEN_TRAIT_DESCRIPTION)

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

        icons_indexes, reference_names_indexes = helpers_replacement.replace_placeholders(paragraph)

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


def _generate_base_layer(card: cards.Card, layer: illustrator_com.Layer) -> None:
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
    number_cards = len(helpers.get_all_printable_cards())
    card_number = helpers.get_card_printing_number(card)
    page_items["Number"].Contents = f"{card_number:03d}/{number_cards:03d}"

    page_items["Identifier"].Hidden = False
    git_tag = GIT_TAG_NAME
    if git_tag is None:
        git_tag = "TEST"
    page_items["Identifier"].Contents = f"{git_tag} | {card.get_id()}"

    _generate_base_layer_add_art(card, page_items["ArtClipGroup"])

    page_items["OuterBorderLine"].Hidden = True
    page_items["InnerBorderLine"].Hidden = True


def _generate_base_layer_add_art(card: cards.Card, art_clip_group: illustrator_com.GroupItem) -> None:
    art_clip_group.Hidden = False

    page_items = helpers.get_all_page_items_by_name(
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
        raise errors.CardPrintError(
            f"Found multiple arts for card {card.get_id()}: {[file.name for file in art_files]}"
        )
    art_file = art_files[0]
    art_linked_file.File = str(art_file)

    # Adjust picture settings
    art_linked_file.Position = art_border.Position
    art_linked_file.Height = art_border.Height
    art_linked_file.Width = art_border.Width


def _generate_background_color_layer(card: cards.Card, layer: illustrator_com.Layer) -> None:
    layer.Visible = True

    color = card.get_color()

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
            if (
                    isinstance(card, cards.Creature)
                    and card.data.is_token
                    and page_item.PageItems.Item(i).Name == "CostColorBackground"
            ):
                page_item.PageItems.Item(i).Hidden = True
            else:
                page_item.PageItems.Item(i).Hidden = False
