from enum import StrEnum

import src.print_cards.illustrator_com as illustrator_com

# This is needed because EXPORT_OPTIONS_TIFF references instances of COM classes
# helpers_get_app.get_illustrator_app()

# For tokens, a "fake trait" is added describing the token rules
# The following configure the "fake trait" name and description
TOKEN_TRAIT_NAME = "Token Creature"
TOKEN_TRAIT_DESCRIPTION = \
    (
        "This (CREATURE) can't be in the deck at the start of the game. If it leaves the field, remove it from the "
        "game."
    )


# Values must match the layer names in Illustrator
class ILLUSTRATOR_LAYER(StrEnum):
    NON_CREATURE = "NonCreatureLayer"
    CREATURE = "CreatureLayer"
    BASE = "BaseLayer"
    BACKGROUND_COLOR = "BackgroundColorLayer"
    AUXILIARY = "AuxLayer"


# Order must match the layers' order in Illustrator
ILLUSTRATOR_LAYER_LIST = [
    ILLUSTRATOR_LAYER.NON_CREATURE,
    ILLUSTRATOR_LAYER.CREATURE,
    ILLUSTRATOR_LAYER.BASE,
    ILLUSTRATOR_LAYER.BACKGROUND_COLOR,
    ILLUSTRATOR_LAYER.AUXILIARY
]


# Values must match the style names in Illustrator
class ILLUSTRATOR_STYLE(StrEnum):
    DESCRIPTION = "Description"
    TRAIT = "Trait Title"
    REFERENCE_NAME = "Reference Name"
    ICONS = "Icons"
    AUXILIARY = "Auxiliary"


# List of keywords in the text, mapped to the unicode character they should be replaced with
KEYWORD_TO_CHARACTER_DICT = {
    "(CREATURE)": "\uE100",
    "(ACTION)": "\uE101",
    "(AURA)": "\uE102",
    "(FIELD)": "\uE103",

    "(HP)": "\uE200",
    "(SATK)": "\uE201",
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


# This is a function so that execution is delayed until the Illustrator app is loaded (otherwise initializing export
# options will throw an error
def get_export_options() -> illustrator_com.ExportOptionsTIFF:
    export_options = illustrator_com.ExportOptionsTIFF()
    export_options.AntiAliasing = illustrator_com.constants.aiArtOptimized
    export_options.ByteOrder = illustrator_com.constants.aiIBMPC
    export_options.ImageColorSpace = illustrator_com.constants.aiImageCMYK
    export_options.LZWCompression = False
    export_options.Resolution = 300
    export_options.SaveMultipleArtboards = False
    return export_options
