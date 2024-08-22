"""
Run this script to generate all playable cards
"""

import scripts.yacg_python.card_printing as card_printing
import scripts.yacg_python.cards as cards
from scripts.yacg_python.common_vars import BASE_DIR

OUTPUT_DIR = BASE_DIR / "print_output"

DUPLICATE_CARDS = {
    # No color
    "E005": 3,  # Charge
    "E089": 3,  # Double Up
    "E002": 3,  # Counter
    "E004": 3,  # Restart
    "E003": 3,  # Search
    "E001": 3,  # Brainstorm

    # Pink
    "E048": 2,  # Charge
    "E090": 2,  # Double Up
    "E044": 2,  # Counter
    "E046": 2,  # Restart
    "E049": 2,  # Search
    "E043": 2,  # Brainstorm

    # Black
    "C002": 2,  # Charge
    "E091": 2,  # Double Up
    "E056": 2,  # Counter
    "E058": 2,  # Restart
    "E057": 2,  # Search
    "E055": 2,  # Brainstorm

    # Cyan
    "E069": 2,  # Charge
    "E092": 2,  # Double Up
    "E065": 2,  # Counter
    "E067": 2,  # Restart
    "E066": 2,  # Search
    "E064": 2,  # Brainstorm
}

EMPTY_CARDS = {
    cards.Color.NONE: 2,
    cards.Color.ORANGE: 1,
    cards.Color.GREEN: 1,
    cards.Color.BLUE: 1,
    cards.Color.WHITE: 1,
    cards.Color.YELLOW: 1,
    cards.Color.PURPLE: 1,
    cards.Color.PINK: 1,
    cards.Color.BLACK: 1,
    cards.Color.CYAN: 1,
}

cards.import_all_data()

output_dir_fronts = OUTPUT_DIR / "fronts"
output_dir_backs = OUTPUT_DIR / "backs"

print_index = 1
for card in card_printing.get_all_cards_to_print():
    card_id = card.get_id()
    if card_id in DUPLICATE_CARDS:
        card_number_copies = DUPLICATE_CARDS[card_id]
    else:
        card_number_copies = 1

    for copy_index in range(1, card_number_copies + 1):
        file_name = f"{print_index:03d}_{card_id}.{copy_index}"
        if len(list(output_dir_fronts.glob(f"{file_name}*"))) == 0:
            card_printing.print_card(
                card,
                output_dir_fronts,
                front_file_name=f"{file_name}_front",
                skip_back=True
            )
        if len(list(output_dir_backs.glob(f"{file_name}*"))) == 0:
            card_printing.print_card(
                card,
                output_dir_backs,
                back_file_name=f"{file_name}_back",
                skip_front=True
            )
        print_index += 1

for color in EMPTY_CARDS:
    for copy_index in range(1, EMPTY_CARDS[color] + 1):
        file_name = f"{print_index:03d}_{str(color)}.{copy_index}"
        if len(list(output_dir_fronts.glob(f"{file_name}*"))) == 0:
            card_printing.print_blank_card(
                color,
                output_dir_fronts,
                front_file_name=f"{file_name}_front",
                skip_back=True
            )
        if len(list(output_dir_backs.glob(f"{file_name}*"))) == 0:
            card_printing.print_blank_card(
                color,
                output_dir_backs,
                back_file_name=f"{file_name}_back",
                skip_front=True
            )
        print_index += 1
print("Done!")
