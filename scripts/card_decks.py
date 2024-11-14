from tabulate import tabulate

import src.cards as cards
import src.print_cards as print_cards

cards.import_all_data()


def deck_to_string(deck_input) -> str:
    deck_dict = {}  # Keys are card ids, values are number of copies
    all_printed_cards = print_cards.get_all_printable_cards()

    if isinstance(deck_input, dict):
        deck_dict = deck_input
    else:
        for card_id_or_number in deck_input:
            card_id = ""
            if isinstance(card_id_or_number, str):
                card_id = card_id_or_number
            elif isinstance(card_id_or_number, int):
                card_id = all_printed_cards[card_id_or_number - 1].get_id()

            if card_id == "":
                raise ValueError(f"Unknown card '{card_id_or_number}'")

            if card_id not in deck_dict:
                deck_dict[card_id] = 0
            deck_dict[card_id] += 1

    def key_func(card_id: str) -> int:
        card = cards.get_card(card_id)
        card_number = print_cards.get_card_printing_number(card)
        return card_number

    deck_row_list = []
    total_number_cards = 0
    for card_id in sorted(deck_dict.keys(), key=key_func):
        number_copies = deck_dict[card_id]
        card = cards.get_card(card_id)
        card_number = print_cards.get_card_printing_number(card)
        deck_row_list.append([
            f"{card_number:03d}",
            f"x{number_copies}",
            str(card.get_color()).upper(),
            card.get_name(),
            card.get_id()
        ])
        total_number_cards += number_copies

    output = tabulate(deck_row_list)
    if not total_number_cards == 50:
        output += f"\nWARNING: Deck has {total_number_cards}\\50 cards"
    return output


deck_fred_green_pink = [3, 3, 3, 4, 4, 4, 6, 6, 6, 28, 28, 28, 30, 30, 30, 33, 33, 33, 37, 37, 37, 41, 41, 41, 42, 42,
                        42, 43, 43, 43, 136, 136, 138, 138, 138, 141, 141, 141, 142, 142, 142, 143, 143, 143, 144, 144,
                        144, 145, 145, 145]

deck_henrique_first_yellow_self_destruct = [3, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 6, 9, 9, 9, 93, 93, 94, 94, 94, 95, 95,
                                            102, 102, 102, 107, 107, 107, 108, 108, 111, 111, 111, 158, 158, 160, 160,
                                            160, 161, 161, 161, 162, 162, 162, 163, 163, 163, 164, 164, 164]

deck_henrique_orange_cyan = [3, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 6, 8, 8, 8, 10, 10, 10, 18, 18, 25, 25, 25, 166, 166, 166,
                             168, 168, 168, 168, 174, 174, 174, 175, 175, 175, 178, 178, 178, 180, 180, 180, 181, 181,
                             181, 182, 182, 182, 186, 186]

deck_orange_purple = [3, 3, 3, 4, 4, 4, 5, 5, 6, 6, 6, 11, 11, 11, 16, 16, 16, 17, 17, 17, 19, 19, 19, 22, 22, 22, 24,
                      24, 25, 25, 112, 112, 112, 118, 118, 118, 120, 120, 120, 122, 122, 122, 125, 125, 125, 128, 128,
                      128, 130, 130]

deck_subtil_zoo_white_pink = [4, 4, 4, 5, 5, 6, 6, 6, 68, 68, 68, 70, 70, 70, 71, 71, 71, 73, 73, 73, 76, 76, 76, 79,
                              79, 81, 81, 82, 82, 86, 87, 87, 87, 89, 89, 90, 137, 137, 138, 138, 139, 139, 141, 141,
                              141, 143, 143, 143, 145, 145]

deck_stall_no_browns = [49, 49, 49, 52, 52, 52, 55, 55, 55, 56, 56, 56, 58, 58, 58, 63, 63, 63, 64, 64, 64, 65, 65, 65,
                        148, 148, 148, 149, 149, 150, 150, 150, 151, 151, 151, 153, 153, 153, 155, 155, 155, 161, 161,
                        161, 162, 162, 162, 163, 163, 163]

print(deck_to_string(deck_stall_no_browns))
print(sorted(deck_stall_no_browns))
