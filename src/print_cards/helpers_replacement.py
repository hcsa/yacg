import re
from typing import List, Tuple

import src.cards as cards
import src.print_cards.errors as errors
import src.print_cards.illustrator_com as illustrator_com
from src.print_cards.configs import KEYWORD_TO_CHARACTER_DICT


def replace_placeholders(text_frame: illustrator_com.TextFrame) -> Tuple[List[int], List[int]]:
    """
    Replaces all keywords in the text frame with the respective icons, and references with content

    Returns 2 lists:
      - Indices of the character icons (1-indexed)
      - Indices of the characters belonging to references' names (1-indexed)
    """

    # Replace descriptions first, because the replacement text may have keywords or names
    _replace_descriptions(text_frame)
    list_indices_keyword_characters, list_indices_name_characters = _replace_keywords_and_names(text_frame)
    return list_indices_keyword_characters, list_indices_name_characters


def _replace_descriptions(text_frame: illustrator_com.TextFrame):
    text_before = text_frame.Contents

    pattern = r"\(REF\:(?P<id>[^\)]+?)\.DESCRIPTION\)"
    matches = re.finditer(pattern, text_before)

    text_after = ""
    index_last_match_end = 0
    for match in matches:
        match_replacement = _get_replacement_text_for_description_match(match)

        text_after += text_before[index_last_match_end:match.start()] + match_replacement
        index_last_match_end = match.end()

    text_after += text_before[index_last_match_end:]

    text_frame.Contents = text_after


def _get_replacement_text_for_description_match(match: re.Match) -> str:
    trait_id = str(match["id"])
    trait = cards.get_mechanic(trait_id)
    if not isinstance(trait, cards.Trait):
        raise errors.CardPrintError(f"Found reference '{match.group(0)}', but '{trait_id}' isn't a trait ID")

    replacement = trait.data.description
    if replacement[-1] == ".":  # Remove ending full stop
        replacement = replacement[:-1]

    return replacement


def _replace_keywords_and_names(text_frame: illustrator_com.TextFrame) -> Tuple[List[int], List[int]]:
    """
    Returns 2 lists:
      - Indices of the character icons (1-indexed)
      - Indices of the characters belonging to references' names (1-indexed)
    """

    text_before = text_frame.Contents

    pattern_keywords = "|".join(
        # Escape the values because they have special characters
        re.escape(keyword) for keyword in KEYWORD_TO_CHARACTER_DICT
    )
    pattern_names = r"\(REF\:(?P<id>[^\)]+?)\.NAME\)"
    pattern = rf"({pattern_keywords})|({pattern_names})"
    matches = re.finditer(pattern, text_before)

    text_after = ""
    index_last_match_end = 0
    list_indices_keyword_characters: List[int] = []
    list_indices_name_characters: List[int] = []
    for match in matches:
        if match.group(0) in KEYWORD_TO_CHARACTER_DICT:
            match_replacement = KEYWORD_TO_CHARACTER_DICT[match.group(0)]

            list_indices_keyword_characters += range(len(text_after) + 1, len(text_after) + 1 + len(match_replacement))
        else:
            match_replacement = _get_replacement_text_for_name_match(match)

            list_indices_name_characters += range(len(text_after) + 1, len(text_after) + 1 + len(match_replacement))

        text_after += text_before[index_last_match_end:match.start()] + match_replacement
        index_last_match_end = match.end()

    text_after += text_before[index_last_match_end:]

    text_frame.Contents = text_after
    return list_indices_keyword_characters, list_indices_name_characters


def _get_replacement_text_for_name_match(match: re.match) -> str:
    mechanic_id = str(match["id"])
    mechanic = cards.get_mechanic(mechanic_id)

    replacement = mechanic.get_name()
    if replacement == "" and isinstance(mechanic, cards.Card):
        replacement = f"(Card {mechanic_id})"
    if replacement == "" and isinstance(mechanic, cards.Trait):
        replacement = f"(Trait {mechanic_id})"

    return replacement
