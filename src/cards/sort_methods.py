from enum import Enum

from src.cards.abstract_classes import Card
from src.cards.creature import Creature
from src.cards.effect import Effect


def _sort_key_name(card: Card):
    return card.get_name()


def _sort_key_canonical(card: Card):
    """
    Sort by (from top to bottom):
    - Is playable
    - Dev stage (doesn't apply if card is playable)
    - Is a token creature
    - Color
    - Is creature or not (first creatures, then effects)
    - Cost total
    - Name
    """

    return (
        int(not card.is_playable()),
        (card.get_dev_stage().sort_key if not card.is_playable() else 0),
        int(isinstance(card, Creature) and card.data.is_token),
        card.get_color().sort_key if card.get_color() is not None else float('inf'),
        int(isinstance(card, Effect)),
        card.get_cost_total() if card.get_cost_total() is not None else float('inf'),
        card.get_name()
    )


class SortMethod(Enum):
    """
    Used to sort lists of cards
    """

    # Must be functions that take a card and return a key (used to order a list of cards)

    SORT_BY_NAME = _sort_key_name
    SORT_CANONICAL = _sort_key_canonical
