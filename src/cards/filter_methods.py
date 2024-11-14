from enum import Enum

from src.cards.abstract_classes import Card


def _no_filter(card: Card):
    return True


def _is_playable(card: Card):
    return card.is_playable()


class FilterMethod(Enum):
    """
    Used to filter out cards from lists
    """

    # Must be functions that take a card and return a bool (True to keep a card, False otherwise)

    NO_FILTER = _no_filter
    IS_PLAYABLE = _is_playable
