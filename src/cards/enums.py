from enum import Enum, StrEnum


class _GameElementIdPrefix(StrEnum):
    """
    A prefix, added before all ids of game elements/card types

    Differs for different types of game elements/card types

    If you change this, you'll have to change all ids of game elements
    """

    MECHANIC = "M"
    TRAIT = "T"
    EFFECT = "E"
    CREATURE = "C"


class Color(Enum):
    """
    A card's color
    """

    NONE = (
        "None",
        0
    )
    ORANGE = (
        "Orange",
        1
    )
    GREEN = (
        "Green",
        2
    )
    BLUE = (
        "Blue",
        3
    )
    WHITE = (
        "White",
        4
    )
    YELLOW = (
        "Yellow",
        5
    )
    PURPLE = (
        "Purple",
        6
    )
    PINK = (
        "Pink",
        7
    )
    BLACK = (
        "Black",
        8
    )
    CYAN = (
        "Cyan",
        9
    )

    def __new__(cls, *args):
        if len(args) != 2:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 2 args: a human-readable name and a value used for sorting"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        obj._sort_key_ = int(args[1])
        return obj

    def __str__(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def name(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def sort_key(self) -> int:
        return self._sort_key_


class EffectType(Enum):
    """
    A type of effect
    """

    ACTION = (
        "Action",
        "Has an immediate effect. Goes to the discard pile after resolved",
        0
    )
    AURA = (
        "Aura",
        "Has a continued effect on a creature. "
        "Stays attached to it after resolved, goes to the discard pile whenever the attached creature is no longer on "
        "the field",
        1
    )
    FIELD = (
        "Field",
        "Has a continued effect. Stays on the board after resolved, until the end of the round",
        2
    )

    def __new__(cls, *args):
        if len(args) != 3:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 3 args: a human-readable name, a description and a value used for sorting"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        obj._description_ = str(args[1])
        obj._sort_key_ = int(args[2])
        return obj

    def __str__(self) -> str:
        return f"{self._value_} ({self._description_})"

    @property  # Ensures the attribute is read-only
    def name(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def description(self) -> str:
        return self._description_

    @property  # Ensures the attribute is read-only
    def sort_key(self) -> int:
        return self._sort_key_


class DevStage(Enum):
    """
    A development stage of a card
    """

    CONCEPTION = (
        "Conception",
        "Still in conception, not all fields may be filled in",
        False,
        1000
    )
    ALPHA_0 = (
        "Alpha-0",
        "Ready to be used, never tried out",
        True,
        100
    )
    ALPHA_1 = (
        "Alpha-1",
        "Has been used at least once",
        True,
        101
    )
    DISCONTINUED = (
        "Discontinued",
        "Replaced by another card or abandoned entirely",
        False,
        9000
    )

    def __new__(cls, *args):
        if len(args) != 4:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 4 args: a human-readable name, a description, whether it's playable and a value used for "
                "sorting"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        obj._description_ = str(args[1])
        obj._is_playable_ = bool(args[2])
        obj._sort_key_ = int(args[3])
        return obj

    def __str__(self) -> str:
        return f"{self._value_} ({self._description_})"

    @property  # Ensures the attribute is read-only
    def name(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def description(self) -> str:
        return self._description_

    @property  # Ensures the attribute is read-only
    def is_playable(self) -> bool:
        return self._is_playable_

    @property  # Ensures the attribute is read-only
    def sort_key(self) -> int:
        return self._sort_key_


class TraitType(Enum):
    """
    A type of trait
    """

    ENTRY = (
        "Entry",
        "Its effect happens when the creature enters the field"
    )
    COMBAT = (
        "Combat",
        "Its effect happens when the creature is in combat"
    )
    OTHER = (
        "Other",
        "Doesn't fit any of the above"
    )

    def __new__(cls, *args):
        if len(args) != 2:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 2 args: a human-readable name and a description"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        obj._description_ = str(args[1])
        return obj

    def __str__(self) -> str:
        return f"{self._value_} ({self._description_})"

    @property  # Ensures the attribute is read-only
    def name(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def description(self) -> str:
        return self._description_
