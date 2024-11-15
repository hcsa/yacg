from typing import List, Callable, Any

from src.cards.abstract_classes import Card, Mechanic
from src.cards.creature import Creature
from src.cards.effect import Effect
from src.cards.enums import _MechanicIdPrefix
from src.cards.filter_methods import FilterMethod
from src.cards.sort_methods import SortMethod
from src.cards.trait import Trait


def import_all_data() -> None:
    # Creatures have traits, so traits must be imported first
    Trait.import_all_from_yaml()
    Creature.import_all_from_yaml()
    Effect.import_all_from_yaml()


def export_all_data() -> None:
    Trait.export_all_to_yaml()
    Creature.export_all_to_yaml()
    Effect.export_all_to_yaml()


def get_mechanic(mechanic_id: str) -> Mechanic:
    if mechanic_id.startswith(_MechanicIdPrefix.TRAIT):
        return Trait.get_trait(mechanic_id)
    if mechanic_id.startswith(_MechanicIdPrefix.CREATURE):
        return Creature.get_creature(mechanic_id)
    if mechanic_id.startswith(_MechanicIdPrefix.EFFECT):
        return Effect.get_effect(mechanic_id)
    raise ValueError(f"Mechanic id '{mechanic_id}' has an unexpected prefix")


def get_card(card_id: str) -> Card:
    if card_id.startswith(_MechanicIdPrefix.CREATURE):
        return Creature.get_creature(card_id)
    if card_id.startswith(_MechanicIdPrefix.EFFECT):
        return Effect.get_effect(card_id)
    raise ValueError(f"Card id '{card_id}' has an unexpected prefix")


def get_all_cards(
        filter_method: FilterMethod = FilterMethod.NO_FILTER,
        sort_method: SortMethod = SortMethod.SORT_CANONICAL
) -> List[Card]:
    filter_method: Callable[[Card], bool]
    sort_method: Callable[[Card], Any]

    output = (
            [c for c in Creature.get_creature_dict().values() if filter_method(c)]
            + [c for c in Effect.get_effect_dict().values() if filter_method(c)]
    )
    output = sorted(output, key=sort_method)
    return output
