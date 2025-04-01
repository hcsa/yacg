from typing import List, Callable, Any

from src.cards.abstract_classes import Card, GameElement
from src.cards.attack import Attack
from src.cards.creature import Creature
from src.cards.effect import Effect
from src.cards.enums import _GameElementIdPrefix
from src.cards.filter_methods import FilterMethod
from src.cards.mechanic import Mechanic
from src.cards.sort_methods import SortMethod
from src.cards.trait import Trait


def import_all_data() -> None:
    # Creatures have traits and attacks, so these must be imported first
    Mechanic.import_all_from_yaml()
    Trait.import_all_from_yaml()
    Attack.import_all_from_yaml()
    Creature.import_all_from_yaml()
    Effect.import_all_from_yaml()


def export_all_data() -> None:
    Mechanic.export_all_to_yaml()
    Trait.export_all_to_yaml()
    Attack.export_all_to_yaml()
    Creature.export_all_to_yaml()
    Effect.export_all_to_yaml()


def get_game_element(game_element_id: str) -> GameElement:
    if game_element_id.startswith(_GameElementIdPrefix.MECHANIC):
        return Mechanic.get_mechanic(game_element_id)
    if game_element_id.startswith(_GameElementIdPrefix.TRAIT):
        return Trait.get_trait(game_element_id)
    if game_element_id.startswith(_GameElementIdPrefix.ATTACK):
        return Attack.get_attack(game_element_id)
    if game_element_id.startswith(_GameElementIdPrefix.CREATURE):
        return Creature.get_creature(game_element_id)
    if game_element_id.startswith(_GameElementIdPrefix.EFFECT):
        return Effect.get_effect(game_element_id)
    raise ValueError(f"Game element id '{game_element_id}' has an unexpected prefix")


def get_card(card_id: str) -> Card:
    if card_id.startswith(_GameElementIdPrefix.CREATURE):
        return Creature.get_creature(card_id)
    if card_id.startswith(_GameElementIdPrefix.EFFECT):
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
