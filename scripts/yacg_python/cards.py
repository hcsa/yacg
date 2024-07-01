from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, Optional, Self, ClassVar, List

import yaml

from scripts.yacg_python.common_vars import CREATURE_DATA_PATH, TRAIT_DATA_PATH, EFFECT_DATA_PATH


class Color(Enum):
    """
    A card's color
    """

    NONE = (
        "None",
    )
    ORANGE = (
        "Orange",
    )
    GREEN = (
        "Green",
    )
    BLUE = (
        "Blue",
    )
    WHITE = (
        "White",
    )
    YELLOW = (
        "Yellow",
    )
    PURPLE = (
        "Purple",
    )
    PINK = (
        "Pink",
    )
    BLACK = (
        "Black",
    )
    CYAN = (
        "Cyan",
    )

    def __new__(cls, *args):
        if len(args) != 1:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 1 args: a human-readable name"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        return obj

    def __str__(self) -> str:
        return self._value_

    @property  # Ensures the attribute is read-only
    def name(self) -> str:
        return self._value_


class EffectType(Enum):
    """
    A type of effect
    """

    ACTION = (
        "Action",
        "Has an immediate effect. Goes to the discard pile after resolved"
    )
    FIELD = (
        "Field",
        "Has a continued effect. Stays on the board after resolved, until the end of the round"
    )
    AURA = (
        "Aura",
        "Has a continued effect on a creature. "
        "Stays attached to it after resolved, goes to the discard pile whenever the attached creature is no longer on "
        "the field"
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


class DevStage(Enum):
    """
    A development stage of a card
    """

    CONCEPTION = (
        "Conception",
        "Still in conception, not all fields may be filled in",
        100
    )
    ALPHA_0 = (
        "Alpha-0",
        "Ready to be used, never tried out",
        200
    )
    ALPHA_1 = (
        "Alpha-1",
        "Has been used at least once",
        201
    )
    DISCONTINUED = (
        "Discontinued",
        "Replaced by another card or abandoned entirely",
        9000
    )

    def __new__(cls, *args):
        if len(args) != 3:
            raise RuntimeError(
                f"A {cls.__name__} is ill-defined. "
                "It requires 3 args: a human-readable name, a description and an order number (used for sorting)"
            )
        obj = object.__new__(cls)
        obj._value_ = str(args[0])
        obj._description_ = str(args[1])
        obj._order_ = int(args[2])
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
    def order(self) -> int:
        return self._order_


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


class CardData(ABC):
    @abstractmethod
    def get_id(self) -> str:
        pass

    @abstractmethod
    def get_name(self) -> str:
        pass


class Card(CardData):
    @abstractmethod
    def get_color(self) -> Color:
        pass

    @abstractmethod
    def get_cost_total(self) -> int:
        pass

    @abstractmethod
    def get_cost_color(self) -> int:
        pass


@dataclass(frozen=True)
class TraitData:
    name: str
    description: str = ""


@dataclass(frozen=True)
class TraitMetadata:
    id: str
    type: TraitType
    value: Optional[int] = None
    dev_stage: DevStage = DevStage.CONCEPTION
    dev_name: str = ""
    order: Optional[int] = None
    summary: str = ""
    notes: str = ""


@dataclass(frozen=True)
class Trait(CardData):
    _id_prefix: ClassVar[str] = "T"
    _trait_dict: ClassVar[Dict[str, Self]] = {}

    data: TraitData
    metadata: TraitMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(self._id_prefix):
            raise ValueError(f"Trait's ID '{self.metadata.id}' doesn't start with prefix '{self._id_prefix}'")
        if self.metadata.id in self._trait_dict:
            raise ValueError(f"Trait with ID '{self.metadata.id}' already exists")
        self._trait_dict[self.metadata.id] = self

    def get_id(self) -> str:
        return self.metadata.id

    def get_name(self) -> str:
        if not self.data.name == "":
            return self.data.name
        elif not self.metadata.dev_name == "":
            return f"({self.metadata.dev_name})"
        return ""

    # Class method for _trait_dict
    # Implemented a class method, so it's read-only and is documented in a way IntelliSense can read it
    @classmethod
    def get_trait_dict(cls) -> Dict[str, Self]:
        """
        Returns a dict containing all traits, indexed by ID
        """

        return cls._trait_dict

    @classmethod
    def get_trait(cls, trait_id: str) -> Self:
        return cls._trait_dict[trait_id]

    @classmethod
    def import_from_yaml(cls, trait_id: str) -> Self:
        """
        Reads the trait data from the corresponding YAML file.
        Returns the imported trait.
        """

        yaml_path = TRAIT_DATA_PATH / f"{trait_id}.yaml"
        with open(yaml_path, "r") as f:
            yaml_data = yaml.safe_load(f)["trait"]

        trait_data = TraitData(
            name=(
                str(yaml_data["data"]["name"])
                if yaml_data["data"]["name"] is not None
                else ""
            ),
            description=str(yaml_data["data"]["description"]).strip()
        )
        trait_metadata = TraitMetadata(
            id=str(yaml_data["metadata"]["id"]),
            type=TraitType(str(yaml_data["metadata"]["type"])),
            value=(
                int(yaml_data["metadata"]["value"])
                if yaml_data["metadata"]["value"] is not None
                else None
            ),
            dev_stage=DevStage(str(yaml_data["metadata"]["dev-stage"])),
            dev_name=(
                str(yaml_data["metadata"]["dev-name"])
                if yaml_data["metadata"]["dev-name"] is not None
                else ""
            ),
            order=(
                int(yaml_data["metadata"]["order"])
                if yaml_data["metadata"]["order"] is not None
                else None
            ),
            summary=(
                str(yaml_data["metadata"]["summary"])
                if yaml_data["metadata"]["summary"] is not None
                else ""
            ),
            notes=str(yaml_data["metadata"]["notes"]).replace("\n      ", "\n").strip(),
        )
        trait = Trait(
            data=trait_data,
            metadata=trait_metadata,
        )
        return trait

    def export_to_yaml(self) -> None:
        """
        Writes the trait data to the corresponding YAML file.
        """

        notes_str = self.metadata.notes.strip().replace("\n", "\n      ")

        yaml_content = f"""
trait:
  data:
    name: {self.data.name}
    description: {self.data.description}

  metadata:
    id: {self.metadata.id}
    type: {self.metadata.type.name}
    value: {self.metadata.value if self.metadata.value is not None else ""}
    dev-stage: {self.metadata.dev_stage.name}
    dev-name: {self.metadata.dev_name}
    order: {self.metadata.order if self.metadata.order is not None else ""}
    summary: {self.metadata.summary}
    notes: |
      {notes_str}"""[1:]

        yaml_path = TRAIT_DATA_PATH / f"{self.metadata.id}.yaml"
        with open(yaml_path, "w") as f:
            f.write(yaml_content)

    @classmethod
    def import_all_from_yaml(cls) -> None:
        """
        Reads all traits' data from the YAML files.
        """

        for yaml_path in TRAIT_DATA_PATH.iterdir():
            trait_id = str(yaml_path.stem)
            _ = cls.import_from_yaml(trait_id)

    @classmethod
    def export_all_to_yaml(cls) -> None:
        """
        Writes all traits' data to YAML files.
        """

        for trait in cls._trait_dict.values():
            trait.export_to_yaml()


@dataclass(frozen=True)
class CreatureData:
    name: str
    color: Optional[Color]
    cost_total: Optional[int]
    cost_color: Optional[int]
    hp: Optional[int]
    atk: Optional[int]
    spe: Optional[int]
    traits: List[Trait] = field(default_factory=list)


@dataclass(frozen=True)
class CreatureMetadata:
    id: str
    value: Optional[int] = None
    dev_stage: DevStage = DevStage.CONCEPTION
    dev_name: str = ""
    order: Optional[int] = None
    summary: str = ""
    notes: str = ""


@dataclass(frozen=True)
class Creature(Card):
    _id_prefix: ClassVar[str] = "C"
    _creature_dict: ClassVar[Dict[str, Self]] = {}

    data: CreatureData
    metadata: CreatureMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(self._id_prefix):
            raise ValueError(f"Creature's ID '{self.metadata.id}' doesn't start with prefix '{self._id_prefix}'")
        if self.metadata.id in self._creature_dict:
            raise ValueError(f"Creature with ID '{self.metadata.id}' already exists")
        self._creature_dict[self.metadata.id] = self

    def get_id(self) -> str:
        return self.metadata.id

    def get_name(self) -> str:
        if not self.data.name == "":
            return self.data.name
        elif not self.metadata.dev_name == "":
            return f"({self.metadata.dev_name})"
        return ""

    def get_color(self) -> Color:
        return self.data.color

    def get_cost_total(self) -> int:
        return self.data.cost_total

    def get_cost_color(self) -> int:
        return self.data.cost_color

    # Class method for _creature_dict
    # Implemented a class method, so it's read-only and is documented in a way IntelliSense can read it
    @classmethod
    def get_creature_dict(cls) -> Dict[str, Self]:
        """
        Returns a dict containing all creatures, indexed by ID
        """

        return cls._creature_dict

    @classmethod
    def get_creature(cls, creature_id: str) -> Self:
        return cls._creature_dict[creature_id]

    @classmethod
    def import_from_yaml(cls, creature_id: str) -> Self:
        """
        Reads the creature data from the corresponding YAML file.
        Returns the imported creature.
        """

        yaml_path = CREATURE_DATA_PATH / f"{creature_id}.yaml"
        with open(yaml_path, "r") as f:
            yaml_data = yaml.safe_load(f)["creature"]

        traits_list = []
        if "traits" in yaml_data["data"]:
            trait_dict = Trait.get_trait_dict()
            for trait in yaml_data["data"]["traits"]:
                trait_id = str(trait["id"])
                if trait_id not in trait_dict:
                    raise ValueError(f"This creature has unknown trait '{trait_id}'")
                traits_list.append(trait_dict[trait_id])

        creature_data = CreatureData(
            name=(
                str(yaml_data["data"]["name"])
                if yaml_data["data"]["name"] is not None
                else ""
            ),
            color=(
                Color(str(yaml_data["data"]["color"]))
                if yaml_data["data"]["color"] is not None
                else None
            ),
            cost_total=(
                int(yaml_data["data"]["cost-total"])
                if yaml_data["data"]["cost-total"] is not None
                else None
            ),
            cost_color=(
                int(yaml_data["data"]["cost-color"])
                if yaml_data["data"]["cost-color"] is not None
                else None
            ),
            hp=(
                int(yaml_data["data"]["hp"])
                if yaml_data["data"]["hp"] is not None
                else None
            ),
            atk=(
                int(yaml_data["data"]["atk"])
                if yaml_data["data"]["atk"] is not None
                else None
            ),
            spe=(
                int(yaml_data["data"]["spe"])
                if yaml_data["data"]["spe"] is not None
                else None
            ),
            traits=traits_list,
        )
        creature_metadata = CreatureMetadata(
            id=str(yaml_data["metadata"]["id"]),
            value=(
                int(yaml_data["metadata"]["value"])
                if yaml_data["metadata"]["value"] is not None
                else None
            ),
            dev_stage=DevStage(str(yaml_data["metadata"]["dev-stage"])),
            dev_name=(
                str(yaml_data["metadata"]["dev-name"])
                if yaml_data["metadata"]["dev-name"] is not None
                else ""
            ),
            order=(
                int(yaml_data["metadata"]["order"])
                if yaml_data["metadata"]["order"] is not None
                else None
            ),
            summary=(
                str(yaml_data["metadata"]["summary"])
                if yaml_data["metadata"]["summary"] is not None
                else ""
            ),
            notes=str(yaml_data["metadata"]["notes"]).replace("\n      ", "\n").strip(),
        )
        creature = Creature(
            data=creature_data,
            metadata=creature_metadata,
        )
        return creature

    def export_to_yaml(self) -> None:
        """
        Writes the creature data to the corresponding YAML file.
        """

        notes_str = self.metadata.notes.strip().replace("\n", "\n      ")
        traits_str = ""
        if len(self.data.traits) > 0:
            traits_str += "    traits:\n"
            for trait in self.data.traits:
                trait_str = f"""
      - name: {trait.data.name}
        description: {trait.data.description}
        id: {trait.metadata.id}"""[1:]
                traits_str += trait_str + "\n"

        yaml_content = f"""
creature:
  data:
    name: {self.data.name}
    color: {self.data.color.name if self.data.color is not None else ""}
    cost-total: {self.data.cost_total if self.data.cost_total is not None else ""}
    cost-color: {self.data.cost_color if self.data.cost_color is not None else ""}
    hp: {self.data.hp if self.data.hp is not None else ""}
    atk: {self.data.atk if self.data.atk is not None else ""}
    spe: {self.data.spe if self.data.spe is not None else ""}
{traits_str}
  metadata:
    id: {self.metadata.id}
    value: {self.metadata.value if self.metadata.value is not None else ""}
    dev-stage: {self.metadata.dev_stage.name}
    dev-name: {self.metadata.dev_name}
    order: {self.metadata.order if self.metadata.order is not None else ""}
    summary: {self.metadata.summary}
    notes: |
      {notes_str}"""[1:]

        yaml_path = CREATURE_DATA_PATH / f"{self.metadata.id}.yaml"
        with open(yaml_path, "w") as f:
            f.write(yaml_content)

    @classmethod
    def import_all_from_yaml(cls) -> None:
        """
        Reads all creatures' data from the YAML files.
        """

        for yaml_path in CREATURE_DATA_PATH.iterdir():
            creature_id = str(yaml_path.stem)
            _ = cls.import_from_yaml(creature_id)

    @classmethod
    def export_all_to_yaml(cls) -> None:
        """
        Writes all creature's data to YAML files.
        """

        for creature in cls._creature_dict.values():
            creature.export_to_yaml()


@dataclass(frozen=True)
class EffectData:
    name: str
    color: Optional[Color]
    type: EffectType
    cost_total: Optional[int]
    cost_color: Optional[int]
    description: str = ""


@dataclass(frozen=True)
class EffectMetadata:
    id: str
    value: Optional[int] = None
    dev_stage: DevStage = DevStage.CONCEPTION
    dev_name: str = ""
    order: Optional[int] = None
    summary: str = ""
    notes: str = ""


@dataclass(frozen=True)
class Effect(Card):
    _id_prefix: ClassVar[str] = "E"
    _effect_dict: ClassVar[Dict[str, Self]] = {}

    data: EffectData
    metadata: EffectMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(self._id_prefix):
            raise ValueError(f"Effect's ID '{self.metadata.id}' doesn't start with prefix '{self._id_prefix}'")
        if self.metadata.id in self._effect_dict:
            raise ValueError(f"Effect with ID '{self.metadata.id}' already exists")
        self._effect_dict[self.metadata.id] = self

    def get_color(self) -> Color:
        return self.data.color

    def get_id(self) -> str:
        return self.metadata.id

    def get_name(self) -> str:
        if not self.data.name == "":
            return self.data.name
        elif not self.metadata.dev_name == "":
            return f"({self.metadata.dev_name})"
        return ""

    def get_cost_total(self) -> int:
        return self.data.cost_total

    def get_cost_color(self) -> int:
        return self.data.cost_color

    # Class method for _effect_dict
    # Implemented a class method, so it's read-only and is documented in a way IntelliSense can read it
    @classmethod
    def get_effect_dict(cls) -> Dict[str, Self]:
        """
        Returns a dict containing all effects, indexed by ID
        """

        return cls._effect_dict

    @classmethod
    def get_effect(cls, effect_id: str) -> Self:
        return cls._effect_dict[effect_id]

    @classmethod
    def import_from_yaml(cls, effect_id: str) -> Self:
        """
        Reads the effect data from the corresponding YAML file.
        Returns the imported effect.
        """

        yaml_path = EFFECT_DATA_PATH / f"{effect_id}.yaml"
        with open(yaml_path, "r") as f:
            yaml_data = yaml.safe_load(f)["effect"]

        effect_data = EffectData(
            name=(
                str(yaml_data["data"]["name"])
                if yaml_data["data"]["name"] is not None
                else ""
            ),
            color=(
                Color(str(yaml_data["data"]["color"]))
                if yaml_data["data"]["color"] is not None
                else None
            ),
            type=EffectType(str(yaml_data["data"]["type"])),
            cost_total=(
                int(yaml_data["data"]["cost-total"])
                if yaml_data["data"]["cost-total"] is not None
                else None
            ),
            cost_color=(
                int(yaml_data["data"]["cost-color"])
                if yaml_data["data"]["cost-color"] is not None
                else None
            ),
            description=str(yaml_data["data"]["description"]).strip()
        )
        effect_metadata = EffectMetadata(
            id=str(yaml_data["metadata"]["id"]),
            dev_stage=DevStage(str(yaml_data["metadata"]["dev-stage"])),
            dev_name=(
                str(yaml_data["metadata"]["dev-name"])
                if yaml_data["metadata"]["dev-name"] is not None
                else ""
            ),
            order=(
                int(yaml_data["metadata"]["order"])
                if yaml_data["metadata"]["order"] is not None
                else None
            ),
            summary=(
                str(yaml_data["metadata"]["summary"])
                if yaml_data["metadata"]["summary"] is not None
                else ""
            ),
            notes=str(yaml_data["metadata"]["notes"]).replace("\n      ", "\n").strip(),
        )
        effect = Effect(
            data=effect_data,
            metadata=effect_metadata,
        )
        return effect

    def export_to_yaml(self) -> None:
        """
        Writes the effect data to the corresponding YAML file.
        """

        description_str = self.data.description.strip().replace("\n", "\n      ")
        notes_str = self.metadata.notes.strip().replace("\n", "\n      ")

        yaml_content = f"""
effect:
  data:
    name: {self.data.name}
    color: {self.data.color.name if self.data.color is not None else ""}
    type: {self.data.type.name}
    cost-total: {self.data.cost_total if self.data.cost_total is not None else ""}
    cost-color: {self.data.cost_color if self.data.cost_color is not None else ""}
    description: |
      {description_str}

  metadata:
    id: {self.metadata.id}
    dev-stage: {self.metadata.dev_stage.name}
    dev-name: {self.metadata.dev_name}
    order: {self.metadata.order if self.metadata.order is not None else ""}
    summary: {self.metadata.summary}
    notes: |
      {notes_str}"""[1:]

        yaml_path = EFFECT_DATA_PATH / f"{self.metadata.id}.yaml"
        with open(yaml_path, "w") as f:
            f.write(yaml_content)

    @classmethod
    def import_all_from_yaml(cls) -> None:
        """
        Reads all effects' data from the YAML files.
        """

        for yaml_path in EFFECT_DATA_PATH.iterdir():
            effect_id = str(yaml_path.stem)
            _ = cls.import_from_yaml(effect_id)

    @classmethod
    def export_all_to_yaml(cls) -> None:
        """
        Writes all effects' data to YAML files.
        """

        for effect in cls._effect_dict.values():
            effect.export_to_yaml()


def import_all_data() -> None:
    """
    Imports all card data: Traits, Creatures and Effects.
    """

    # Traits must be imported before creatures
    Trait.import_all_from_yaml()
    Creature.import_all_from_yaml()
    Effect.import_all_from_yaml()


def export_all_data() -> None:
    """
    Exports all card data: Traits, Creatures and Effects.
    """

    Trait.export_all_to_yaml()
    Creature.export_all_to_yaml()
    Effect.export_all_to_yaml()


def get_card_data(card_data_id: str) -> CardData:
    if card_data_id.startswith(Trait._id_prefix):
        return Trait.get_trait(card_data_id)
    if card_data_id.startswith(Creature._id_prefix):
        return Creature.get_creature(card_data_id)
    if card_data_id.startswith(Effect._id_prefix):
        return Effect.get_effect(card_data_id)
    raise ValueError(f"Card data id '{card_data_id}' has an unexpected prefix")


def get_card(card_id: str) -> Card:
    if card_id.startswith(Creature._id_prefix):
        return Creature.get_creature(card_id)
    if card_id.startswith(Effect._id_prefix):
        return Effect.get_effect(card_id)
    raise ValueError(f"Card id '{card_id}' has an unexpected prefix")
