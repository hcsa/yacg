from dataclasses import dataclass, field
from typing import Optional, Self, ClassVar, Dict, List

import yaml

from src.cards.abstract_classes import Card
from src.cards.enums import Color, DevStage, _GameElementIdPrefix
from src.cards.trait import Trait
from src.cards.attack import Attack
from src.utils import CREATURE_DATA_PATH


@dataclass(frozen=True)
class CreatureData:
    name: str
    color: Optional[Color]
    is_token: bool
    cost_total: Optional[int]
    cost_color: Optional[int]
    hp: Optional[int]
    atk_strong: Optional[int]
    atk_strong_effect: Optional[Attack]
    atk_strong_effect_variable: Optional[int]
    atk_technical: Optional[int]
    atk_technical_effect: Optional[Attack]
    atk_technical_effect_variable: Optional[int]
    spe: Optional[int]
    traits: List[Trait] = field(default_factory=list)
    flavor_text: str = ""


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
    id_prefix: ClassVar[str] = _GameElementIdPrefix.CREATURE
    _creature_dict: ClassVar[Dict[str, Self]] = {}

    data: CreatureData
    metadata: CreatureMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(self.id_prefix):
            raise ValueError(
                f"Creature's ID '{self.metadata.id}' doesn't start with prefix '{self.id_prefix}'"
            )
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

    def get_dev_stage(self) -> DevStage:
        return self.metadata.dev_stage

    def get_color(self) -> Optional[Color]:
        return self.data.color

    def get_cost_total(self) -> Optional[int]:
        return self.data.cost_total

    def get_cost_color(self) -> Optional[int]:
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

        atk_strong_effect = None
        atk_strong_effect_variable = None
        if "atk-strong-effect" in yaml_data["data"]:
            atk_strong_effect_id = str(yaml_data["data"]["atk-strong-effect"]["id"])
            atk_strong_effect = Attack.get_attack(atk_strong_effect_id)
            if "variable" in yaml_data["data"]["atk-strong-effect"]:
                atk_strong_effect_variable = int(yaml_data["data"]["atk-strong-effect"]["variable"])

        atk_technical_effect = None
        atk_technical_effect_variable = None
        if "atk-technical-effect" in yaml_data["data"]:
            atk_technical_effect_id = str(yaml_data["data"]["atk-technical-effect"]["id"])
            atk_technical_effect = Attack.get_attack(atk_technical_effect_id)
            if "variable" in yaml_data["data"]["atk-technical-effect"]:
                atk_technical_effect_variable = int(yaml_data["data"]["atk-technical-effect"]["variable"])

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
            is_token=(
                bool(yaml_data["data"]["is-token"])
                if yaml_data["data"]["is-token"] is not None
                else False
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
            atk_strong=(
                int(yaml_data["data"]["atk-strong"])
                if yaml_data["data"]["atk-strong"] is not None
                else None
            ),
            atk_strong_effect=atk_strong_effect,
            atk_strong_effect_variable=atk_strong_effect_variable,
            atk_technical=(
                int(yaml_data["data"]["atk-technical"])
                if yaml_data["data"]["atk-technical"] is not None
                else None
            ),
            atk_technical_effect=atk_technical_effect,
            atk_technical_effect_variable=atk_technical_effect_variable,
            spe=(
                int(yaml_data["data"]["spe"])
                if yaml_data["data"]["spe"] is not None
                else None
            ),
            traits=traits_list,
            flavor_text=str(yaml_data["data"]["flavor-text"]).strip()
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

        flavor_text_str = ""
        if not self.data.flavor_text == "":
            flavor_text_str += "    flavor-text: |\n"
            flavor_text_str += "      "
            flavor_text_str += self.data.flavor_text.strip().replace("\n", "\n      ")

        notes_str = ""
        if not self.metadata.notes == "":
            notes_str += "    notes: |\n"
            notes_str += "      "
            notes_str += self.metadata.notes.strip().replace("\n", "\n      ")

        atk_strong_str = (
                "    atk-strong: "
                + (str(self.data.atk_strong) if self.data.atk_strong is not None else "")
                + "\n"
        )

        atk_technical_str = (
                "    atk-technical: "
                + (str(self.data.atk_technical) if self.data.atk_technical is not None else "")
                + "\n"
        )

        traits_str = ""
        if len(self.data.traits) > 0:
            traits_str += "    traits:\n"
            for trait in self.data.traits:
                traits_str += f"      - name: {trait.data.name}\n"
                traits_str += f"        description: {trait.data.description}\n"
                traits_str += f"        id: {trait.metadata.id}\n"

        atk_strong_effect_str = ""
        if self.data.atk_strong_effect is not None:
            atk_strong_effect = self.data.atk_strong_effect
            atk_strong_effect_str += f"    atk-strong-effect:\n"
            atk_strong_effect_str += f"      name: {atk_strong_effect.data.name}\n"
            if self.data.atk_strong_effect_variable is not None:
                atk_strong_effect_str += f"      variable: {self.data.atk_strong_effect_variable}\n"
            atk_strong_effect_str += f"      description: {atk_strong_effect.data.description}\n"
            atk_strong_effect_str += f"      id: {atk_strong_effect.metadata.id}\n"

        atk_technical_effect_str = ""
        if self.data.atk_technical_effect is not None:
            atk_technical_effect = self.data.atk_technical_effect
            atk_technical_effect_str += f"    atk-technical-effect:\n"
            atk_technical_effect_str += f"      name: {atk_technical_effect.data.name}\n"
            if self.data.atk_technical_effect_variable is not None:
                atk_technical_effect_str += f"      variable: {self.data.atk_technical_effect_variable}\n"
            atk_technical_effect_str += f"      description: {atk_technical_effect.data.description}\n"
            atk_technical_effect_str += f"      id: {atk_technical_effect.metadata.id}\n"

        yaml_content = "creature:\n"
        yaml_content += f"  data:\n"
        yaml_content += f"    name: {self.data.name}\n"
        yaml_content += f"    color: {self.data.color.name if self.data.color is not None else ""}\n"
        yaml_content += f"    is-token: {self.data.is_token}\n"
        yaml_content += f"    cost-total: {self.data.cost_total if self.data.cost_total is not None else ""}\n"
        yaml_content += f"    cost-color: {self.data.cost_color if self.data.cost_color is not None else ""}\n"
        yaml_content += f"    hp: {self.data.hp if self.data.hp is not None else ""}\n"
        yaml_content += f"    spe: {self.data.spe if self.data.spe is not None else ""}\n"
        yaml_content += traits_str
        yaml_content += atk_strong_str
        yaml_content += atk_strong_effect_str
        yaml_content += atk_technical_str
        yaml_content += atk_technical_effect_str
        yaml_content += flavor_text_str
        yaml_content += f"\n"
        yaml_content += f"  metadata:\n"
        yaml_content += f"    id: {self.metadata.id}\n"
        yaml_content += f"    value: {self.metadata.value if self.metadata.value is not None else ""}\n"
        yaml_content += f"    dev-stage: {self.metadata.dev_stage.name}\n"
        yaml_content += f"    dev-name: {self.metadata.dev_name}\n"
        yaml_content += f"    order: {self.metadata.order if self.metadata.order is not None else ""}\n"
        yaml_content += f"    summary: {self.metadata.summary}\n"
        yaml_content += notes_str

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
