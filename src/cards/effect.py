from dataclasses import dataclass
from typing import Optional, Self, ClassVar, Dict

import yaml

from src.cards.abstract_classes import Card
from src.cards.enums import Color, DevStage, EffectType, _MechanicIdPrefix
from src.utils import EFFECT_DATA_PATH


@dataclass(frozen=True)
class EffectData:
    name: str
    color: Optional[Color]
    type: EffectType
    cost_total: Optional[int]
    cost_color: Optional[int]
    description: str = ""
    flavor_text: str = ""


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
    _effect_dict: ClassVar[Dict[str, Self]] = {}

    data: EffectData
    metadata: EffectMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(_MechanicIdPrefix.EFFECT):
            raise ValueError(f"Effect's ID '{self.metadata.id}' doesn't start with prefix '{_MechanicIdPrefix.EFFECT}'")
        if self.metadata.id in self._effect_dict:
            raise ValueError(f"Effect with ID '{self.metadata.id}' already exists")
        self._effect_dict[self.metadata.id] = self

    def get_color(self) -> Optional[Color]:
        return self.data.color

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

    def get_cost_total(self) -> Optional[int]:
        return self.data.cost_total

    def get_cost_color(self) -> Optional[int]:
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
            description=str(yaml_data["data"]["description"]).strip(),
            flavor_text=str(yaml_data["data"]["flavor-text"]).strip()
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
        flavor_text_str = self.data.flavor_text.strip().replace("\n", "\n      ")
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
    flavor-text: |
      {flavor_text_str}

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
