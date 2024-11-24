from dataclasses import dataclass
from typing import Optional, Self, ClassVar, Dict

import yaml

from src.cards.abstract_classes import GameElement
from src.cards.enums import DevStage, TraitType, _MechanicIdPrefix
from src.utils import TRAIT_DATA_PATH


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
class Trait(GameElement):
    _trait_dict: ClassVar[Dict[str, Self]] = {}

    data: TraitData
    metadata: TraitMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(_MechanicIdPrefix.TRAIT):
            raise ValueError(f"Trait's ID '{self.metadata.id}' doesn't start with prefix '{_MechanicIdPrefix.TRAIT}'")
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

    def get_dev_stage(self) -> DevStage:
        return self.metadata.dev_stage

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
