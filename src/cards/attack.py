from dataclasses import dataclass, field
from typing import Optional, Self, ClassVar, Dict, List

import yaml

from src.cards.abstract_classes import GameElement
from src.cards.enums import Color, DevStage, _GameElementIdPrefix
from src.utils import ATTACK_PATH, YAML_ENCODING


@dataclass(frozen=True)
class AttackData:
    name: str
    description: str = ""


@dataclass(frozen=True)
class AttackColors:
    primary: List[Color] = field(default_factory=list)
    secondary: List[Color] = field(default_factory=list)
    tertiary: List[Color] = field(default_factory=list)


@dataclass(frozen=True)
class AttackMetadata:
    id: str
    colors: AttackColors
    value: Optional[int] = None
    dev_stage: DevStage = DevStage.CONCEPTION
    dev_name: str = ""
    order: Optional[int] = None
    summary: str = ""
    notes: str = ""


@dataclass(frozen=True)
class Attack(GameElement):
    id_prefix: ClassVar[str] = _GameElementIdPrefix.ATTACK
    _attack_dict: ClassVar[Dict[str, Self]] = {}

    data: AttackData
    metadata: AttackMetadata

    def __post_init__(self):
        if not self.metadata.id.startswith(self.id_prefix):
            raise ValueError(
                f"Attack's ID '{self.metadata.id}' doesn't start with prefix '{self.id_prefix}'"
            )
        if self.metadata.id in self._attack_dict:
            raise ValueError(f"Attack with ID '{self.metadata.id}' already exists")
        self._attack_dict[self.metadata.id] = self

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

    # Class method for _attack_dict
    # Implemented a class method, so it's read-only and is documented in a way IntelliSense can read it
    @classmethod
    def get_attack_dict(cls) -> Dict[str, Self]:
        """
        Returns a dict containing all attacks, indexed by ID
        """

        return cls._attack_dict

    @classmethod
    def get_attack(cls, attack_id: str) -> Self:
        return cls._attack_dict[attack_id]

    @classmethod
    def import_from_yaml(cls, attack_id: str) -> Self:
        """
        Reads the attack data from the corresponding YAML file.
        Returns the imported attack.
        """

        yaml_path = ATTACK_PATH / f"{attack_id}.yaml"
        with open(yaml_path, "r", encoding=YAML_ENCODING) as f:
            yaml_data = yaml.safe_load(f)["attack"]

        primary_colors_list = []
        secondary_colors_list = []
        tertiary_colors_list = []
        if yaml_data["metadata"]["colors"] is not None:
            if "primary" in yaml_data["metadata"]["colors"]:
                for color_id in yaml_data["metadata"]["colors"]["primary"]:
                    color = Color(str(color_id))
                    primary_colors_list.append(color)
            if "secondary" in yaml_data["metadata"]["colors"]:
                for color_id in yaml_data["metadata"]["colors"]["secondary"]:
                    color = Color(str(color_id))
                    secondary_colors_list.append(color)
            if "tertiary" in yaml_data["metadata"]["colors"]:
                for color_id in yaml_data["metadata"]["colors"]["tertiary"]:
                    color = Color(str(color_id))
                    tertiary_colors_list.append(color)

        attack_data = AttackData(
            name=(
                str(yaml_data["data"]["name"])
                if yaml_data["data"]["name"] is not None
                else ""
            ),
            description=str(yaml_data["data"]["description"]).strip()
        )
        attack_colors = AttackColors(
            primary=primary_colors_list,
            secondary=secondary_colors_list,
            tertiary=tertiary_colors_list,
        )
        attack_metadata = AttackMetadata(
            id=str(yaml_data["metadata"]["id"]),
            colors=attack_colors,
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
            notes=(
                str(yaml_data["metadata"]["notes"]).strip()
                if "notes" in yaml_data["metadata"]
                else ""
            ),
        )
        attack = Attack(
            data=attack_data,
            metadata=attack_metadata,
        )
        return attack

    def export_to_yaml(self) -> None:
        """
        Writes the attack data to the corresponding YAML file.
        """

        notes_str = ""
        if not self.metadata.notes == "":
            notes_str += "    notes: |\n"
            notes_str += "      "
            notes_str += self.metadata.notes.strip().replace("\n", "\n      ")
            notes_str += "\n"

        colors_str = "    colors:\n"
        if len(self.metadata.colors.primary) > 0:
            colors_str += "      primary:\n"
            for color in self.metadata.colors.primary:
                colors_str += f"        - {color.name}\n"
        if len(self.metadata.colors.secondary) > 0:
            colors_str += "      secondary:\n"
            for color in self.metadata.colors.secondary:
                colors_str += f"        - {color.name}\n"
        if len(self.metadata.colors.tertiary) > 0:
            colors_str += "      tertiary:\n"
            for color in self.metadata.colors.tertiary:
                colors_str += f"        - {color.name}\n"

        yaml_content = "attack:\n"
        yaml_content += f"  data:\n"
        yaml_content += f"    name: {self.data.name}\n"
        yaml_content += f"    description: {self.data.description}\n"
        yaml_content += f"\n"
        yaml_content += f"  metadata:\n"
        yaml_content += f"    id: {self.metadata.id}\n"
        yaml_content += colors_str
        yaml_content += f"    value: {self.metadata.value if self.metadata.value is not None else ""}\n"
        yaml_content += f"    dev-stage: {self.metadata.dev_stage.name}\n"
        yaml_content += f"    dev-name: {self.metadata.dev_name}\n"
        yaml_content += f"    order: {self.metadata.order if self.metadata.order is not None else ""}\n"
        yaml_content += f"    summary: {self.metadata.summary}\n"
        yaml_content += notes_str

        yaml_path = ATTACK_PATH / f"{self.metadata.id}.yaml"
        with open(yaml_path, "w", encoding=YAML_ENCODING) as f:
            f.write(yaml_content)

    @classmethod
    def import_all_from_yaml(cls) -> None:
        """
        Reads all attacks' data from the YAML files.
        """

        for yaml_path in ATTACK_PATH.iterdir():
            attack_id = str(yaml_path.stem)
            _ = cls.import_from_yaml(attack_id)

    @classmethod
    def export_all_to_yaml(cls) -> None:
        """
        Writes all attacks' data to YAML files.
        """

        for attack in cls._attack_dict.values():
            attack.export_to_yaml()
