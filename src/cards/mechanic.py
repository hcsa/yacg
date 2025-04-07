from dataclasses import dataclass, field
from typing import Optional, Self, ClassVar, Dict, List

import yaml

from src.cards.abstract_classes import GameElement
from src.cards.enums import Color, DevStage, _GameElementIdPrefix
from src.utils import MECHANIC_DATA_PATH, YAML_ENCODING


@dataclass(frozen=True)
class MechanicColors:
    primary: List[Color] = field(default_factory=list)
    secondary: List[Color] = field(default_factory=list)
    tertiary: List[Color] = field(default_factory=list)


@dataclass(frozen=True)
class Mechanic(GameElement):
    id_prefix: ClassVar[str] = _GameElementIdPrefix.MECHANIC
    _mechanic_dict: ClassVar[Dict[str, Self]] = {}

    name: str
    colors: MechanicColors
    id: str
    dev_stage: DevStage = DevStage.CONCEPTION
    order: Optional[int] = None
    notes: str = ""

    def __post_init__(self):
        if not self.id.startswith(self.id_prefix):
            raise ValueError(f"Mechanic's ID '{self.id}' doesn't start with prefix '{self.id_prefix}'")
        if self.id in self._mechanic_dict:
            raise ValueError(f"Mechanic with ID '{self.id}' already exists")
        self._mechanic_dict[self.id] = self

    def get_id(self) -> str:
        return self.id

    def get_name(self) -> str:
        return self.name

    def get_dev_stage(self) -> DevStage:
        return self.dev_stage

    # Class method for _mechanic_dict
    # Implemented a class method, so it's read-only and is documented in a way IntelliSense can read it
    @classmethod
    def get_mechanic_dict(cls) -> Dict[str, Self]:
        """
        Returns a dict containing all mechanics, indexed by ID
        """

        return cls._mechanic_dict

    @classmethod
    def get_mechanic(cls, mechanic_id: str) -> Self:
        return cls.get_mechanic_dict()[mechanic_id]

    @classmethod
    def import_from_yaml(cls, mechanic_id: str) -> Self:
        """
        Reads the mechanic data from the corresponding YAML file.
        Returns the imported mechanic.
        """

        yaml_path = MECHANIC_DATA_PATH / f"{mechanic_id}.yaml"
        with open(yaml_path, "r", encoding=YAML_ENCODING) as f:
            yaml_data = yaml.safe_load(f)["mechanic"]

        primary_colors_list = []
        secondary_colors_list = []
        tertiary_colors_list = []
        if yaml_data["colors"] is not None:
            if "primary" in yaml_data["colors"]:
                for color_id in yaml_data["colors"]["primary"]:
                    color = Color(str(color_id))
                    primary_colors_list.append(color)
            if "secondary" in yaml_data["colors"]:
                for color_id in yaml_data["colors"]["secondary"]:
                    color = Color(str(color_id))
                    secondary_colors_list.append(color)
            if "tertiary" in yaml_data["colors"]:
                for color_id in yaml_data["colors"]["tertiary"]:
                    color = Color(str(color_id))
                    tertiary_colors_list.append(color)

        mechanic_colors = MechanicColors(
            primary=primary_colors_list,
            secondary=secondary_colors_list,
            tertiary=tertiary_colors_list,
        )
        mechanic = Mechanic(
            name=str(yaml_data["name"]).strip(),
            colors=mechanic_colors,
            id=str(yaml_data["id"]),
            dev_stage=DevStage(str(yaml_data["dev-stage"])),
            order=(
                int(yaml_data["order"])
                if yaml_data["order"] is not None
                else None
            ),
            notes=(
                str(yaml_data["notes"]).strip()
                if "notes" in yaml_data
                else ""
            ),
        )
        return mechanic

    def export_to_yaml(self) -> None:
        """
        Writes the mechanic data to the corresponding YAML file.
        """

        name_str = "  name: |\n"   # Has a newline because mechanics' names may have quotes
        name_str += f"    {self.name}\n"

        colors_str = "  colors:\n"
        if len(self.colors.primary) > 0:
            colors_str += "    primary:\n"
            for color in self.colors.primary:
                colors_str += f"      - {color.name}\n"
        if len(self.colors.secondary) > 0:
            colors_str += "    secondary:\n"
            for color in self.colors.secondary:
                colors_str += f"      - {color.name}\n"
        if len(self.colors.tertiary) > 0:
            colors_str += "    tertiary:\n"
            for color in self.colors.tertiary:
                colors_str += f"      - {color.name}\n"

        notes_str = ""
        if not self.notes == "":
            notes_str += "  notes: |\n"
            notes_str += "    "
            notes_str += self.notes.strip().replace("\n", "\n      ")
            notes_str += "\n"

        yaml_content = "mechanic:\n"
        yaml_content += name_str
        yaml_content += f"  id: {self.id}\n"
        yaml_content += colors_str
        yaml_content += f"  dev-stage: {self.dev_stage.name}\n"
        yaml_content += f"  order: {self.order if self.order is not None else ""}\n"
        yaml_content += notes_str

        yaml_path = MECHANIC_DATA_PATH / f"{self.id}.yaml"
        with open(yaml_path, "w", encoding=YAML_ENCODING) as f:
            f.write(yaml_content)

    @classmethod
    def import_all_from_yaml(cls) -> None:
        """
        Reads all mechanics' data from the YAML files.
        """

        for yaml_path in MECHANIC_DATA_PATH.iterdir():
            mechanic_id = str(yaml_path.stem)
            _ = cls.import_from_yaml(mechanic_id)

    @classmethod
    def export_all_to_yaml(cls) -> None:
        """
        Writes all mechanics' data to YAML files.
        """

        for mechanic in cls._mechanic_dict.values():
            mechanic.export_to_yaml()
