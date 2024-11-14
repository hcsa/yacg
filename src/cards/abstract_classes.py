from abc import ABC, abstractmethod
from typing import Optional

from src.cards.enums import Color, DevStage


class Mechanic(ABC):
    """
    Implements a game mechanic relevant enough to be sorted and version-controlled
    """

    @abstractmethod
    def get_id(self) -> str:
        pass

    @abstractmethod
    def get_name(self) -> str:
        pass

    @abstractmethod
    def get_dev_stage(self) -> DevStage:
        pass


class Card(Mechanic):
    """
    Implements a playable card
    """

    @abstractmethod
    def get_id(self) -> str:
        pass

    @abstractmethod
    def get_name(self) -> str:
        pass

    @abstractmethod
    def get_dev_stage(self) -> DevStage:
        pass

    @abstractmethod
    def get_color(self) -> Optional[Color]:
        pass

    @abstractmethod
    def get_cost_total(self) -> Optional[int]:
        pass

    @abstractmethod
    def get_cost_color(self) -> Optional[int]:
        pass

    def is_playable(self) -> bool:
        """
        Returns True if card is ready to be used, False otherwise (eg, card is discontinued)
        """

        return self.get_dev_stage().is_playable
