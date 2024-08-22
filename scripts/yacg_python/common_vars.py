import os
import sys
from pathlib import Path
from typing import Optional

from git import Repo


def _get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        # We're running from .exe
        # .exe is located in the root
        base_dir = Path(os.path.dirname(sys.executable))
    else:
        try:
            base_dir = Path(os.path.realpath(__file__)).parent.parent.parent
            # If this doesn't return NameError, then we're running in non-interactive mode
            # This file is located in [BASE_DIR] / scripts / yacg_python / common_vars.py
        except NameError:
            # We're running in interactive mode
            # This file is located in [BASE_DIR] / scripts / yacg_python / common_vars.py
            base_dir = Path(os.getcwd()).parent

    return base_dir


BASE_DIR = _get_base_dir()

EXCEL_PATH = BASE_DIR / "Cards.xlsx"
EXCEL_BACKUP_PATH = BASE_DIR / "Cards (backup).xlsx"
EXCEL_TEMPLATE_PATH = BASE_DIR / "excel_base.xlsx"

CARD_TEMPLATE_PATH = BASE_DIR / "card_design" / "card_template.ai"
CARD_ARTS_DIR = BASE_DIR / "card_design" / "card_arts"

CARD_DATA_PATH = BASE_DIR / "card_data"
CREATURE_DATA_PATH = CARD_DATA_PATH / "creatures"
EFFECT_DATA_PATH = CARD_DATA_PATH / "effects"
TRAIT_DATA_PATH = CARD_DATA_PATH / "traits"

VALUES_DATA_PATH = BASE_DIR / "dev_data" / "values.yaml"


def _get_git_tag() -> Optional[str]:
    repo = Repo.init(BASE_DIR)
    commit = repo.head.commit
    for tag in repo.tags:
        if tag.commit == commit:
            return tag.name
    return None


GIT_TAG_NAME = _get_git_tag()
