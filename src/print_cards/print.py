import os
import shutil
from pathlib import Path
from typing import Optional

import src.cards as cards
import src.print_cards.generate_back as generate_back
import src.print_cards.generate_blank_front as generate_blank_front
import src.print_cards.generate_front as generate_front
import src.print_cards.helpers as helpers
import src.print_cards.illustrator_com as illustrator_com
from src.utils import CARD_TEMPLATE_PATH


def print_card(
        card: cards.Card,
        output_dir: Path,
        skip_front: bool = False,
        skip_back: bool = False,
        front_file_name: Optional[str] = None,
        back_file_name: Optional[str] = None,
) -> None:
    app = helpers.get_illustrator_app()

    if not skip_front:
        file_name = front_file_name
        if file_name is None:
            file_name = f"{card.get_id()}_front"  # Extension is added on export

        _print_card_front(card, app, output_dir, file_name)

    if not skip_back:
        file_name = back_file_name
        if file_name is None:
            file_name = f"{card.get_id()}_front"  # Extension is added on export

        _print_card_back(card.get_color(), app, output_dir, file_name)


def _print_card_front(
        card: cards.Card,
        app: illustrator_com.Application,
        output_dir: Path,
        file_name: str,
) -> None:
    output_path = output_dir / file_name
    temp_file_path = output_dir / f"{file_name}.temp"

    shutil.copy2(CARD_TEMPLATE_PATH, temp_file_path)
    temp_document = app.Open(temp_file_path)
    generate_front.generate_front(card, temp_document)
    helpers.export_to_tiff(temp_document, output_path)
    temp_document.Close(illustrator_com.constants.aiDoNotSaveChanges)
    os.remove(temp_file_path)


def _print_card_back(
        color: cards.Color,
        app: illustrator_com.Application,
        output_dir: Path,
        file_name: str,
) -> None:
    output_path = output_dir / file_name
    temp_file_path = output_dir / f"{file_name}.temp"

    shutil.copy2(CARD_TEMPLATE_PATH, temp_file_path)
    temp_document = app.Open(temp_file_path)
    generate_back.generate_back(color, temp_document)
    helpers.export_to_tiff(temp_document, output_path)
    temp_document.Close(illustrator_com.constants.aiDoNotSaveChanges)
    os.remove(temp_file_path)


def print_blank_card(
        color: cards.Color,
        output_dir: Path,
        skip_front: bool = False,
        skip_back: bool = False,
        front_file_name: Optional[str] = None,
        back_file_name: Optional[str] = None,
) -> None:
    app = helpers.get_illustrator_app()

    if not skip_front:
        file_name = front_file_name
        if file_name is None:
            file_name = f"{str(color)}_front"  # Extension is added on export

        _print_blank_card_front(color, app, output_dir, file_name)

    if not skip_back:
        file_name = back_file_name
        if file_name is None:
            file_name = f"{str(color)}_front"  # Extension is added on export

        _print_card_back(color, app, output_dir, file_name)


def _print_blank_card_front(
        color: cards.Color,
        app: illustrator_com.Application,
        output_dir: Path,
        file_name: str,
) -> None:
    output_path = output_dir / file_name
    temp_file_path = output_dir / f"{file_name}.temp"

    shutil.copy2(CARD_TEMPLATE_PATH, temp_file_path)
    temp_document = app.Open(temp_file_path)
    generate_blank_front.generate_blank_front(color, temp_document)
    helpers.export_to_tiff(temp_document, output_path)
    temp_document.Close(illustrator_com.constants.aiDoNotSaveChanges)
    os.remove(temp_file_path)
