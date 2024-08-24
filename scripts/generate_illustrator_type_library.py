"""
Run this script to generate the type library for Adobe Illustrator

Requires some work beforehand, check README for more details
"""

import sys

from win32com.client import makepy

from scripts.yacg_python.common_vars import BASE_DIR

sys.argv = [
    "makepy",
    "-v",
    "-o",
    str(BASE_DIR / "scripts" / "yacg_python" / "illustrator_com.py"),  # Location where type library will be generated
    "Illustrator.Application"  # Can be commented out, a window will pop up to choose the type library
]

makepy.main()
