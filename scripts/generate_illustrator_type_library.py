"""
Generates type library to use with Adobe Illustrator

To be run from root, as:
python ./scripts/generate_illustrator_type_library.py

Requires some work beforehand, check README for more details
"""

import sys

from win32com.client import makepy


sys.argv = [
    "makepy",
    "-v",
    "-o",
    r"./scripts/yacg_python/illustrador_com.py",   # Location where type library will be generated
    "Illustrator.Application"   # Can be commented out, a window will pop up to choose the type library
]

makepy.main()
