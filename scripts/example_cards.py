"""
This is an example script that showcases the "cards" module
"""

import src.cards as cards

cards.import_all_data()
c: cards.Card = cards.get_card("CRE-012")
print(c.get_name())
cards.export_all_data()
