"""
This is an example script that showcases the "cards" module
"""

import src.cards as cards

cards.import_all_data()
c: cards.Card = cards.get_card("C012")
print(c.get_name())
