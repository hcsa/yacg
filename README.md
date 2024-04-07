# YACG

## Development

Requires
- Python 3.12
- Adobe Illustrator 2022 (26.1)

### Compile executables
To create the .exe files, install pyinstaller (`pip install pyinstaller`) and run

```bash
pyinstaller scripts/import_to_excel.py -F -c --clean --specpath ./_pyinstaller_cache/ --distpath ./ --workpath ./_pyinstaller_cache/

pyinstaller scripts/export_from_excel.py -F -c --clean --specpath ./_pyinstaller_cache/ --distpath ./ --workpath ./_pyinstaller_cache/
```

### Modify Adobe Illustrator scripts

These scripts are written in ExtendScript, which is an extension of JavaScript used by Adobe softwares.

The best way to work on these is through VSCode, with the extension "ExtendScript Development Pack". You can run "Types for Adobe: Set-Up Types for Adobe" so VSCode is aware of pre-configured objects that ExtendScript uses.

## Cards YAML structure

### Creature

```yaml
creature:
  # All fields that have gameplay influence
  data:
    name: The Fred
    color: Yellow
    cost-total: 4
    cost-color: 2
    hp: 3
    atk: 1
    spd: 2
    traits: # Max 4
      # Traits are matched by ID, name and description are redundant.
      # They're here to be human-friendly.
      - name: Subtle
        description: Can be cast at any point while the opponent isn't looking
        id: TXXX

  # All fields that have no gameplay influence
  metadata:
    id: CXXX
    # This is a heuristic of how good a card is.
    # Positive value indicates it's too good, negative values indicates it's too bad.
    # It's derived from stats and traits, look into Excel for full computation.
    # Don't try too hard to make this 0, this is just a heuristic.
    value: 10
    # Check Excel for what these mean.
    dev-stage: Discontinued
    # While card has no serious name.
    dev-name: Template creature
    # Used to order cards in Excel.
    order: 0
    # The creature's main idea.
    # Eg: "2-cost red card", "black card with arrogant and defeatist traits, seems funny".
    summary: Showcase creature YAML
    # Notes during card development.
    # Fill this with notes on usage, balancing, etc.
    # Eg: "any cost less than 3 makes this busted", "rejected due to having no counter-play", "value is -20 but that's fine, Haste + Moxie makes up for it").
    notes: |
      * This is a template creature card. Exists purely as a template. Will never be printed. Isn't that so sad?
      * Grouped the fields in "data" and "metadata" groups.
        This way it's obvious what fields have gameplay influence and what fields don't.
```

### Effect

```yaml
effect:
  # All fields that have gameplay influence
  data:
    name: Henrique's Idea
    color: Orange
    # Either "Action", "Field" or "Aura"
    type: Field
    cost-total: 7
    cost-color: 3
    description: |
      Propose a new mechanic. That mechanic is valid for this game.

  # All fields that have no gameplay influence
  metadata:
    id: EXXX
    # Check Excel for what these mean.
    dev-stage: Discontinued
    # While card has no serious name.
    dev-name: Template effect
    # Used to order cards in Excel.
    order: 0
    # The creature's main idea.
    # Eg: "cheap +Atk aura", "black action that gives energy".
    summary: Showcase effect YAML
    # Notes during card development.
    # Fill this with notes on usage, balancing, etc.
    # Eg: "any cost less than 3 makes this busted", "rejected due to having no counter-play", "changed colors, fits blue more").
    notes: |
      * This is a template effect card. Exists purely as a template. Will never be printed. Isn't that so sad?
      * Grouped the fields in "data" and "metadata" groups.
        This way it's obvious what fields have gameplay influence and what fields don't.
```

### Trait

```yaml
trait:
  # All fields that have gameplay influence
  data:
    name: Subtle
    description: |
      Can be cast at any point while the opponent isn't looking

  # All fields that have no gameplay influence
  metadata:
    id: TXXX
    # Either "Cast" (has effect when creature's cast), "Combat" (has effect when creature is in battle) or "Other".
    type: Other
    # How much it's worth for a card to have this. Bad traits have negative value
    value: 35
    # Check Excel for what these mean.
    dev-stage: Discontinued
    # While card has no serious name.
    dev-name: Template trait
    # Used to order cards in Excel.
    order: 0
    # The creature's main idea.
    # Eg: "+Atk if kills creature", "-Spd in the rain".
    summary: Showcase trait YAML
    # Notes during card development.
    # Fill this with notes on usage, balancing, etc.
    # Eg: "creatures with this must cost at least 3", "rejected due to having no counter-play", "can't be paired with arrogance".
    notes: |
      * This is a template trait. Exists purely as a template. Will never be used. Isn't that so sad?
      * Grouped the fields in "data" and "metadata" groups.
        This way it's obvious what fields have gameplay influence and what fields don't.
```