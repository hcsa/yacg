# YACG

## Development

Requires

- Python 3.12
- Adobe Illustrator 2023 (v27.3)
- Installing fonts in the [fonts](/card_design/fonts) folder

### Compile executables

To create the .exe files, install pyinstaller (`pip install pyinstaller`) and run

```bash
pyinstaller scripts/yaml_to_excel.py -F -c --clean --specpath ./_pyinstaller_cache/ --distpath ./ --workpath ./_pyinstaller_cache/
```

```bash
pyinstaller scripts/excel_to_yaml.py -F -c --clean --specpath ./_pyinstaller_cache/ --distpath ./ --workpath ./_pyinstaller_cache/
```

### Fonts

This project uses two fonts, which are stored in the [fonts](/card_design/fonts) folder.

- The [icons font](./card_design/fonts/YACG-Icons.ttf) contains all icons used in cards. This font is used in card
  texts, to replace certain keywords with appropriate icons. These are stored in Unicode characters, in the block from
  U+E000 to U+F8FF (which is reserved for private use).

| **Icon** | **Keyword** | **Unicode** |
|:--------:|:-----------:|:-----------:|
| Creature | (CREATURE)  |    E100     |
|  Action  |  (ACTION)   |    E101     |
|   Aura   |   (AURA)    |    E102     |
|  Field   |   (FIELD)   |    E103     |
|    HP    |    (HP)     |    E200     |
|   Atk    |    (ATK)    |    E201     |
|   Spe    |    (SPE)    |    E202     |
| No color |  (NOCOLOR)  |    E300     |
|  Black   |   (BLACK)   |    E301     |
|   Blue   |   (BLUE)    |    E302     |
|   Cyan   |   (CYAN)    |    E303     |
|  Green   |   (GREEN)   |    E304     |
|  Orange  |  (ORANGE)   |    E305     |
|   Pink   |   (PINK)    |    E306     |
|  Purple  |  (PURPLE)   |    E307     |
|  White   |   (WHITE)   |    E308     |
|  Yellow  |  (YELLOW)   |    E309     |

- The [auxiliary font](./card_design/fonts/YACG-Auxiliary.ttf) is used when generating the cards. It contains a
  character for all the Unicode characters form the icons font, plus all the characters from U+0020 to U+007E (Basic
  Latin Unicode block, except for the C0 controls and the delete control).

The icons font is created from the .svg that live in the [icons](./card_design/icons) folder. A way of creating the font
is by uploading the .svg files to [Icomoon](https://icomoon.io/). The Icomoon project is stored as a JSON
in [icomoon_project](/card_design/icomoon_project/) folder.
> In Icomoon, strokes get ignored when generating fonts. You can convert them to fills using Illustrator — these are
> stored [here](/card_design/icons/expanded). See more details [here](https://icomoon.io/docs/#stroke-to-fill).

### Modify Illustrator scripts

These scripts are written in Python, using the `pywin32` module to access the Illustrator's COM objects and modify the
files.

The best documentation available is
the [Adobe Illustrator Scripting Guide](https://ai-scripting.docsforadobe.dev/jsobjref/javascript-object-reference.html).
Even though it's for Javascript, most of the time the references are close enough to the COM objects' interface, so they
work the same for Python. The main references are summarized in the following object model:

![Main Illustrator objects](illustrator_object_model.png)

The type library for Python has been generated to [illustrator_com.py](./scripts/yacg_python/illustrador_com.py).
Unfortunately,
the objects' properties are encoded in `_prop_map_get_` attributes, and IntelliSense's autocomplete can't make sense of
them. To get exact references for the COM objects' interface, you can:

- Use [OleViewDotNet](https://github.com/tyranid/oleviewdotnet), an open-source tool that lets you navigate the
  interface and invoke objects. You can find an overview [here](https://stackoverflow.com/a/42944052).

- Use "oleviewer.exe", a tool from
  the [Windows SDK tools](https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/).

#### Generate Illustrator type library

The type library is generated using the `makepy` CLI that comes with the `pywin32` module. This requires some work
beforehand

- Install `pywin32` module (run `pip install pywin32` or equivalent)
- Go to `[path/to/Python3/or/venv]/Scripts`, there should be a `pywin32_postinstall.py` file there
- Run `python pywin32_postinstall.py -install` with admin priviledges

After that, running [generate_illustrator_type_library.py](./scripts/generate_illustrator_type_library.py) will generate
the type library.

The documentation for type library generation can be
found [here](https://timgolden.me.uk/pywin32-docs/html/com/win32com/HTML/QuickStartClientCom.html).

## Cards YAML structure

### Creature

```yaml
creature:
  # All fields that have gameplay influence
  data:
    name: The Fred
    color: Yellow
    is-token: False
    cost-total: 4
    cost-color: 2
    hp: 3
    atk: 1
    spe: 2
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