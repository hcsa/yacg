import win32com.client.gencache
import yacg_python.illustrador_com as Illustrator

# combrowse.main()
# controler: Illustrator._Application = win32com.client.Dispatch("Illustrator.Application")
app = win32com.client.GetActiveObject("Illustrator.Application")._dispobj_
# win32com.client.gencache.EnsureModule('{11C6EAB4-152A-4829-9DDB-5DE14E87A5DD}', 0, 1, 0)
# out = win32com.client.gencache.EnsureModule('{95CD20AA-AD72-11D3-B086-0010A4F5C335}', 0, 1, 0)
# combrowse.main()

# controller = win32com.client.Dispatch("Illustrator.Application")

print("HERE")
fontToChange = app.TextFonts.GetFontByName("Times-Bold")
doc = app.ActiveDocument
doc.Layers.Item("CreatureLayer").TextFrames.Item("TextDescription").Words.Item(3).CharacterAttributes.TextFont = fontToChange
print("HERE 2")