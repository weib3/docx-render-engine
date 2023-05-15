import json
from src.document import MyDocument

f = open("demo.json", "r", encoding="UTF-8")
data = json.load(f)
f.close()

document = MyDocument()
document.render(data, "demo.docx")
