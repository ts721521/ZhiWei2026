
from pypdf import PdfWriter
print([x for x in dir(PdfWriter) if x.startswith("add")])
