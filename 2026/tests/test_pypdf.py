
from pypdf import PdfWriter, PdfReader

try:
    writer = PdfWriter()
    # Create a dummy blank page
    writer.add_blank_page(width=595, height=842)
    writer.add_blank_page(width=595, height=842)
    
    # Try to add a link
    # signature: add_link(page_nr, target_page_nr, rect, border=None)
    # rect is [x1, y1, x2, y2]
    writer.add_link(0, 1, [50, 50, 200, 100])
    
    print("pypdf add_link seems available and signature matches expectation.")
except Exception as e:
    print(f"Error: {e}")
