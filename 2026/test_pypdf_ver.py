
import pypdf
print(f"pypdf version: {pypdf.__version__}")
try:
    from pypdf.annotations import FreeText
    print("pypdf.annotations exists")
    print(dir(pypdf.annotations))
except ImportError:
    print("pypdf.annotations not found")
