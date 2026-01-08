
import pypdf
import sys
print(f"pypdf version: {pypdf.__version__}", flush=True)
try:
    import pypdf.annotations
    print("pypdf.annotations imported", flush=True)
except ImportError as e:
    print(f"ImportError: {e}", flush=True)
