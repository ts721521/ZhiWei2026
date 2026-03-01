import os,sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from office_converter import OfficeConverter
cfg = r'd:\GitHub\ZhiWei2026\2026\configs\scenarios\notebooklm\config.notebooklm_full_md_merge_run.json'
conv = OfficeConverter(cfg, interactive=False)
conv.run()
