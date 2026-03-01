# -*- coding: utf-8 -*-
"""Shared constants extracted from office_converter.py."""

# Office constants
wdFormatPDF = 17
xlTypePDF = 0
ppSaveAsPDF = 32
ppFixedFormatTypePDF = 2
xlPDF_SaveAs = 57
xlRepairFile = 1

# Engine types
ENGINE_WPS = "wps"
ENGINE_MS = "ms"
ENGINE_ASK = "ask"

# Process cleanup strategies
KILL_MODE_ASK = "ask"
KILL_MODE_AUTO = "auto"
KILL_MODE_KEEP = "keep"

# Main run modes
MODE_CONVERT_ONLY = "convert_only"
MODE_MERGE_ONLY = "merge_only"
MODE_CONVERT_THEN_MERGE = "convert_then_merge"
MODE_COLLECT_ONLY = "collect_only"  # collect and deduplicate mode
MODE_MSHELP_ONLY = "mshelp_only"  # dedicated mode for MSHelpViewer API docs

# Merge & convert sub-modes under MODE_MERGE_ONLY
MERGE_CONVERT_SUBMODE_MERGE_ONLY = "merge_only"
MERGE_CONVERT_SUBMODE_PDF_TO_MD = "pdf_to_md"

# collect_only sub-modes
COLLECT_MODE_COPY_AND_INDEX = "copy_and_index"  # dedup + copy + Excel
COLLECT_MODE_INDEX_ONLY = "index_only"  # Excel only, no copying

# Merge modes
MERGE_MODE_CATEGORY = "category_split"  # split by Price_/Word_/Excel_ categories
MERGE_MODE_ALL_IN_ONE = "all_in_one"  # merge all PDFs into one file

# Content processing strategy (conversion mode only)
STRATEGY_STANDARD = "standard"  # classify only by extension
STRATEGY_SMART_TAG = "smart_tag"  # filename/content keyword hit -> Price_
STRATEGY_PRICE_ONLY = "price_only"  # process only keyword-matching files

ERR_RPC_SERVER_BUSY = -2147417846
ERR_RPC_SERVER_UNAVAILABLE = -2147023174
DEFAULT_SHORT_ID_LEN = 8
