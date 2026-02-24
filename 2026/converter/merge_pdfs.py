# -*- coding: utf-8 -*-
"""PDF merge workflow extracted from office_converter.py."""

import logging
import os
import traceback
from datetime import datetime


def merge_pdfs(
    converter,
    *,
    has_pypdf,
    has_openpyxl,
    pdf_writer_cls,
    pdf_reader_cls,
    workbook_cls,
    pythoncom_mod,
    win32_client,
):
    if not converter.config.get("enable_merge", True):
        return []
    if not has_pypdf:
        print("\n[INFO] pypdf not found. Skip merge step. Run: pip install pypdf")
        logging.warning("pypdf not found. Skip merge step.")
        return []

    try:
        from pypdf.generic import (
            DictionaryObject,
            NumberObject,
            FloatObject,
            NameObject,
            TextStringObject,
            ArrayObject,
            RectangleObject,
        )

        has_pypdf_generic = True
    except ImportError:
        has_pypdf_generic = False

    print("\n" + "=" * 60)
    print("  Start PDF merging ...")
    print(f"  Merge mode: {converter.get_readable_merge_mode()} ({converter.merge_mode})")
    print(f"  Merge output dir: {converter.merge_output_dir}")
    if converter.enable_merge_index:
        print("  [Option] Enable index page generation (with clickable links)")
    if converter.enable_merge_excel:
        print("  [Option] Enable Excel list output (one row per source file)")
    print("=" * 60)

    wb_merge = None
    ws_merge = None
    merge_excel_path = None
    if converter.enable_merge_excel:
        if not has_openpyxl:
            print("  [WARN] openpyxl not found. Excel merge list is disabled.")
        else:
            timestamp_excel = datetime.now().strftime("%Y%m%d_%H%M%S")
            merge_excel_path = os.path.join(
                converter.merge_output_dir, f"Merge_List_{timestamp_excel}.xlsx"
            )
            wb_merge = workbook_cls()
            ws_merge = wb_merge.active
            ws_merge.title = "MergeList"
            ws_merge.append(["Merged File", "Source Files"])
            ws_merge.column_dimensions["A"].width = 40
            ws_merge.column_dimensions["B"].width = 60

    merge_tasks = converter._get_merge_tasks()

    total_tasks = len(merge_tasks)
    print(f"  Total merge tasks: {total_tasks}")

    word_app = None
    if converter.enable_merge_index and total_tasks > 0:
        try:
            pythoncom_mod.CoInitialize()
            word_app = win32_client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
        except Exception as exc:
            logging.error(f"Failed to start Word. Cannot generate index page: {exc}")
            word_app = None

    generated_outputs = []
    generated_map_outputs = []
    merge_index_records = []
    merge_batch_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    converter.merge_excel_path = None

    for idx, (output_filename, group) in enumerate(merge_tasks, 1):
        print(f"  Processing [{idx}/{total_tasks}]: {output_filename}")

        if ws_merge:
            for sub_file in group:
                ws_merge.append([output_filename, os.path.basename(sub_file)])

        output_path = os.path.join(converter.merge_output_dir, output_filename)

        try:
            merger = pdf_writer_cls()
            current_page_index = 0
            map_records = []
            short_id_taken = set()

            index_pdf_path = None
            file_basenames = [os.path.basename(path) for path in group]

            file_start_pages = []

            if converter.enable_merge_index and word_app:
                index_pdf_path = converter._create_index_doc_and_convert(
                    word_app, file_basenames, "File Index"
                )

            if index_pdf_path and os.path.exists(index_pdf_path):
                idx_reader = pdf_reader_cls(index_pdf_path)
                index_page_count = len(idx_reader.pages)

                for page in idx_reader.pages:
                    merger.add_page(page)

                current_page_index += index_page_count
            else:
                index_page_count = 0

            for source_idx, pdf_file in enumerate(group, 1):
                fname = os.path.basename(pdf_file)

                file_start_pages.append(current_page_index)

                try:
                    reader = pdf_reader_cls(pdf_file)
                    source_page_count = len(reader.pages)
                    source_md5 = converter._compute_md5(pdf_file)
                    source_short_id = converter._build_short_id(source_md5, short_id_taken)
                    bookmark_title = fname
                    if converter.config.get("bookmark_with_short_id", True):
                        bookmark_title = f"[ID:{source_short_id}] {fname}"

                    merger.add_outline_item(bookmark_title, current_page_index)

                    start_page_1based = current_page_index + 1
                    end_page_1based = current_page_index + source_page_count

                    for page in reader.pages:
                        merger.add_page(page)
                    current_page_index += source_page_count

                    source_abs_path = os.path.abspath(pdf_file)
                    source_rel_path = source_abs_path
                    try:
                        source_rel_path = os.path.relpath(
                            source_abs_path,
                            converter._get_source_root_for_path(source_abs_path),
                        )
                    except Exception:
                        pass

                    map_records.append(
                        {
                            "merge_batch_id": merge_batch_id,
                            "merged_pdf_name": os.path.basename(output_path),
                            "merged_pdf_path": output_path,
                            "source_index": source_idx,
                            "source_filename": fname,
                            "source_abspath": source_abs_path,
                            "source_relpath": source_rel_path,
                            "source_md5": source_md5,
                            "source_short_id": source_short_id,
                            "start_page_1based": start_page_1based,
                            "end_page_1based": end_page_1based,
                            "page_count": source_page_count,
                            "bookmark_title": bookmark_title,
                        }
                    )

                    if converter.config.get("privacy", {}).get("mask_md5_in_logs", True):
                        md5_log = converter._mask_md5(source_md5)
                    else:
                        md5_log = source_md5
                    logging.info(
                        f"merge map record: {fname} | pages {start_page_1based}-{end_page_1based} | ID={source_short_id} | MD5={md5_log}"
                    )
                except Exception as exc:
                    logging.error(f"merge read failed {pdf_file}: {exc}")

            if index_page_count > 0 and has_pypdf_generic:
                start_y = 715
                line_height = 20

                lines_per_page = 32

                for i, target_page_num in enumerate(file_start_pages):
                    page_idx = i // lines_per_page
                    row_idx = i % lines_per_page
                    if page_idx >= index_page_count:
                        break

                    idx_page = merger.pages[page_idx]

                    rect_top = start_y - (row_idx * line_height)
                    rect_bottom = rect_top - line_height
                    rect = [72, rect_bottom, 520, rect_top]

                    target_page_obj = merger.pages[target_page_num]
                    target_page_ref = getattr(target_page_obj, "indirect_ref", None)
                    if target_page_ref is None:
                        target_page_ref = getattr(
                            target_page_obj,
                            "indirect_reference",
                            None,
                        )

                    if target_page_ref is None:
                        logging.warning(
                            f"cannot resolve target page reference {target_page_num} (indirect_ref/reference). skip index link."
                        )
                        continue

                    link_annotation = DictionaryObject()
                    link_annotation.update(
                        {
                            NameObject("/Type"): NameObject("/Annot"),
                            NameObject("/Subtype"): NameObject("/Link"),
                            NameObject("/Rect"): ArrayObject(
                                [FloatObject(c) for c in rect]
                            ),
                            NameObject("/Border"): ArrayObject(
                                [NumberObject(0), NumberObject(0), NumberObject(0)]
                            ),
                            NameObject("/Dest"): ArrayObject(
                                [target_page_ref, NameObject("/Fit")]
                            ),
                        }
                    )

                    if "/Annots" not in idx_page:
                        idx_page[NameObject("/Annots")] = ArrayObject()

                    idx_page["/Annots"].append(link_annotation)

            merger.write(output_path)
            merger.close()

            if not os.path.exists(output_path):
                raise RuntimeError(f"merged output file not generated: {output_path}")
            try:
                if os.path.getsize(output_path) <= 0:
                    raise RuntimeError(
                        f"merged output file size invalid (0 bytes): {output_path}"
                    )
            except OSError:
                pass

            generated_outputs.append(output_path)
            if map_records:
                merged_pdf_md5 = ""
                try:
                    merged_pdf_md5 = converter._compute_md5(output_path)
                except Exception:
                    pass
                for rec in map_records:
                    rec["merged_pdf_md5"] = merged_pdf_md5
                merge_index_records.extend(map_records)

            if converter.config.get("enable_merge_map", True):
                try:
                    csv_path, json_path = converter._write_merge_map(output_path, map_records)
                    if csv_path and json_path:
                        generated_map_outputs.extend([csv_path, json_path])
                        logging.info(f"map file generated: {csv_path}")
                        logging.info(f"map file generated: {json_path}")
                except Exception as exc:
                    logging.error(f"failed to write map files {output_path}: {exc}")

            if index_pdf_path and os.path.exists(index_pdf_path):
                os.remove(index_pdf_path)

        except Exception as exc:
            print(f" [FAILED] {exc}")
            logging.error(f"merge task failed {output_filename}: {exc}")
            traceback.print_exc()

    if word_app:
        try:
            word_app.Quit()
        except Exception:
            pass
        pythoncom_mod.CoUninitialize()

    converter.merge_index_records = merge_index_records
    converter.generated_merge_outputs = list(generated_outputs)
    converter.generated_map_outputs = list(generated_map_outputs)

    if wb_merge:
        try:
            if merge_index_records:
                ws_merge_index = wb_merge.create_sheet("MergeIndex")
                converter._write_merge_index_sheet(ws_merge_index, merge_index_records)
            if converter.conversion_index_records:
                ws_conv_index = wb_merge.create_sheet("ConvertedPDFs")
                converter._write_conversion_index_sheet(
                    ws_conv_index,
                    converter.conversion_index_records,
                )
            if ws_merge:
                converter._style_header_row(ws_merge)
                converter._auto_fit_sheet(ws_merge)
            wb_merge.save(merge_excel_path)
            converter.merge_excel_path = merge_excel_path
            print(f"\n  Excel index saved: {merge_excel_path}")
        except Exception as exc:
            logging.error(f"failed to save Excel merge list: {exc}")

    if total_tasks <= 0:
        print("\n  [INFO] No merge tasks generated. Ensure PDF files exist in scan results.")
    elif len(generated_outputs) <= 0:
        print(
            "\n  [INFO] Merge tasks executed, but no output was generated. Check logs and output permissions."
        )
    else:
        print("\n  Merged output files:")
        for path in generated_outputs:
            print(f"  - {path}")
        if generated_map_outputs:
            print("\n  Map files:")
            for path in generated_map_outputs:
                print(f"  - {path}")

    return generated_outputs
