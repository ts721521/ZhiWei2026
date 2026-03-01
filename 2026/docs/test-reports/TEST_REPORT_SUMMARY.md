# ????

## 2026-02-26 ?????Google Drive ???????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 394 tests in 10.884s`?`OK`
- ??????? `tests/test_gdrive_upload_module.py`?gdrive_upload ?? 13 ?????`gdrive_upload.py` ? datetime.utcnow ???? datetime.now(timezone.utc)?

## 2026-02-26 ??????????????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 394 tests in 5.698s`?`OK`
- ?????run_workflow ????/??????????? scripts/diagnose_convert_scope.py ?????

## ??


```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

## ???????????????2026-02-24?

| ??     | ?? |
|----------|------|
| ???? | 280  |
| ??     | 280  |
| ??     | 0    |
| ??     | 0    |
| ???   | 0    |
| ??     | OK   |

??????

```text
Ran 280 tests in 8.234s
OK
```

**??**??????????????????????? Schema ?????????????????

?????????????????

## ??????

- ???????????????????/???????????????????
- GUI ????Run Mode ??????? Mixin ???????????
- ?????????????/????????????????????
- ????? i18n??? schema ????? key ??????
- ?????`pypdf` API ??????

## ??????????????

- `test_task_list_filter_sort.py`????????/?????????????????
- `test_default_config_schema.py`?`ui.task_current_config_only` ???? schema ????
- `test_nonfatal_ui_error_reporting.py`?UI ?????????????????
- `test_gui_run_mode_state_behavior.py`?run mode ????????? `except` ???

## ?????

- [TEST_REPORT_GUI_TASK_MODE.md](TEST_REPORT_GUI_TASK_MODE.md)
- [TEST_REPORT_TASK_MANAGER.md](TEST_REPORT_TASK_MANAGER.md)
- [TEST_REPORT_CONVERTER_RESUME.md](TEST_REPORT_CONVERTER_RESUME.md)
- [TEST_REPORT_MERGE_CONVERT_PIPELINE.md](TEST_REPORT_MERGE_CONVERT_PIPELINE.md)
- [TEST_REPORT_OUTPUT_CONTROLS.md](TEST_REPORT_OUTPUT_CONTROLS.md)

## 2026-02-24 ???????markdown_quality_report ????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 193 tests in 3.559s`
- ???`FAILED (failures=1, errors=4, skipped=2)`
- ???
  - `test_gui_task_mode.TestGuiTaskTabVisibility.test_classic_mode_hides_task_tab`????? Tab ?????
- ???
  - `test_pypdf`
  - `test_pypdf_dir`
  - `test_pypdf_ver`
  - `test_pypdf_ver2`
  - ?? 4 ??? `ModuleNotFoundError: No module named 'pypdf'`

## 2026-02-24 ???????markdown_render ????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 195 tests in 1.979s`
- ???`FAILED (failures=1, errors=4, skipped=2)`
- ???
  - `test_gui_task_mode.TestGuiTaskTabVisibility.test_classic_mode_hides_task_tab`????? Tab ?????
- ???
  - `test_pypdf`
  - `test_pypdf_dir`
  - `test_pypdf_ver`
  - `test_pypdf_ver2`
  - ?? 4 ??? `ModuleNotFoundError: No module named 'pypdf'`

## 2026-02-24 ???????merge_tasks ????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 199 tests in 7.877s`
- ???`OK`

## 2026-02-24 ???????pdf_markdown_export / cab_convert / merge_index_doc / convert_thread / records_json_export?

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???
  - `Ran 201 tests in 6.379s`?`OK`
  - `Ran 205 tests in 7.455s`?`OK`
  - `Ran 207 tests in 6.919s`?`OK`
  - `Ran 209 tests in 6.617s`?`OK`

## 2026-02-24 ???????excel_defined_names ????

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 211 tests in 6.658s`
- ???`OK`


## 2026-02-24 ???????error_recording / index_runtime / markdown_render(table helper)?

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 229 tests in 6.692s`
- ???`OK`

## 2026-02-24 ????????fast_md_engine + prompt_wrapper + traceability???

- ???`python -m unittest discover -s tests -p "test_*.py" -v`
- ???`Ran 238 tests in 7.350s`
- ???`OK`

## 2026-02-24 ??????????merge_markdowns / merge_mode_pipeline / trace_map ????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 245 tests in 7.709s`
- ????`OK`

## 2026-02-24 ??????????_convert_on_mac / confirm_config_in_terminal / _get_local_app??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 251 tests in 7.778s`
- ????`OK`

## 2026-02-24 ??????????_extract_sheet_charts / _setup_excel_pages / _write_excel_structured_json_exports??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 257 tests in 11.077s`
- ????`OK`

## 2026-02-24 ??????????_export_pdf_markdown / _write_chromadb_export??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 261 tests in 7.324s`
- ????`OK`

## 2026-02-24 ??????????scan_excel_content_in_thread / convert_logic_in_thread / _build_ai_output_path_from_source / load_config??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 269 tests in 7.794s`
- ????`OK`

## 2026-02-24 ??????????V6.0 ?????? 8.1??markitdown ??? MVP??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 271 tests in 7.447s`
- ????`OK`

## 2026-02-24 ???????????GUI mixin ??????? gui/mixins??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 271 tests in 7.932s`
- ????`OK`

## 2026-02-24 ???????????GUI ?????? + ?????????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 271 tests in 6.717s`
- ????`OK`

## 2026-02-24 ???????????legacy_shims ???? gui ????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 271 tests in 6.688s`
- ????`OK`

## 2026-02-24 ???????????mixin ?????????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 271 tests in 6.964s`
- ????`OK`

## 2026-02-24 ???????????????????????? + ??????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 274 tests in 6.978s`
- ????`OK`

## 2026-02-24 ???????????????????????? pre-commit + CI??
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 276 tests in 11.165s`
- ????`OK`

## 2026-02-24 ??????????batch_parallel ???? + scan ??????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 277 tests in 7.654s`
- ????`OK`

## 2026-02-24 ?????????????? Schema ???
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 280 tests in 8.234s`
- ????`OK`

## 2026-02-24 ????????????????????????????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 282 tests in 7.325s`
- ????`OK`

## 2026-02-24 ???????????run_workflow / scan_convert_candidates ????????
- ????`python -m unittest discover -s tests -p "test_*.py" -v`
- ?????`Ran 282 tests in 7.730s`
- ????`OK`

## 2026-02-25 TaskWorkflowMixin exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 287 tests in 7.575s`
- Status: `OK`
## 2026-02-25 TaskWorkflowMixin exception narrowing regression (round 2)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 288 tests in 6.917s`
- Status: `OK`
## 2026-02-25 No folder-open during tests regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 290 tests in 6.623s`
- Status: `OK`
## 2026-02-25 MiscUIMixin exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 291 tests in 7.886s`
- Status: `OK`
## 2026-02-25 Artifact meta exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 291 tests in 6.219s`
- Status: `OK`
## 2026-02-25 AI paths exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 294 tests in 7.750s`
- Status: `OK`
## 2026-02-25 Checkpoint utils exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 295 tests in 6.126s`
- Status: `OK`
## 2026-02-25 Config load exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 296 tests in 6.100s`
- Status: `OK`
## 2026-02-25 CAB convert exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 297 tests in 7.740s`
- Status: `OK`
## 2026-02-25 Default config exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 298 tests in 5.903s`
- Status: `OK`
## 2026-02-25 CAB extract exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 299 tests in 6.056s`
- Status: `OK`## 2026-02-25 Failure stage exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 300 tests in 6.299s`
- Status: `OK`
## 2026-02-25 Index runtime exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 301 tests in 5.951s`
- Status: `OK`
## 2026-02-25 Platform utils exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 302 tests in 6.140s`
- Status: `OK`
## 2026-02-25 Local office app exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 303 tests in 6.162s`
- Status: `OK`
## 2026-02-25 Callback utils exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 304 tests in 6.126s`
- Status: `OK`
## 2026-02-25 Sandbox guard exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 305 tests in 6.432s`
- Status: `OK`
## 2026-02-25 Interactive prompts exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 306 tests in 6.414s`
- Status: `OK`
## 2026-02-25 File registry exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 307 tests in 6.493s`
- Status: `OK`
## 2026-02-25 Office cycle exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 308 tests in 7.795s`
- Status: `OK`
## 2026-02-25 Source roots exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 309 tests in 7.979s`
- Status: `OK`
## 2026-02-25 Traceability exception narrowing regression
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 310 tests in 6.162s`
- Status: `OK`
## 2026-02-25 Batch exception narrowing regression (file_ops + excel utils)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 314 tests in 6.554s`
- Status: `OK`
## 2026-02-25 Batch exception narrowing regression (failure/mshelp/incremental/update)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 320 tests in 6.756s`
- Status: `OK`
## 2026-02-25 Batch exception narrowing regression (trace/merge/mshelp/pdf_md)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 324 tests in 6.187s`
- Status: `OK`
## 2026-02-25 Batch exception narrowing regression (trace/merge/mshelp/pdf_md round 2)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 324 tests in 6.112s`
- Status: `OK`

## 2026-02-25 Batch exception narrowing regression (batch/chromadb/excel/convert/corpus)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 338 tests in 6.203s`
- Status: `OK`

## 2026-02-25 Batch exception narrowing regression (collect/incremental/mshelp/process/safe_exec)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 347 tests in 5.926s`
- Status: `OK`

## 2026-02-25 Batch exception narrowing regression (merge_pdfs + office_converter final)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 349 tests in 5.924s`
- Status: `OK`

## 2026-02-25 Split round (interactive choices extraction)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 352 tests in 5.771s`
- Status: `OK`

## 2026-02-25 Split round (runtime helpers extraction batch)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 358 tests in 5.948s`
- Status: `OK`

## 2026-02-25 Split round (bootstrap state + lifecycle extraction)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 362 tests in 5.847s`
- Status: `OK`

## 2026-02-25 Split round (path config + perf metrics extraction)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 367 tests in 5.745s`
- Status: `OK`

## 2026-02-25 Split round (merge/prompt/excel/markdown extraction batch)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 371 tests in 6.198s`
- Status: `OK`

## 2026-02-25 Split round (cli wizard + display + lifecycle kill app extraction)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 375 tests in 5.962s`
- Status: `OK`

## 2026-02-25 Split round (state writeback extraction batch)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 377 tests in 5.785s`
- Status: `OK`

## 2026-02-25 Split round (output state writeback extraction batch)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 377 tests in 6.122s`
- Status: `OK`

## 2026-02-25 Split round (failed/process/merge-map wrapper extraction)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 377 tests in 6.163s`
- Status: `OK`
## 2026-02-25 Split round (final six-method extraction closure)
- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 381 tests in 8.897s`
- Status: `OK`
## 2026-02-25 Docs sync round (handover + summary refresh)
- Scope: docs only
- Status: completed and pushed

## 2026-02-27 GUI toggle for Markdown image manifest

- Command: `python -m unittest discover -s tests -p "test_*.py" -v`
- Result: `Ran 401 tests in 6.578s`
- Status: `OK`
- Change summary:
  - Added GUI toggle for `enable_markdown_image_manifest` (runtime AI export area).
  - Wired config load/save/compose/dirty snapshot paths and manual runtime execution path.
  - Wired task runtime override/apply paths so task mode uses the same setting.
