# -*- coding: utf-8 -*-
"""Locator/query related methods extracted from OfficeGUI."""

import json
import os
import subprocess
import glob
from types import SimpleNamespace

from converter.traceability import normalize_short_id_for_match
from locate_source import locate_by_page, locate_by_short_id
from search_adapter import EverythingAdapter, build_listary_query


class LocatorMixin:
    def get_locator_map_dir(self):
        target = self.var_target_folder.get().strip()
        if not target:
            return ""
        return os.path.join(target, "_MERGED")

    def _set_locator_action_state(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn_name in (
            "btn_locator_open_file",
            "btn_locator_open_dir",
            "btn_locator_everything",
            "btn_locator_copy_listary",
        ):
            btn = getattr(self, btn_name, None)
            if btn is not None:
                btn.configure(state=state)

    def refresh_locator_maps(self):
        map_dir = self.get_locator_map_dir()
        self.locator_short_id_index = {}
        self.last_locate_record = None
        self._set_locator_action_state(False)
        if not map_dir or not os.path.isdir(map_dir):
            self.cb_locator_merged.configure(values=[])
            self.var_locator_result.set(self.tr("msg_locator_no_merged_dir"))
            return

        merged_names = []
        map_files = glob.glob(os.path.join(map_dir, "*.map.json"))
        map_files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for p in map_files:
            stem = os.path.basename(p)[:-9]  # remove .map.json
            merged_names.append(f"{stem}.pdf")
            try:
                with open(p, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                for rec in payload.get("records", []):
                    sid_raw = str(rec.get("source_short_id", "")).strip().upper()
                    sid_norm = normalize_short_id_for_match(sid_raw)
                    if sid_raw:
                        self.locator_short_id_index.setdefault(sid_raw, []).append(rec)
                    if sid_norm and sid_norm != sid_raw:
                        self.locator_short_id_index.setdefault(sid_norm, []).append(rec)
            except Exception:
                continue

        self.cb_locator_merged.configure(values=merged_names)
        current = self.var_locator_merged.get().strip()
        if merged_names:
            if current not in merged_names:
                self.var_locator_merged.set(merged_names[0])
        else:
            self.var_locator_merged.set("")
        self.var_locator_result.set(
            self.tr("msg_locator_loaded_maps").format(len(merged_names))
        )

    def run_locator_query(self):
        if not self.validate_runtime_inputs(silent=False, scope="locator"):
            return
        map_dir = self.get_locator_map_dir()
        if not os.path.isdir(map_dir):
            self.var_locator_result.set(self.tr("msg_locator_map_missing"))
            return

        merged_name = self.var_locator_merged.get().strip()
        page_raw = self.var_locator_page.get().strip()
        short_id = self.var_locator_short_id.get().strip()
        priority_note = ""

        result = None
        if page_raw and short_id and merged_name:
            priority_note = f"{self.tr('msg_locator_priority_page')} "
        if page_raw and merged_name:
            try:
                page = int(page_raw)
            except ValueError:
                self.var_locator_result.set(self.tr("msg_locator_invalid_page"))
                return
            result = locate_by_page(merged_name, page, map_dir)
        elif short_id:
            sid = short_id.upper()
            sid_norm = normalize_short_id_for_match(sid)
            cache_hits = self.locator_short_id_index.get(sid, [])
            if not cache_hits and sid_norm:
                cache_hits = self.locator_short_id_index.get(sid_norm, [])
            if len(cache_hits) == 1:
                r = cache_hits[0]
                t = SimpleNamespace(
                    source_filename=r.get("source_filename", ""),
                    start_page_1based=int(r.get("start_page_1based", 0)),
                    end_page_1based=int(r.get("end_page_1based", 0)),
                    source_short_id=r.get("source_short_id", ""),
                    source_abspath=r.get("source_abspath", ""),
                    source_md5=r.get("source_md5", ""),
                )
                result = SimpleNamespace()
                result.record = t
                result.alternatives = []
                result.status = "ok"
            elif len(cache_hits) > 1:
                result = SimpleNamespace()
                result.record = None
                result.status = "ambiguous"
                alts = []
                for r in cache_hits[:2]:
                    a = SimpleNamespace(
                        source_filename=r.get("source_filename", ""),
                        start_page_1based=int(r.get("start_page_1based", 0)),
                        end_page_1based=int(r.get("end_page_1based", 0)),
                    )
                    alts.append(a)
                result.alternatives = alts
            else:
                result = locate_by_short_id(short_id, map_dir)
        else:
            self.var_locator_result.set(self.tr("msg_locator_need_input"))
            return

        if result.record:
            self.last_locate_record = result.record
            self.var_locator_result.set(
                priority_note
                + self.tr("msg_locator_hit").format(
                    result.record.source_filename,
                    result.record.start_page_1based,
                    result.record.end_page_1based,
                    result.record.source_short_id,
                )
            )
            if self._read_config_value(["listary", "copy_query_on_locate"], False):
                self.copy_listary_query(silent=True)
            self._set_locator_action_state(True)
            return

        self.last_locate_record = None
        self._set_locator_action_state(False)
        if result.alternatives:
            alt = ", ".join(
                [
                    f"{x.source_filename}({x.start_page_1based}-{x.end_page_1based})"
                    for x in result.alternatives[:2]
                ]
            )
            self.var_locator_result.set(
                priority_note + self.tr("msg_locator_miss_alt").format(alt)
            )
        else:
            self.var_locator_result.set(
                priority_note + self.tr("msg_locator_status").format(result.status)
            )

    def open_locator_file(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        path = self.last_locate_record.source_abspath
        if not os.path.exists(path):
            self.var_locator_result.set(
                self.tr("msg_locator_file_missing").format(path)
            )
            return
        self._open_path(path)

    def open_locator_folder(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        path = self.last_locate_record.source_abspath
        folder = os.path.dirname(path)
        if not os.path.isdir(folder):
            self.var_locator_result.set(
                self.tr("msg_locator_dir_missing").format(folder)
            )
            return
        if sys.platform == "win32":
            subprocess.run(["explorer", "/select,", path], check=False)
        else:
            self._open_path(folder)

    def search_with_everything(self):
        if not self.last_locate_record:
            self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return

        if not self._read_config_value(["everything", "enabled"], True):
            self.var_locator_result.set(self.tr("msg_locator_everything_disabled"))
            return

        es_path = self._read_config_value(["everything", "es_path"], "")
        timeout_ms = self._read_config_value(["everything", "timeout_ms"], 1500)
        prefer_path_exact = self._read_config_value(
            ["everything", "prefer_path_exact"], True
        )
        adapter = EverythingAdapter(es_path=es_path, timeout_ms=int(timeout_ms))
        if not adapter.is_available():
            self.var_locator_result.set(self.tr("msg_locator_everything_notfound"))
            return

        directory = (
            os.path.dirname(self.last_locate_record.source_abspath)
            if prefer_path_exact
            else ""
        )
        ret = adapter.run_query(self.last_locate_record.source_filename, directory)
        if ret.ok:
            self.var_locator_result.set(self.tr("msg_locator_everything_ok"))
        else:
            self.var_locator_result.set(
                self.tr("msg_locator_everything_fail").format(ret.stderr)
            )

    def copy_listary_query(self, silent=False):
        if not self.last_locate_record:
            if not silent:
                self.var_locator_result.set(self.tr("msg_locator_run_first"))
            return
        query = build_listary_query(
            self.last_locate_record.source_short_id,
            self.last_locate_record.source_md5,
            self.last_locate_record.source_filename,
            self.last_locate_record.source_abspath,
        )
        self.clipboard_clear()
        self.clipboard_append(query)
        self.update_idletasks()
        if not silent:
            self.var_locator_result.set(self.tr("msg_locator_listary_copied"))

    def _read_config_value(self, key_path, default_value):
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            val = cfg
            for k in key_path:
                if not isinstance(val, dict):
                    return default_value
                val = val.get(k)
            return default_value if val is None else val
        except Exception:
            return default_value

