# -*- coding: utf-8 -*-
"""File hashing and hyperlink helpers extracted from office_converter.py."""

import hashlib
import os

from converter.constants import DEFAULT_SHORT_ID_LEN


def compute_md5(path, block_size=1024 * 1024):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(block_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def mask_md5(md5_value):
    if not md5_value or len(md5_value) < 12:
        return md5_value
    return f"{md5_value[:8]}...{md5_value[-4:]}"


def build_short_id(md5_value, taken_ids, default_len=DEFAULT_SHORT_ID_LEN):
    length = default_len
    while length <= len(md5_value):
        candidate = md5_value[:length].upper()
        if candidate not in taken_ids:
            taken_ids.add(candidate)
            return candidate
        length += 2
    candidate = md5_value.upper()
    taken_ids.add(candidate)
    return candidate


def compute_file_hash(path, block_size=1024 * 1024):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(block_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def make_file_hyperlink(path: str) -> str:
    path = os.path.abspath(path)
    return "file:///" + path.replace("\\", "/")
