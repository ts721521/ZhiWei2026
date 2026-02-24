# -*- coding: utf-8 -*-
"""Checkpoint helpers extracted from office_converter.py."""

import hashlib
import json
import logging
import os
from datetime import datetime


def get_checkpoint_path(config):
    """Get checkpoint file path for current config."""
    target_folder = config.get("target_folder", "")
    if not target_folder:
        return None
    checkpoint_dir = os.path.join(target_folder, "_AI", "checkpoints")
    os.makedirs(checkpoint_dir, exist_ok=True)
    # Use source-folder hash to keep checkpoint filename stable per source root.
    source_key = config.get("source_folder", "default")
    source_hash = hashlib.md5(source_key.encode()).hexdigest()[:12]
    return os.path.join(checkpoint_dir, f"batch_{source_hash}.json")


def save_checkpoint(checkpoint, checkpoint_path):
    """Persist checkpoint data to file."""
    if not checkpoint or not checkpoint_path:
        return
    try:
        checkpoint["updated_at"] = datetime.now().isoformat(timespec="seconds")
        with open(checkpoint_path, "w", encoding="utf-8") as f:
            json.dump(checkpoint, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.warning(f"Failed to save checkpoint: {e}")


def mark_file_done_in_checkpoint(checkpoint, file_path):
    """Mark one file as completed in checkpoint."""
    if not checkpoint:
        return checkpoint

    completed = checkpoint.setdefault("completed_files", [])
    file_path_normalized = os.path.abspath(file_path)
    if file_path_normalized not in completed:
        completed.append(file_path_normalized)

    planned = checkpoint.get("planned_files", [])
    if len(completed) >= len(planned):
        checkpoint["status"] = "completed"
    return checkpoint


def clear_checkpoint_file(checkpoint_path):
    """Remove checkpoint file if exists."""
    if checkpoint_path and os.path.exists(checkpoint_path):
        try:
            os.remove(checkpoint_path)
            logging.info(f"Checkpoint cleared: {checkpoint_path}")
        except Exception as e:
            logging.warning(f"Failed to clear checkpoint: {e}")
