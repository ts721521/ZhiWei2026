# -*- coding: utf-8 -*-
"""Terminal config confirmation helpers extracted from office_converter.py."""

import os


def confirm_config_in_terminal(
    converter,
    *,
    input_fn=input,
    print_fn=print,
    abspath_fn=os.path.abspath,
):
    converter.print_step_title("Step 1/4: Confirm Source and Target")
    print_fn("Current paths:")
    print_fn(f"  source: {converter.config['source_folder']}")
    print_fn(f"  target: {converter.config['target_folder']}")
    print_fn("-" * 60)

    while True:
        choice = input_fn("Modify these paths? [y/N]: ").strip().lower()
        if choice in ("", "n"):
            break
        if choice == "y":
            print_fn("\n=== Edit Paths ===")
            print_fn(f"Current source: {converter.config['source_folder']}")
            new_source = (
                input_fn("New source (Enter to keep): ")
                .strip()
                .replace('"', "")
                .replace("'", "")
            )
            if new_source:
                converter.config["source_folder"] = abspath_fn(new_source)

            print_fn(f"\nCurrent target: {converter.config['target_folder']}")
            new_target = (
                input_fn("New target (Enter to keep): ")
                .strip()
                .replace('"', "")
                .replace("'", "")
            )
            if new_target:
                converter.config["target_folder"] = abspath_fn(new_target)

            converter.save_config()
            print_fn("Config saved.")
            print_fn("-" * 60)
            print_fn("Updated paths:")
            print_fn(f"  source: {converter.config['source_folder']}")
            print_fn(f"  target: {converter.config['target_folder']}")
            print_fn("-" * 60)
