# -*- coding: utf-8 -*-
"""Excel defined names extractor extracted from office_converter.py."""

import json


def extract_workbook_defined_names(wb_formula):
    names = []
    if wb_formula is None:
        return names

    dn_container = getattr(wb_formula, "defined_names", None)
    if dn_container is None:
        return names

    dn_objects = []
    raw_list = getattr(dn_container, "definedName", None)
    if raw_list:
        dn_objects.extend(list(raw_list))
    try:
        for _, dn in dn_container.items():
            if isinstance(dn, (list, tuple, set)):
                dn_objects.extend(list(dn))
            else:
                dn_objects.append(dn)
    except (TypeError, ValueError, AttributeError):
        pass

    seen = set()
    for dn in dn_objects:
        if dn is None:
            continue
        name = str(getattr(dn, "name", "") or "")
        local_sheet_id = getattr(dn, "localSheetId", None)
        hidden = bool(getattr(dn, "hidden", False))
        attr_text = str(getattr(dn, "attr_text", "") or "")
        comment = str(getattr(dn, "comment", "") or "")

        destinations = []
        try:
            for sheet_name, ref in dn.destinations:
                destinations.append({"sheet": str(sheet_name), "ref": str(ref)})
        except (TypeError, ValueError, AttributeError):
            pass

        dedup_key = (
            name,
            str(local_sheet_id),
            hidden,
            attr_text,
            json.dumps(destinations, ensure_ascii=False, sort_keys=True),
        )
        if dedup_key in seen:
            continue
        seen.add(dedup_key)

        names.append(
            {
                "name": name,
                "local_sheet_id": local_sheet_id,
                "hidden": hidden,
                "attr_text": attr_text,
                "is_formula": bool(attr_text.startswith("=")),
                "comment": comment,
                "destinations": destinations,
            }
        )

    names.sort(key=lambda x: (x.get("name", ""), x.get("local_sheet_id") or -1))
    return names
