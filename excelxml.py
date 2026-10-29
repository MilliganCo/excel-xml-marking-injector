"""Insert marking codes from two XLSX tables into a separate XML document."""

import argparse
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET

import pandas as pd


def _text(value):
    return "" if pd.isna(value) else str(value)


def normalize_gtin(value):
    """Pad a textual identifier to 14 digits; this is not check-digit validation."""
    text = _text(value).strip()
    if not re.fullmatch(r"[0-9]{1,14}", text):
        raise ValueError("GTIN must contain between 1 and 14 ASCII digits")
    return text.zfill(14)


def _require_columns(frame, columns, name):
    if not set(columns).issubset(frame.columns):
        raise ValueError(f"{name}: required columns are {', '.join(columns)}")


def build_markings(frame):
    _require_columns(frame, ("GTIN", "Код маркировки"), "markings")
    result, owners = {}, {}
    for _, row in frame.iterrows():
        raw_gtin, code = _text(row["GTIN"]), _text(row["Код маркировки"])
        if not raw_gtin.strip() and not code.strip():
            continue
        gtin = normalize_gtin(raw_gtin)
        if not code.strip():
            raise ValueError("A marking row has an empty code")
        if "\r" in code:
            raise ValueError("A marking code contains a carriage return, normalized by XML parsers")
        if any(not (c in "\t\n\r" or 0x20 <= ord(c) <= 0xD7FF
                    or 0xE000 <= ord(c) <= 0xFFFD
                    or 0x10000 <= ord(c) <= 0x10FFFF) for c in code):
            raise ValueError("A marking code contains a character forbidden in XML 1.0")
        if code in owners and owners[code] != gtin:
            raise ValueError("A marking code is assigned to multiple GTINs")
        owners[code] = gtin
        codes = result.setdefault(gtin, [])
        if code not in codes:
            codes.append(code)
    return result


def build_reference(frame):
    _require_columns(frame, ("Баркод", "ГТИН"), "reference")
    result = {}
    for _, row in frame.iterrows():
        barcode, raw_gtin = _text(row["Баркод"]).strip(), _text(row["ГТИН"]).strip()
        if not barcode:
            continue
        gtin = normalize_gtin(raw_gtin) if raw_gtin else None
        previous = result.get(barcode)
        if previous is not None and gtin is not None and previous != gtin:
            raise ValueError("A reference barcode maps to multiple GTINs")
        result[barcode] = previous or gtin
    return result


def inject_markings(root, markings_frame, reference_frame):
    """Preflight a non-namespaced document, then append raw text exactly once.

    Missing mappings are counted and skipped. Ambiguous allocation fails before
    mutation: this utility cannot divide physical codes among repeated line items.
    """
    markings = build_markings(markings_frame)
    reference = build_reference(reference_frame)
    products = root.findall(".//СведТов")
    if not products:
        raise ValueError("No supported non-namespaced СведТов elements found")
    counts = {name: 0 for name in (
        "products", "matched_products", "added_codes", "existing_codes",
        "missing_additional_info", "missing_barcode", "missing_reference",
        "missing_reference_gtin", "missing_markings")}
    counts["products"] = len(products)
    existing_owners = {}
    for product in products:
        for element in product.findall(".//НомСредИдентТов/КИЗ"):
            code = element.text or ""
            if code in existing_owners:
                raise ValueError("The XML contains a duplicate existing marking code")
            existing_owners[code] = product
    plans, allocated_gtins = [], set()
    for product in products:
        details = product.findall(".//ДопСведТов")
        if not details:
            counts["missing_additional_info"] += 1
            continue
        if len(details) != 1:
            raise ValueError("A product has multiple ДопСведТов elements")
        detail = details[0]
        barcode = (detail.get("КодТов") or "").strip()
        if not barcode:
            counts["missing_barcode"] += 1
            continue
        if barcode not in reference:
            counts["missing_reference"] += 1
            continue
        gtin = reference[barcode]
        if gtin is None:
            counts["missing_reference_gtin"] += 1
            continue
        if gtin not in markings:
            counts["missing_markings"] += 1
            continue
        if gtin in allocated_gtins:
            raise ValueError("Multiple XML products need the same GTIN; code allocation is ambiguous")
        allocated_gtins.add(gtin)
        containers = detail.findall("НомСредИдентТов")
        if len(containers) > 1:
            raise ValueError("A product has multiple НомСредИдентТов containers")
        to_add = []
        for code in markings[gtin]:
            owner = existing_owners.get(code)
            if owner is not None and owner is not product:
                raise ValueError("A marking code already belongs to another XML product")
            if owner is product:
                counts["existing_codes"] += 1
            else:
                to_add.append(code)
        plans.append((detail, containers[0] if containers else None, to_add))
        counts["matched_products"] += 1
    for detail, container, codes in plans:
        if codes and container is None:
            container = ET.SubElement(detail, "НомСредИдентТов")
        for code in codes:
            # ElementTree escapes reserved characters when serializing XML.
            ET.SubElement(container, "КИЗ").text = code
            counts["added_codes"] += 1
    return counts


def process_files(markings_path, reference_path, xml_path, output_path,
                  *, overwrite=False, encoding="windows-1251"):
    inputs = [Path(p).resolve() for p in (markings_path, reference_path, xml_path)]
    output = Path(output_path).resolve()
    if any(output == source or (output.exists() and output.samefile(source))
           for source in inputs):
        raise ValueError("Output must be separate from every input file")
    if output.exists() and not overwrite:
        raise FileExistsError("Output already exists; use --overwrite explicitly")
    # Text cells preserve leading zeroes; Excel cannot recover digits it already discarded.
    markings = pd.read_excel(inputs[0], dtype=str, keep_default_na=False)
    reference = pd.read_excel(inputs[1], dtype=str, keep_default_na=False)
    tree = ET.parse(inputs[2])
    counts = inject_markings(tree.getroot(), markings, reference)
    # Serialize fully before opening the output, so encoding errors leave it untouched.
    payload = ET.tostring(tree.getroot(), encoding=encoding, xml_declaration=True)
    with output.open("wb" if overwrite else "xb") as handle:
        handle.write(payload)
    return counts


def select_file(title, *, save=False):
    """Create a short-lived dialog only when explicitly running the program."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    try:
        if save:
            return filedialog.asksaveasfilename(parent=root, title=title,
                defaultextension=".xml", filetypes=[("XML files", "*.xml")])
        return filedialog.askopenfilename(parent=root, title=title,
            filetypes=[("Excel / XML files", "*.xlsx *.xml")])
    finally:
        root.destroy()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--markings", help="XLSX: GTIN, Код маркировки")
    parser.add_argument("--reference", help="XLSX: Баркод, ГТИН")
    parser.add_argument("--xml", help="Input XML")
    parser.add_argument("--out", help="Separate output XML")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--encoding", default="windows-1251")
    args = parser.parse_args(argv)
    try:
        supplied = (args.markings, args.reference, args.xml, args.out)
        if any(supplied) and not all(supplied):
            parser.error("provide all four file arguments, or none to use dialogs")
        if not any(supplied):
            supplied = (select_file("Файл с маркировками"),
                        select_file("Справочник артикулов"),
                        select_file("Исходный XML"),
                        select_file("Сохранить отдельный XML", save=True))
            if not all(supplied):
                print("Cancelled; no output written.", file=sys.stderr)
                return 2
        counts = process_files(*supplied, overwrite=args.overwrite, encoding=args.encoding)
    except (OSError, ValueError, ET.ParseError, LookupError) as exc:
        # Do not print raw table contents, identifiers or marking codes.
        print(f"Failed: {type(exc).__name__}. Check inputs and output settings.", file=sys.stderr)
        return 1
    print("; ".join(f"{key}={value}" for key, value in counts.items()))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
