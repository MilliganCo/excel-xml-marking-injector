# Excel XML marking injector

Local Python utility that joins an XLSX list of marking codes to an XLSX barcode/GTIN reference, then adds `КИЗ` elements to a separate XML file. This source snapshot contains no original business documents, marking codes, credentials or output files.

## Run

Use Python 3.10+ with Tkinter if you want file dialogs. Install dependencies in a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python excelxml.py --markings input/markings.xlsx --reference input/reference.xlsx --xml input/document.xml --out output/marked.xml
```

Create the output directory first. Run `python excelxml.py` without file arguments to choose the four paths in dialogs. Supplying only some file arguments is an error. Importing `excelxml` opens no dialogs and reads no user files. Existing output files require `--overwrite`; input files cannot be chosen as output. Default XML output encoding is Windows-1251; `--encoding utf-8` is available. The console prints counts, not marking codes or product identifiers.

## Input contract

- First worksheet of markings XLSX: exact columns `GTIN`, `Код маркировки`.
- First worksheet of reference XLSX: exact columns `Баркод`, `ГТИН`.
- Keep identifiers in **text cells** in Excel. XLSX import requests strings; it cannot recover leading zeroes or digits already discarded by Excel numeric formatting. Barcodes match as stripped text. GTIN input must contain 1–14 ASCII digits; whitespace is trimmed before padding to 14 digits. This is normalization, not GS1 check-digit validation.
- XML uses non-namespaced `СведТов` elements containing one `ДопСведТов` with a `КодТов` attribute. Codes are appended under direct `НомСредИдентТов/КИЗ` children. An unsupported document with no such product elements is rejected.

Reserved XML characters are assigned as raw text; ElementTree serializes them once. Literal text such as `&amp;` remains literal text after parsing the output. XLSX strings such as `NA` are preserved instead of being treated as missing values. XML 1.0 forbidden controls, including a raw GS separator (`U+001D`), are rejected; carriage returns are also rejected because XML parsers normalize them. The program does not guess a replacement encoding.

Repeated marking rows for one GTIN are deduplicated. Identical reference rows are accepted; conflicting nonempty GTINs for one barcode fail before changes. Missing identifiers/mappings are skipped with separate counts. A code assigned to different GTINs, duplicate existing XML codes, multiple containers, or multiple XML lines needing the same GTIN are rejected before appending: allocation among quantities/line items is unspecified. Running again with the same valid inputs does not duplicate existing codes on that product.

## Limits and private data

Use only trusted, local documents and inspect the separate output before any business exchange. This prototype does not validate an XSD, quantities, code authenticity, digital signatures, current marketplace/tax service compatibility, or document acceptance. ElementTree may rewrite formatting, namespace prefixes and declarations and omit comments; editing signed XML invalidates its signature. Namespaced documents need a separate schema-aware implementation. Output writes are not crash-atomic. A file write failure can leave a partial output, including an explicitly overwritten file; keep backups.

No network/API calls or configuration secrets are required. Keep actual XLSX/XML files and logs private; the ignore rules are safeguards, not a data audit. This snapshot intentionally supplies no new license or third-party attribution claims beyond its declared Python dependencies.

## Offline checks

```powershell
python -B -X utf8 -m unittest discover -s tests -v
```

Tests use invented codes and temporary XLSX/XML files under ignored `work/`; they do not open Excel, dialogs, browsers, services or user data. See `VALIDATION.md` for the recorded environment and limits.
