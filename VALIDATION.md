# Validation

Validation is limited to source correctness on invented data. No original XLSX/XML, Excel application, browser, production service or network API is used.

Baseline: two focused regression checks against the copied legacy source failed as expected: import attempted to open Tk; a marking containing reserved XML characters did not survive one serialize/parse cycle unchanged.

Final command (from this directory):

```powershell
python -B -X utf8 -m unittest discover -s tests -v
```

Result: **18 tests passed**, exit code 0, on 2026-10-29. Environment: Windows, CPython 3.12.10, pandas 3.0.3, openpyxl 3.1.5. Test runner reported 0.836 seconds; this is a single observed run, not a performance guarantee.

Tests cover actual pandas/openpyxl XLSX I/O, ElementTree Windows-1251 output, input immutability, leading zeroes, reserved characters/literal entities, import safety, missing mappings, conflicting/duplicate IDs, idempotence, invalid XML controls, output overwrite safeguards and cancellation. Two additional regression checks first failed on the new file workflow: pandas interpreted a literal `NA` code as missing; carriage returns were accepted despite XML newline normalization. They passed after disabling default NA conversion and rejecting carriage returns. Temporary library files also stay inside the synthetic test directory.

Not verified: real exchange schemas/services, digital signatures, business code allocation, GS1 validity, hostile input parsing, GUI appearance, crash recovery, other Python/package versions. Dependencies are ranges, not a lockfile or a claim of latest compatibility.
