# Pol-Tax

Desktop app (Windows, PySide6) for reading Polish electronic financial statements
(*e-Sprawozdania finansowe*) filed with the KRS. It:

- opens a statement XML and downloads the Ministry of Finance XSD schemas to get human-readable labels,
- shows each section (Wprowadzenie, Aktywa, Pasywa, RZiS, zmiany w kapitale, przepływy pieniężne,
  informacja dodatkowa o podatku dochodowym) as a table and exports it to Excel,
- extracts PDF attachments embedded in the XML,
- runs a "Sumowanie" consistency check (subtotals, assets = liabilities, cross-statement links).

## Setup

Requires [uv](https://docs.astral.sh/uv/). On Windows:

```console
winget install --id astral-sh.uv -e
```

Then, in the repo:

```console
uv sync
uv run python main.py
```

## Using the app

- Open a statement with **Otwórz XML** (Ctrl+O), by dragging the file onto the window, from **Ostatnio otwierane**,
  or with *Open with → Pol-Tax* in Explorer.
- Statements can be downloaded from the
  [Repozytorium Dokumentów Finansowych](https://ekrs.ms.gov.pl/rdf/pd/search_df) (search by KRS number).
- The consistency check runs automatically after opening. Failed checks are listed under the table with the
  actual and expected value, the difference and the rule; click one to jump to the position. Checked rows are
  green or red, hover a row to see its checks. **Raport w Excelu** opens the full report with a
  *Niezgodności* sheet.
- **Otwórz w Excelu** (Ctrl+E) opens the displayed part of the statement in Excel.

## Command line

Convert many statements at once without the GUI (`--check` adds the consistency check workbook):

```console
uv run python cli.py sprawozdanie1.xml sprawozdanie2.xml -o wyniki --check
```

Settings, the XSD cache, Excel files and extracted attachments live in `Documents\pol-tax`.
The consistency rules are defined in [`XML/rules.py`](XML/rules.py).

## Tests

```console
uv run pytest
```

## Build a Windows executable

```console
uv run python scripts/build.py
```

This creates the app in `dist/pol-tax/` with PyInstaller (configured in `pol-tax.spec`) and, if
[Inno Setup 6](https://jrsoftware.org/isinfo.php) is installed (`winget install JRSoftware.InnoSetup`),
the installer `dist/Pol-Tax_<version>.exe`. Use `--no-installer` to skip the installer.

The build never changes tracked files. The version lives only in `pyproject.toml`.

## Releasing

1. Bump the version: `uv version --bump patch` (or `minor` / `major`)
2. Commit, tag and push:

   ```console
   git commit -am "Release 0.2.0"
   git tag v0.2.0
   git push --follow-tags
   ```

The `Release` GitHub Actions workflow checks that the tag matches `pyproject.toml`, runs the tests, builds the
installer and attaches it to a GitHub release. The `CI` workflow runs the tests and a build on every push.
