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

## Tests

```console
uv run pytest
```

## Build a Windows executable

```console
package.bat
```

This bumps the build number and creates `dist/pol-tax` with PyInstaller.
The Inno Setup scripts in `installer/` package it into an installer.
