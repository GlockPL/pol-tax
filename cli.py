"""
Batch conversion of e-financial statements without the GUI:

    uv run python cli.py sprawozdanie1.xml sprawozdanie2.xml -o wyniki --check
    uv run python cli.py sprawozdanie.xml -o wyniki --workbook   # one workbook per statement
"""
import argparse
import logging
import sys
from pathlib import Path

from Settings import Settings
from XML.export import default_workbook_name, export_statement, export_summation, export_workbook


def main(argv=None):
    arg_parser = argparse.ArgumentParser(description="Konwersja e-sprawozdań finansowych (XML z KRS) na pliki Excel")
    arg_parser.add_argument("files", nargs="+", type=Path, help="Pliki XML do przetworzenia")
    arg_parser.add_argument("-o", "--output", type=Path, help="Folder na pliki Excel (domyślnie z ustawień programu)")
    arg_parser.add_argument("-c", "--check", action="store_true", help="Wykonaj również sprawdzenie sum")
    arg_parser.add_argument("-w", "--workbook", action="store_true",
                            help="Zapisz całe sprawozdanie w jednym pliku Excel (z kontrolą sum)")
    arg_parser.add_argument("-v", "--verbose", action="store_true", help="Pokaż szczegółowe logi")
    args = arg_parser.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.WARNING,
                        format='%(levelname)s - %(message)s')
    settings = Settings()
    excel_dir = args.output or Path(settings.get_excel_dir())

    failed = 0
    for xml_path in args.files:
        try:
            result = export_statement(xml_path, settings, excel_dir)
            outputs = [path for path, _, _ in result.sections.values()]
            summary = ""
            summation = None
            if args.check or args.workbook:
                summation = export_summation(result.parser, excel_dir, result.prefix)
                outputs.append(summation.path)
                summary = f" (sprawdzono {len(summation.results)} reguł, niezgodności: {len(summation.failures)})"
            if args.workbook:
                outputs.append(export_workbook(result, excel_dir / default_workbook_name(result),
                                               settings.get_keywords(), summation))
            print(f"{xml_path}:{summary}")
            for output in outputs + result.attachments:
                print(f"  {output}")
        except Exception as e:
            failed += 1
            logging.debug("Processing failed", exc_info=True)
            print(f"{xml_path}: BŁĄD - {e}", file=sys.stderr)

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
