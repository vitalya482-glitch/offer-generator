from __future__ import annotations

import argparse
from pathlib import Path

from brands.registry import get_brand_module
from core.models import OfferContext
from gui.main_window import run_gui


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="SAM Offer Generator")
    parser.add_argument("--gui", action="store_true", help="Запустить графический интерфейс")
    parser.add_argument("--brand", default="Stulz", choices=["Stulz", "Riello", "DC Eltek", "Generator"])
    parser.add_argument("--project-dir", default=".", help="Папка проекта на сервере")
    parser.add_argument("--template", help="Word-шаблон .docx")
    parser.add_argument("--calc", help="Excel-калькуляция .xlsx/.xlsm")
    parser.add_argument("--output-dir", help="Папка сохранения результата")
    parser.add_argument("--client", default="ТОО Example")
    parser.add_argument("--sheet", default=None)
    args = parser.parse_args(argv)

    if args.gui or (not args.template and not args.calc):
        run_gui()
        return 0

    project_dir = Path(args.project_dir).resolve()
    context = OfferContext(
        brand=args.brand,
        project_dir=project_dir,
        template_path=Path(args.template).resolve(),
        calc_path=Path(args.calc).resolve(),
        output_dir=Path(args.output_dir).resolve() if args.output_dir else project_dir,
        client_name=args.client,
        sheet_name=args.sheet,
        pdf_dir=project_dir if project_dir.exists() else None,
    )
    module = get_brand_module(args.brand)
    out = module.make_offer(context)
    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
