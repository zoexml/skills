from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
WORKBOOK_GENERATOR = SCRIPT_DIR / "generate_store_diff_workbook.py"
PPT_GENERATOR = SCRIPT_DIR / "generate_store_diff_ppt.py"


def load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate store category role analysis workbook and PPT.")
    parser.add_argument("--store-xlsx", required=True, type=Path)
    parser.add_argument("--enterprise-xlsx", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--tag", required=True)
    args = parser.parse_args()

    store_xlsx = args.store_xlsx.expanduser().resolve()
    enterprise_xlsx = args.enterprise_xlsx.expanduser().resolve()
    output_dir = args.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not store_xlsx.exists():
        raise FileNotFoundError(store_xlsx)
    if not enterprise_xlsx.exists():
        raise FileNotFoundError(enterprise_xlsx)

    detail_xlsx = output_dir / f"各门店代表性差异品类_{args.tag}.xlsx"
    pptx = output_dir / f"各门店代表性差异品类分析报告_{args.tag}.pptx"

    workbook = load_module(WORKBOOK_GENERATOR, "store_category_role_workbook")
    workbook.STORE_ROLE_XLSX = store_xlsx
    workbook.ENTERPRISE_XLSX = enterprise_xlsx
    workbook.OUTPUT_XLSX = detail_xlsx
    workbook.BASE_DIR = output_dir
    workbook.main()

    ppt = load_module(PPT_GENERATOR, "store_category_role_ppt")
    ppt.BASE_DIR = output_dir
    ppt.INPUT_XLSX = detail_xlsx
    ppt.OUTPUT_PPTX = pptx
    ppt.STORE_ROLE_XLSX = store_xlsx
    ppt.ENTERPRISE_ROLE_XLSX = enterprise_xlsx
    ppt.main()

    print(detail_xlsx)
    print(pptx)


if __name__ == "__main__":
    main()
