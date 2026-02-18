#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор PDF из данных CSV/JSON и HTML-шаблонов.
Использует WeasyPrint для создания PDF с поддержкой кириллицы.
"""

import csv
import json
import sys

# Поддержка кириллицы в консоли Windows (если доступно)
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass
import os
import platform
import subprocess
from pathlib import Path

from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

# Директории относительно скрипта
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

# Расширения
CSV_EXT = (".csv",)
JSON_EXT = (".json",)
HTML_EXT = (".html",)


def ensure_dirs():
    """Создаёт необходимые директории."""
    DATA_DIR.mkdir(exist_ok=True)
    TEMPLATES_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def get_data_files():
    """Возвращает списки CSV и JSON файлов из /data."""
    csv_files = []
    json_files = []
    if not DATA_DIR.exists():
        return csv_files, json_files
    for f in sorted(DATA_DIR.iterdir()):
        if f.is_file() and f.suffix.lower() in CSV_EXT:
            csv_files.append(f)
        elif f.is_file() and f.suffix.lower() in JSON_EXT:
            json_files.append(f)
    return csv_files, json_files


def get_templates():
    """Возвращает список HTML-шаблонов из /templates."""
    if not TEMPLATES_DIR.exists():
        return []
    return sorted(
        f for f in TEMPLATES_DIR.iterdir()
        if f.is_file() and f.suffix.lower() in HTML_EXT
    )


def load_csv(path: Path):
    """Загружает CSV через pandas (если есть) или csv."""
    if HAS_PANDAS:
        df = pd.read_csv(path, encoding="utf-8")
        return df.to_dict(orient="records")
    with open(path, encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        return list(reader)


def load_json(path: Path):
    """Загружает JSON."""
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_invoices_from_data(data, file_path: Path):
    """
    Преобразует загруженные данные в структуру {invoice_id: [items]}.
    Поддерживает CSV (invoice_id, product, price, qty) и JSON (список с invoice_id и items).
    """
    invoices = {}
    ext = file_path.suffix.lower()

    if ext == ".json":
        if isinstance(data, list):
            for inv in data:
                iid = inv.get("invoice_id") or inv.get("id") or str(len(invoices))
                items = inv.get("items", [inv])
                invoices[iid] = [
                    {
                        "product": it.get("product", it.get("name", "")),
                        "price": it.get("price", 0),
                        "qty": it.get("qty", it.get("quantity", 1)),
                    }
                    for it in items
                ]
        else:
            invoices["default"] = [{"product": "-", "price": 0, "qty": 1}]
        return invoices

    # CSV: ожидаем колонки invoice_id (или id), product, price, qty
    for row in data:
        iid = str(row.get("invoice_id") or row.get("id") or "")
        if not iid:
            continue
        item = {
            "product": str(row.get("product", row.get("name", ""))),
            "price": row.get("price", 0),
            "qty": int(row.get("qty", row.get("quantity", 1))),
        }
        if iid not in invoices:
            invoices[iid] = []
        invoices[iid].append(item)

    return invoices


def render_html(template_path: Path, context: dict) -> str:
    """Рендерит HTML из Jinja2-шаблона."""
    env = Environment(
        loader=FileSystemLoader(str(template_path.parent)),
        autoescape=True,
    )
    template = env.get_template(template_path.name)
    return template.render(**context)


def generate_pdf(html_content: str, output_path: Path):
    """Генерирует PDF через WeasyPrint с поддержкой кириллицы."""
    html = HTML(string=html_content, base_url=str(BASE_DIR))
    # DejaVu Sans / Liberation / Arial поддерживают кириллицу на Windows и macOS
    css = CSS(string="""
        body { font-family: 'DejaVu Sans', 'Liberation Sans', 'Arial', sans-serif; }
    """)
    html.write_pdf(output_path, stylesheets=[css])


def open_pdf(path: Path):
    """Открывает PDF в системной программе (Windows / macOS / Linux)."""
    path_str = str(path.resolve())
    system = platform.system()
    if system == "Windows":
        os.startfile(path_str)
    elif system == "Darwin":
        subprocess.run(["open", path_str], check=False)
    else:
        subprocess.run(["xdg-open", path_str], check=False)


def print_menu(title: str, items: list, extra_info=None):
    """Печатает нумерованное меню."""
    print()
    print(f"  {title}")
    print("  " + "-" * 40)
    for i, item in enumerate(items, 1):
        info = extra_info(i, item) if extra_info else ""
        print(f"    {i}. {item} {info}")
    print("  " + "-" * 40)


def select_option(prompt: str, max_num: int) -> int:
    """Запрашивает выбор пользователя и возвращает индекс (0-based)."""
    while True:
        try:
            s = input(prompt).strip()
            n = int(s)
            if 1 <= n <= max_num:
                return n - 1
        except ValueError:
            pass
        print("  Введите число от 1 до", max_num)


def main():
    ensure_dirs()

    csv_files, json_files = get_data_files()
    data_files = csv_files + json_files
    templates = get_templates()

    if not data_files:
        print("\n  В директории /data нет CSV или JSON файлов.")
        print("  Добавьте файлы и запустите скрипт снова.")
        sys.exit(1)

    if not templates:
        print("\n  В директории /templates нет HTML-шаблонов.")
        print("  Добавьте шаблоны и запустите скрипт снова.")
        sys.exit(1)

    # Вывод доступных файлов
    print("\n" + "=" * 50)
    print("  ГЕНЕРАТОР PDF")
    print("=" * 50)
    print_menu(
        "Доступные файлы с данными:",
        [f.name for f in data_files],
    )
    idx_data = select_option("  Выберите файл данных (номер): ", len(data_files))
    data_path = data_files[idx_data]

    # Вывод доступных шаблонов
    print_menu(
        "Доступные HTML-шаблоны:",
        [t.name for t in templates],
    )
    idx_tpl = select_option("  Выберите шаблон (номер): ", len(templates))
    template_path = templates[idx_tpl]

    # Загрузка данных
    ext = data_path.suffix.lower()
    if ext == ".csv":
        raw = load_csv(data_path)
    else:
        raw = load_json(data_path)

    invoices = parse_invoices_from_data(raw, data_path)
    if not invoices:
        print("\n  В выбранном файле нет данных с invoice_id.")
        sys.exit(1)

    # Выбор invoice
    inv_ids = sorted(invoices.keys())
    print_menu(
        "Доступные счета (invoice id):",
        inv_ids,
    )
    idx_inv = select_option("  Выберите счёт (номер): ", len(inv_ids))
    invoice_id = inv_ids[idx_inv]
    items = invoices[invoice_id]

    # Генерация PDF
    context = {"invoice_id": invoice_id, "items": items}
    html_content = render_html(template_path, context)
    safe_id = "".join(c if c.isalnum() or c in "-_" else "_" for c in invoice_id)
    output_path = OUTPUT_DIR / f"invoice_{safe_id}.pdf"
    generate_pdf(html_content, output_path)

    print(f"\n  PDF сохранён: {output_path}")
    open_pdf(output_path)
    print("  PDF открыт в системной программе.\n")


if __name__ == "__main__":
    main()
