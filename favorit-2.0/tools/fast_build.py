#!/usr/bin/env python3
"""Fast PDF build/preflight for the existing Favorit JS generator. Does not edit it."""
import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time


def run(args, timeout, env=None):
    result = subprocess.run(args, env=env, capture_output=True, text=True, timeout=timeout)
    if result.returncode:
        raise RuntimeError(f"{args[0]}: {result.stderr.strip() or result.stdout.strip()}")
    return result


def check_pdf(path, claim, vin, damage_names):
    import fitz
    pdf = fitz.open(path)
    if len(pdf) < 4:
        raise RuntimeError(f"В PDF всего {len(pdf)} стр., ожидается минимум 4")
    pages = [p.get_text() for p in pdf]
    faults = []
    for i, text in enumerate(pages, 1):
        if claim and claim not in text:
            faults.append(f"номер убытка отсутствует на листе {i}")
        if not pdf[i-1].get_images(full=True):
            faults.append(f"нет изображения факсимиле на листе {i}")
    # VIN is printed as 17 separate physical cells in the existing DOCX form.
    compact = re.sub(r"\s+", "", pages[0]).upper()
    if vin and vin.upper() not in compact:
        faults.append("VIN не найден на первом листе")
    if "Повреждения, указанные" not in pages[2]:
        faults.append("таблица повреждений начинается не на третьем листе")
    if "Повреждения, указанные" in pages[0] or "Повреждения, указанные" in pages[1]:
        faults.append("таблица повреждений попала на первый или второй лист")
    damaged = " ".join(pages[2:-1]).lower()
    damaged = re.sub(r"\s+", " ", damaged)
    for name in damage_names:
        name = re.sub(r"\s+", " ", name.strip().lower())
        if name and name not in damaged:
            faults.append(f"нет позиции повреждения: {name}")
    if faults:
        raise RuntimeError("PDF не прошёл проверку: " + "; ".join(faults))
    return len(pdf)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workdir", type=Path, required=True, help="Папка с заполненным generator.js и штатным DOCX")
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--claim", default="")
    ap.add_argument("--vin", default="")
    ap.add_argument("--damage", action="append", default=[], help="Название позиции для контроля; можно повторять")
    ap.add_argument("--test", action="store_true", help="Разрешить неполные тестовые реквизиты")
    args = ap.parse_args()
    src = args.workdir.resolve()
    generator = src / "generator.js"
    if not generator.is_file():
        ap.error(f"Не найден {generator}")
    if args.vin and len(args.vin) != 17:
        ap.error("VIN должен содержать 17 символов")
    t0 = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="favorit_fast_") as td:
        temp = Path(td)
        docx = temp / "act.docx"
        env = dict(os.environ, AKT_OUT=str(docx))
        if args.test:
            env["AKT_TEST"] = "1"
        run(["node", str(generator)], 25, env=env)
        t1 = time.monotonic()
        profile = (temp / "lo-profile").as_uri()
        run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(docx)], 35)
        t2 = time.monotonic()
        pdf = temp / "act.pdf"
        if not pdf.exists():
            raise RuntimeError("LibreOffice не создал PDF")
        pages = check_pdf(pdf, args.claim, args.vin, args.damage)
        dest = args.output.resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(pdf, dest)
        t3 = time.monotonic()
    print(json.dumps({"ok": True, "pdf": str(dest), "pages": pages, "seconds": {"docx": round(t1-t0, 2), "pdf": round(t2-t1, 2), "check_copy": round(t3-t2, 2), "total": round(t3-t0, 2)}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
