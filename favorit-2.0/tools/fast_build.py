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


def read_js_data(generator):
    """Read only JS input block in isolated Node VM without executing the generator."""
    js = r"""
const fs = require('fs'), vm = require('vm');
const source = fs.readFileSync(process.argv[1], 'utf8');
const marker = '// Конец блока данных.';
if (source.split(marker).length !== 2) throw Error('Граница блока данных отсутствует');
const input = source.split(marker)[0];
process.stdout.write(vm.runInNewContext(
  input + '\nJSON.stringify({D,damages})', {JSON}, {timeout:1000}));
"""
    return json.loads(run(["node", "-e", js, str(generator)], 5).stdout)


def norm(text):
    return re.sub(r"\\s+", "", text).casefold().replace("ё", "е")


def check_pdf(path, claim, vin, damage_names, expected_count=None):
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
    if expected_count is not None:
        # In the autonomous form each damage row's index is printed on its own line.
        numbers = [int(n) for n in re.findall(
            r"(?m)^\\s*(\\d{1,3})\\s*$", "\\n".join(pages[2:-1]))]
        if numbers != list(range(1, expected_count + 1)):
            faults.append(f"Строки/нумерация ДТП: {numbers}, ожидалось 1..{expected_count}")
    if faults:
        raise RuntimeError("PDF не прошёл проверку: " + "; ".join(faults))
    return len(pdf)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workdir", type=Path, help="Папка с заполненным generator.js и штатным DOCX")
    ap.add_argument("--generator", type=Path, help="Точный путь готового JS с блоком D (ускоренный повтор)")
    ap.add_argument("--core-fnv", type=int, help="Контрольный FNV-1a суффикса JS из GitHub main")
    ap.add_argument("--output", type=Path, required=True)
    ap.add_argument("--claim", default="")
    ap.add_argument("--vin", default="")
    ap.add_argument("--damage", action="append", default=[], help="Название позиции для контроля; можно повторять")
    ap.add_argument("--approved-damages", type=Path, help="Отдельный согласованный список деталей: UTF-8, строка = деталь")
    ap.add_argument("--test", action="store_true", help="Разрешить неполные тестовые реквизиты")
    args = ap.parse_args()
    if not args.generator and not args.workdir:
        ap.error("Нужен --generator или --workdir")
    generator = args.generator.resolve() if args.generator else args.workdir.resolve() / "generator.js"
    if not generator.is_file():
        ap.error(f"Не найден {generator}")
    data = read_js_data(generator)
    js_claim = str(data["D"].get("claimNo") or "")
    js_vin = str(data["D"].get("vin") or "")
    if args.claim and args.claim != js_claim:
        ap.error("Номер убытка аргумента не соответствует JS")
    if args.vin and args.vin != js_vin:
        ap.error("VIN аргумента не соответствует JS")
    claim = args.claim or js_claim
    vin = args.vin or js_vin
    if vin and len(vin) != 17:
        ap.error("VIN должен содержать 17 символов")
    if not args.test and (not claim or not vin):
        ap.error("В JS не заполнены номер убытка / VIN")
    names = [d["detail"] for d in data["damages"]]
    if not args.test and not names:
        ap.error("В JS нет списка ДТП")
    for name in args.damage:
        if norm(name) not in {norm(n) for n in names}:
            ap.error(f"Переданное повреждение отсутствует в JS: {name}")
    if args.approved_damages:
        approved = [line.strip() for line in args.approved_damages.read_text(encoding="utf-8-sig").splitlines()
                    if line.strip() and not line.lstrip().startswith("#")]
        if [norm(x) for x in approved] != [norm(x) for x in names]:
            ap.error("Список ДТП в JS отличается от согласованного; сборка остановлена")
    if args.core_fnv is None and not args.test:
        ap.error("Нужен --core-fnv из актуального GitHub-генератора")
    if args.core_fnv is not None:
        source = generator.read_text(encoding="utf-8")
        marker = "// Конец блока данных."
        if source.count(marker) != 1:
            ap.error("Граница блока D не найдена или неоднозначна")
        core = source[source.index(marker):]
        h = 2166136261
        for ch in core:
            h = ((h ^ ord(ch)) * 16777619) & 0xffffffff
        if h != args.core_fnv:
            ap.error("Генератор не соответствует актуальному GitHub main; сборка заблокирована")
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
        pages = check_pdf(pdf, claim, vin, names, expected_count=len(names))
        dest = args.output.resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(pdf, dest)
        t3 = time.monotonic()
    print(json.dumps({"ok": True, "pdf": str(dest), "pages": pages, "seconds": {"docx": round(t1-t0, 2), "pdf": round(t2-t1, 2), "check_copy": round(t3-t2, 2), "total": round(t3-t0, 2)}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
