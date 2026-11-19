#!/usr/bin/env python3
"""Fast PDF build/preflight for the existing Favorit JS generator. Does not edit it."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import time


DAMAGE_ROWS_PER_SHEET = 17  # Capacity of Favorit_AKT_autonomous.js.
CONCLUSION_HEADING = "Характер описанных повреждений"


EXPERT_SOURCES = ("EXPERT_KNOWLEDGE.md", "CASES.md", "knowledge/CORRECTIONS.md")
EXPERT_CHECKS = (
    "professional_wording",
    "damage_completeness",
    "operational_defects",
    "evidence_classification",
)


def sha256_file(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_expert_review(report_path, generator):
    """Validate the recorded expert review for these exact build inputs."""
    if report_path is None:
        raise ValueError("Реальный акт: обязателен --expert-review с отчётом экспертной проверки")
    try:
        report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"Не удалось прочитать JSON-отчёт экспертной проверки: {exc}") from exc

    if not isinstance(report, dict):
        raise ValueError("Отчёт экспертной проверки должен быть JSON-объектом")
    if report.get("open_errors") != []:
        raise ValueError("В экспертной проверке есть незакрытые ошибки или отсутствует массив open_errors")

    checks = report.get("checks")
    if not isinstance(checks, dict):
        raise ValueError("В отчёте отсутствуют результаты экспертной проверки (checks)")
    for name in EXPERT_CHECKS:
        note = checks.get(name)
        if not isinstance(note, str) or not note.strip():
            raise ValueError(f"В экспертном отчёте нет результата проверки: {name}")

    try:
        generator_hash = sha256_file(generator)
        project_root = Path(__file__).resolve().parents[1]
        sources = {name: sha256_file(project_root / name) for name in EXPERT_SOURCES}
    except OSError as exc:
        raise ValueError(f"Не удалось сверить входной JS или экспертную базу: {exc}") from exc

    if report.get("generator_sha256") != generator_hash:
        raise ValueError("Экспертный отчёт относится к другой редакции входного JS")
    if report.get("sources") != sources:
        raise ValueError("Экспертный отчёт не соответствует файлам экспертной базы сборщика")
    return generator_hash


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
  input + '\nJSON.stringify({D,damages,CHECKS,SIGN_EXPERT})', {JSON}, {timeout:1000}));
"""
    return json.loads(run(["node", "-e", js, str(generator)], 5).stdout)



def norm(text):
    return re.sub(r"\s+", "", text).casefold().replace("ё", "е")


def has_checked(text, label, selected):
    mark = "☒" if selected else "☐"
    return re.search(re.escape(mark) + r"\s*" + re.escape(label), text, re.I) is not None


def check_pair(text, heading, expected, max_lookahead=650):
    idx = text.find(heading)
    if idx < 0:
        return False
    match = re.search(r"([☒☐])\s*Да\s*([☒☐])\s*Нет", text[idx:idx+max_lookahead])
    if not match:
        return False
    yes, no = match.groups()
    if expected is True:
        return yes == "☒" and no == "☐"
    if expected is False:
        return yes == "☐" and no == "☒"
    return yes == "☐" and no == "☐"


def check_pdf(path, claim, vin, damages, expected_count=None, checks=None, data_fields=None):
    import fitz
    pdf = fitz.open(path)
    damage_count = len(damages) if expected_count is None else expected_count
    damage_sheets = max(1, (damage_count + DAMAGE_ROWS_PER_SHEET - 1) // DAMAGE_ROWS_PER_SHEET)
    expected_pages = 3 + damage_sheets
    if len(pdf) != expected_pages:
        raise RuntimeError(f"В PDF {len(pdf)} стр., ожидается {expected_pages} для {damage_count} повреждений")
    pages = [p.get_text() for p in pdf]
    faults = []
    sigs_by_page = []
    for i, body in enumerate(pages, 1):
        if f"лист№{i}актаосмотра№" not in norm(body):
            faults.append(f"нет правильного номера листа {i}")
        if claim and claim not in body:
            faults.append(f"номер убытка отсутствует на листе {i}")
        sigs = set()
        page = pdf[i-1]
        for img in page.get_images(full=True):
            xref = img[0]
            for rect in page.get_image_rects(xref):
                if rect.x0 > page.rect.width * 0.65 and rect.y0 > page.rect.height * 0.50:
                    sigs.add(xref)
        if not sigs:
            faults.append(f"нет факсимиле в зоне подписи эксперта на листе {i}")
        sigs_by_page.append(sigs)
    if sigs_by_page and not set.intersection(*sigs_by_page):
        faults.append("факсимиле различается между листами")
    if vin and norm(vin) not in norm(pages[0]):
        faults.append("VIN не найден на первом листе")
    heading = "Повреждения, указанные"
    if heading not in pages[2]:
        faults.append("таблица повреждений начинается не на третьем листе")
    if heading in pages[0] or heading in pages[1]:
        faults.append("таблица повреждений попала на первый или второй лист")
    if CONCLUSION_HEADING not in pages[-1] or any(CONCLUSION_HEADING in body for body in pages[:-1]):
        faults.append(f"заключение должно находиться на листе {expected_pages}")
    if norm("Осмотр осуществил") not in norm(pages[-1]):
        faults.append(f"нет блока подписей на листе {expected_pages}")
    damaged = norm(" ".join(pages[2:-1]))
    for n, damage in enumerate(damages, 1):
        for field in ("detail", "description"):
            fragment = str(damage.get(field, "")).strip()
            if not fragment or norm(fragment) not in damaged:
                faults.append(f"позиция ДТП {n}: {field} отсутствует или обрезана в PDF")
    numbers = [int(n) for n in re.findall(
        r"(?m)^\s*(\d{1,3})\s*$", "\n".join(pages[2:-1]))]
    if expected_count is not None:
        if len(numbers) != expected_count:
            faults.append(f"число строк ДТП в PDF {len(numbers)}, в JS {expected_count}")
        if numbers != list(range(1, expected_count + 1)):
            faults.append(f"нумерация повреждений: {numbers}, ожидалось 1..{expected_count}")
    if checks:
        mapped = [
            (0, "ОСАГО", str(checks.get("insurance", "")).upper() == "ОСАГО"),
            (0, "КАСКО", str(checks.get("insurance", "")).upper() == "КАСКО"),
            (0, "первичный", checks.get("inspection") == "первичный"),
            (0, "выездной", checks.get("place") == "выездной"),
            (1, "в светлое время суток", checks.get("light") == "day"),
            (1, "в грязном виде", checks.get("cleanliness") in ("грязное", "грязный")),
            (1, "в условиях СТОА", checks.get("stoa") is True),
        ]
        for page_idx, label, selected in mapped:
            if not has_checked(pages[page_idx], label, selected):
                faults.append(f"галочка листа {page_idx+1}: {label} (ожидалось {selected})")
        pair_states = [
            (0, "Идентификационные характеристики", checks.get("vinOk")),
            (0, "Дополнительное оборудование", checks.get("extraEquip")),
            (len(pages)-2, "Осмотр проводился с проведением демонтажных работ", checks.get("dismantling")),
            (len(pages)-1, CONCLUSION_HEADING, checks.get("match")),
            (len(pages)-1, "Скрытые повреждения, возможно наличие", checks.get("hidden")),
        ]
        if data_fields is not None:
            pair_states += [
                (1, "Дефекты эксплуатации транспортного средства", True if data_fields.get("defects") else None),
                (1, "Повреждения, находящиеся в зоне аварийных", True if data_fields.get("uncertainDamages") else None),
            ]
        for page_idx, label, state in pair_states:
            if not check_pair(pages[page_idx], label, state):
                faults.append(f"отметки Да/Нет на листе {page_idx+1}: {label}, ожидалось {state}")
        repairable = checks.get("repairable")
        if repairable is not None:
            if not has_checked(pages[-1], "ТС подлежит ремонту", repairable):
                faults.append("неверно отмечено: ТС подлежит ремонту")
            if not has_checked(pages[-1], "ТС не подлежит ремонту", not repairable):
                faults.append("неверно отмечено: ТС не подлежит ремонту")
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
    ap.add_argument("--expert-review", type=Path, help="JSON-отчёт экспертной проверки для рабочего акта")
    ap.add_argument("--test", action="store_true", help="Разрешить неполные тестовые реквизиты")
    args = ap.parse_args()
    if not args.generator and not args.workdir:
        ap.error("Нужен --generator или --workdir")
    generator = args.generator.resolve() if args.generator else args.workdir.resolve() / "generator.js"
    if not generator.is_file():
        ap.error(f"Не найден {generator}")
    reviewed_hash = None
    if not args.test or args.expert_review is not None:
        try:
            reviewed_hash = check_expert_review(args.expert_review, generator)
        except ValueError as exc:
            ap.error(str(exc))
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
    if not args.test and not args.approved_damages:
        ap.error("Реальный акт: обязателен --approved-damages с согласованным перечнем")
    if not args.test and data["SIGN_EXPERT"] is not True:
        ap.error("В рабочем акте отключено факсимиле эксперта")
    if data["CHECKS"].get("dismantling") is True:
        ap.error("Текущий автономный JS всегда ставит 'Нет' для демонтажа; акт с демонтажом блокирован")
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
        else:
            env.pop("AKT_TEST", None)
        if reviewed_hash is not None and sha256_file(generator) != reviewed_hash:
            ap.error("Входной JS изменён после экспертной проверки")
        run(["node", str(generator)], 25, env=env)
        t1 = time.monotonic()
        profile = (temp / "lo-profile").as_uri()
        run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(docx)], 35)
        t2 = time.monotonic()
        pdf = temp / "act.pdf"
        if not pdf.exists():
            raise RuntimeError("LibreOffice не создал PDF")
        pages = check_pdf(pdf, claim, vin, data["damages"], expected_count=len(names), checks=data["CHECKS"], data_fields=data["D"])
        dest = args.output.resolve()
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(pdf, dest)
        t3 = time.monotonic()
    print(json.dumps({"ok": True, "pdf": str(dest), "pages": pages, "damages": len(names), "seconds": {"docx": round(t1-t0, 2), "pdf": round(t2-t1, 2), "check_copy": round(t3-t2, 2), "total": round(t3-t0, 2)}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
