const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
  Footer, PageNumber, UnderlineType,
  ImageRun, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType
} = require('docx');
const fs = require('fs');

// Движок акта при эвакуации: пустая форма. Данные, подпись и ФИО эксперта передаются снаружи.
module.exports = function build(cfg) {
  const { D, KOMPL, LIGHT, CLEAN, PHOTO, SIGN_EXPERT, OWNER_MODE, OWNER_NOTE, EXPERT, SIG_RLE, OUT_FILE } = cfg;
// ===================================================================
// ПОДПИСИ — настройка под конкретный акт
// ===================================================================

// Подпись хранится как «строки штрихов» (а не картинка в base64): если при
// переписывании ошибиться в символе — сдвинется кусочек одной линии, а не
// сломается вся картинка. Ниже — сборка PNG из этих строк.
const SIG_W = 240;
function sigPng() {
  const zlib = require('zlib');
  const rows = SIG_RLE.split(","), H = rows.length, W = SIG_W;
  const raw = Buffer.alloc((W * 4 + 1) * H);
  rows.forEach((r, y) => {
    const o = y * (W * 4 + 1);
    for (let i = 0; i + 3 < r.length; i += 4) {
      const s = parseInt(r.substr(i, 2), 36), l = parseInt(r.substr(i + 2, 2), 36);
      if (isNaN(s) || isNaN(l)) continue;
      for (let x = s; x < Math.min(W, s + l); x++) { const k = o + 1 + x * 4; raw[k] = 25; raw[k+1] = 35; raw[k+2] = 140; raw[k+3] = 255; }
    }
  });
  const T = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; T[n] = c >>> 0; }
  const crc = b => { let c = 0xffffffff; for (const x of b) c = T[(c ^ x) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ih = Buffer.alloc(13); ih.writeUInt32BE(W, 0); ih.writeUInt32BE(H, 4); ih[8] = 8; ih[9] = 6;
  return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk('IHDR', ih), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const SIG_PNG = sigPng();
const ULINE = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const NBX = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const ulBorders = { top: NBX, left: NBX, right: NBX, bottom: ULINE };
const noBordersSig = { top: NBX, left: NBX, right: NBX, bottom: NBX };

// Подпись «плавает» и заходит на линию снизу, строку эксперта не раздувает
function sigImg() {
  return new ImageRun({
    type: "png", data: SIG_PNG,
    transformation: { width: 58, height: 42 },
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.CHARACTER, offset: 120000 },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
      wrap: { type: TextWrappingType.NONE },
      allowOverlap: true, behindDocument: false, layoutInCell: false, zIndex: 10,
    },
  });
}
function sigCell(children, width, bd, va = VerticalAlign.BOTTOM) {
  return new TableCell({
    borders: bd, width: { size: width, type: WidthType.DXA }, verticalAlign: va,
    margins: { top: 0, bottom: 10, left: 60, right: 60 },
    children: [new Paragraph({ spacing: { before: 0, after: 0 }, children })],
  });
}
// Строка подписи: подпись слева, затем линия ФИО и линия подписи, без рамок таблицы.
// rows: [{ label, fio, note, expert }]
function sigLines(rows, sz, cols = [3100, 3060, 3200]) {
  const sm = (text, size, bold = false) => new TextRun({ text, size, bold, font: "Times New Roman" });
  const trs = rows.map(r => new TableRow({
    height: { value: r.expert && SIGN_EXPERT ? 500 : 330, rule: "atLeast" },
    children: [
      sigCell([sm(r.label, sz, true)], cols[0], noBordersSig),
      sigCell([sm(r.fio || "", sz, true)], cols[1], ulBorders),
      r.expert
        ? sigCell(SIGN_EXPERT ? [sigImg()] : [], cols[2], ulBorders, VerticalAlign.TOP)
        : sigCell([sm(r.note || "", sz - 4, true)], cols[2], ulBorders),
    ],
  }));
  // одна строка подсказок «ФИО / подпись» под всем блоком
  trs.push(new TableRow({ children: [
    sigCell([], cols[0], noBordersSig),
    sigCell([sm("ФИО", sz - 4)], cols[1], noBordersSig, VerticalAlign.TOP),
    sigCell([sm("подпись", sz - 4)], cols[2], noBordersSig, VerticalAlign.TOP),
  ]}));
  return new Table({ width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: cols, rows: trs });
}

const B = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: B, bottom: B, left: B, right: B };
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const f = "Times New Roman";
const BUMP = 4;
const t = (text, sz=16, bd=false, ul=false) => new TextRun({ text, size: sz+BUMP, bold: bd, font: f, underline: ul ? { type: UnderlineType.SINGLE } : undefined });
const p = (children, align=AlignmentType.LEFT) => new Paragraph({ alignment: align, spacing: { before: 0, after: 0 }, children });
const ck = (on) => on ? "☒" : "☐";
const val = (v, sz=16) => t(v || "", sz, true);
function cell(children, width, opts={}) {
  const { cs=1, va=VerticalAlign.CENTER, bd=borders } = opts;
  return new TableCell({ borders: bd, width: { size: width, type: WidthType.DXA }, columnSpan: cs, verticalAlign: va,
    margins: { top: 30, bottom: 30, left: 80, right: 80 }, children: Array.isArray(children) ? children : [children] });
}
function row(...cells) { return new TableRow({ cantSplit: true, children: cells }); }
function tbl(cols, rows) { return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths: cols, rows }); }
function sp() { return new Paragraph({ spacing: { before: 0, after: 80 }, children: [t("")] }); }
function choice(options, chosen, sz=16) {
  const out = [];
  options.forEach((o, i) => { const on = Array.isArray(chosen) ? chosen.includes(o) : chosen === o;
    if (i) out.push(t("   ", sz)); out.push(t(ck(on) + " " + o, sz, on)); });
  return out;
}
function pick(options, chosen, sz=16) {
  const out = [];
  options.forEach((o, i) => { const on = chosen === o; if (i) out.push(t(" / ", sz)); out.push(t(o, sz, on, on)); });
  return out;
}
const listOrNo = (arr, sz=16) => arr.length ? arr.map((x, i) => p([t(`${i+1}. ${x}`, sz)])) : [p([t("Нет", sz, true)])];

function sigTable() {
  return sigLines([
    { label: "Собственник/Доверенное лицо", fio: D.proxy || D.owner, note: OWNER_NOTE },
    { label: "Заинтересованные лица" },
    { label: "Осмотр осуществил", fio: EXPERT, expert: true },
  ], 15+BUMP);
}
function pageFooter() {
  return new Footer({ children: [ p([ t("Лист №",14,true),
    new TextRun({ children: [PageNumber.CURRENT], bold: true, size: 14+BUMP, font: f }),
    t("    Акта осмотра №" + (D.claimNo || "________________"),14,true) ], AlignmentType.RIGHT) ] });
}

// ===== КОМПЛЕКТАЦИЯ =====
const KS = 14;
const KOMPL_ITEMS = [
  ["gear","",["АКПП","МКПП"]], ["lock","Ц.замок (с ДУ)"], ["headwash","Омыв. фар"], ["spoiler","Спойлер доп."],
  ["drive","Привод",["Пер","Зад","Пол"]], ["park","Датч.парк.",["Пер","Зад"]], ["lights","",["Ксенон","адапт","светод"]], ["steps","Подножки доп."],
  ["abs","ABS"], ["camera","Камера"], ["windows","Эл. ст. дверей",["2","4"]], ["guards","Дуги защитные"],
  ["airbag","Airbag",null,(v)=>`(${v===true?"__":v} шт.)`], ["sunroof","Люк",["Эл.","мех."]], ["handles","Ручки двер",["окр","не окр"]], ["alarm","Охр. система"],
  ["climate","",["Кондиц.","Климат"]], ["susp","Рег. подвеска"], ["trim","",["Кожа","велюр","ткань"]], ["sensors","Датч.",["дождя","подогр"]],
  ["mirrors","Зеркала",["эл.","обогр"]], ["gas","Газ. Бал. Обор."], ["alloy","Диски легк.сплав.",null,(v)=>`R${v===true?"__":v}`], ["tires","Шины",null,(v)=>v===true?"":v],
  ["mirrPaint","Зеркала",["окр","не окр"]], ["tow","Прицеп. уст-во"], ["fog","П/тум. фары"], ["liners","Подкрылки доп."],
  ["ads","Реклама наружная"], ["vinyl","Плёнка винил"],
];
function komplRuns([key, label, opts, fmt]) {
  const v = KOMPL[key]; const runs = [t(ck(!!v) + " ", KS, !!v)];
  if (label) runs.push(t(label + ((opts || fmt) ? " " : ""), KS, !!v));
  if (opts) opts.forEach((o, i) => { const on = v === o || (Array.isArray(v) && v.includes(o)); if (i) runs.push(t("/", KS)); runs.push(t(o, KS, on, on)); });
  if (fmt) runs.push(t(fmt(v || true), KS, !!v));
  return runs;
}
const komplRows = [];
for (let i = 0; i < KOMPL_ITEMS.length; i += 4) {
  const chunk = KOMPL_ITEMS.slice(i, i + 4);
  const cells = chunk.map(it => cell(p(komplRuns(it)), 2340));
  if (chunk.length === 2) cells.push(cell(p(komplRuns(["deflector","Дефлектор",["капота","окон пер.","окон зад."]])), 4680, {cs:2}));
  komplRows.push(row(...cells));
}

// -------------------- ЛИСТ 1 --------------------
const dateTxt = D.day ? `«${D.day}» ${D.month} ${D.year4} г.` : "«____» ________________ 20___ г.";
const hm = (a) => a[0] ? [val(a[0]), t(" час. ",16), val(a[1]), t(" мин.",16)] : [t("_____ час. _____ мин.",16)];
const sheet1 = [
  p([t("ФАВОРИТ",22,true)], AlignmentType.CENTER),
  sp(),
  p([t("АКТ ОСМОТРА ТРАНСПОРТНОГО СРЕДСТВА",20,true)], AlignmentType.CENTER),
  p([t(ck(D.before)+" ПЕРЕД ТРАНСПОРТИРОВКОЙ НА ВОССТАНОВИТЕЛЬНЫЙ РЕМОНТ НА СТОА",15,true)], AlignmentType.CENTER),
  p([t(ck(!D.before)+" ПОСЛЕ ТРАНСПОРТИРОВКИ ИЗ ВОССТАНОВИТЕЛЬНОГО РЕМОНТА НА СТОА",15,true)], AlignmentType.CENTER),
  p([t("№ " + (D.claimNo || "__________________"),20,true)], AlignmentType.CENTER),
  sp(),
  p([t("Основание для осмотра: заявка на осмотр № ",16,true), val(D.zayavka)]),
  p([t("Адрес осмотра: ",16,true), val(D.address)]),
  p([t("Осмотр проводился ",16), val(dateTxt), t(" с ",16), ...hm(D.start), t(" по ",16), ...hm(D.end)]),
  sp(),
  tbl([4680,4680], [row(
    cell([ p([t("Осмотр ТС проводится:",16,true)]),
      p([t(ck(LIGHT==="day")+" в светлое время суток",15,LIGHT==="day")]),
      p([t(ck(LIGHT==="art")+" в тёмное время суток с использованием искусственного освещения",15,LIGHT==="art")]),
      p([t(ck(LIGHT==="dark")+" в тёмное время суток без использования искусственного освещения",15,LIGHT==="dark")]) ],4680),
    cell([ p([t("ТС предоставлено на осмотр:",16,true)]),
      p([t(ck(CLEAN==="clean")+" в чистом виде",15,CLEAN==="clean")]),
      p([t(ck(CLEAN==="dirty")+" в грязном виде",15,CLEAN==="dirty")]) ],4680),
  )]),
  sp(),
  p([t("Сведения о ТС",17,true)], AlignmentType.CENTER),
  tbl([3000,6360], [
    row(cell(p([t("Регистрационный знак:",16,true)]),3000), cell(p([val(D.plate)]),6360)),
    row(cell(p([t("VIN (№ кузова)",16,true)]),3000), cell(p([val(D.vin)]),6360)),
    row(cell(p([t("Марка, модель",16,true)]),3000), cell(p([val(D.model)]),6360)),
    row(cell(p([t("Тип кузова",16,true)]),3000), cell(p([val(D.bodyType), t("          Число дверей: ",16), ...pick(["2","3","4","5"], D.doors)]),6360)),
    row(cell(p([t("Дата начала эксплуатации:",16,true)]),3000), cell(p([val(D.startDate), t("          Год выпуска: ",16), val(D.year)]),6360)),
    row(cell(p([t("Показ. одометра",16,true)]),3000), cell(p([val(D.mileage), t(/^\d/.test(D.mileage) ? " км" : "",16,true)]),6360)),
    row(cell(p([t("Шасси (рама) №",16,true)]),3000), cell(p([val(D.chassis)]),6360)),
    row(cell(p([t("Цвет, вид покрытия",16,true)]),3000), cell(p([val(D.color), t("          ",16), ...pick(["акрил","метал.","перл."], D.paint)]),6360)),
    row(cell(p([t("Двигатель",16,true)]),3000), cell([
      p([t("Модель, №: ",16), val(D.engineNo), t("   Мощн.: ",16), val(D.power), t("   Объём: ",16), val(D.volume)]),
      p(choice(["Бенз.","диз.","гибрид.","Впр.","Карбюратор","Турбо"], D.fuel, 15)) ],6360)),
    row(cell(p([t("Свид. о рег. / ПТС",16,true)]),3000), cell(p([t("Серия ",16), val(D.stsSeries), t("  № ",16), val(D.stsNumber)]),6360)),
  ]),
  sp(),
  p([t("Комплектация",17,true)], AlignmentType.CENTER),
  tbl([2340,2340,2340,2340], [...komplRows, row(cell(p([t("Прочее: ",KS), val(KOMPL.other || "", KS)]),9360,{cs:4}))]),
  sp(),
  sigTable(),
];

// -------------------- ЛИСТ 2 --------------------
const sheet2 = [
  new Paragraph({ children: [new PageBreak()] }),
  p([t("1. Повреждения транспортного средства, относящиеся к заявленному событию — ДТП от ",16,true), val(D.eventDate || "____________"), t(":",16,true)]),
  p([val(D.basisActs || "Согласно акту(ам) осмотра по убытку")]),
  sp(),
  p([t("2. Дефекты эксплуатации транспортного средства — повреждения, не относящиеся к заявленному случаю, полученные до заявленного события:",16,true)]),
  ...listOrNo(D.defects),
  sp(),
  p([t("3. Повреждения, находящиеся в зоне аварийных, принадлежность которых к заявленному событию определить не представляется возможным:",16,true)]),
  ...listOrNo(D.unclear),
  sp(),
  sigTable(),
];

// -------------------- ЛИСТ 3 --------------------
const names = Object.keys(D.equip);
const half = Math.ceil(names.length / 2);
const yn = (v, want) => v === null || v === undefined ? "☐" : ck(v === want);
function eqCells(n, name) {
  if (!name) return [cell(p([t("")]),400), cell(p([t("")]),2600), cell(p([t("")]),500), cell(p([t("")]),500)];
  const v = D.equip[name];
  return [ cell(p([t(String(n),14)],AlignmentType.CENTER),400), cell(p([t(name,14)]),2600),
    cell(p([t(yn(v,true),14,v===true)],AlignmentType.CENTER),500), cell(p([t(yn(v,false),14,v===false)],AlignmentType.CENTER),500) ];
}
const eqHead = row(...[0,1].flatMap(() => ["№","Наименование","ДА","НЕТ"].map((h,i)=>cell(p([t(h,13,true)],AlignmentType.CENTER),[400,2600,500,500][i]))));
const eqRows = names.slice(0, half).map((n, i) => row(...eqCells(i+1, n), ...eqCells(half+i+1, names[half+i])));
const sheet3 = [
  new Paragraph({ children: [new PageBreak()] }),
  p([t("4. Дополнительная информация.",16,true)]),
  p([t("Также вместе с ТС передано:",16,true)]),
  tbl([400,2600,500,500,400,2600,500,500], [eqHead, ...eqRows]),
  sp(),
  p([t("Сигнализаторы неисправностей на щитке приборов автомобиля:",16,true)]),
  p(choice(["АБС","Ошибка ДВС","Торм. жидкость","Тем-ра","Давл. масла","ESP"], D.warnings, 15)),
  sp(),
  p([t("Подкапотное пространство осмотрено (выявлены следующие повреждения): ",16,true), val(D.underhood || "повреждений не выявлено")]),
  sp(),
  p([t("Количество топлива (полный, пустой, 50%): ",16,true), val(D.fuelLevel || "____________"), t("      Уровень масла в ДВС: ",16,true), val(D.oilLevel || "____________")]),
  sp(),
  p([t("Фотофиксация дефектов:  ",16,true), ...choice(["ДА","НЕТ"], PHOTO ? "ДА" : "НЕТ")]),
  sp(),
  p([t("Техническое состояние ТС (возможность самостоятельного передвижения):  ",16,true), ...choice(["ИСПРАВНО","НЕИСПРАВНО"], D.techOk ? "ИСПРАВНО" : "НЕИСПРАВНО")]),
  sp(),
  sigTable(),
];

const doc = new Document({ sections: [{
  properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, right: 700, bottom: 900, left: 1000, footer: 400 } } },
  footers: { default: pageFooter() },
  children: [...sheet1, ...sheet2, ...sheet3],
}] });
return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT_FILE, buf); console.log("OK"); });

};
