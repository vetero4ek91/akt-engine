const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
  Footer, PageNumber, UnderlineType,
  ImageRun, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType
} = require('docx');
const fs = require('fs');

// Движок акта: пустая форма. Данные, подпись и ФИО эксперта передаются снаружи.
module.exports = function build(cfg) {
  const { D, KOMPL, LIGHT, CLEAN, SPRAVKA, DEMONTAZH, STOA, MATCH, REPAIRABLE, HIDDEN, KASKO, damages, OUT_FILE, SIGN_EXPERT, OWNER_MODE, OWNER_NOTE, SIG_RLE, EXPERT } = cfg;
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

// ===================================================================
// КОД ШАБЛОНА — ниже ничего не менять
// ===================================================================
const B = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: B, bottom: B, left: B, right: B };
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const ulOnly = { top: NB, left: NB, right: NB, bottom: B };

const f = "Times New Roman";
const BUMP = 4; // +2pt ко всем размерам
const t = (text, sz=18, bd=false, ul=false) => new TextRun({ text, size: sz+BUMP, bold: bd, font: f, underline: ul ? { type: UnderlineType.SINGLE } : undefined });
const p = (children, align=AlignmentType.LEFT, spacing={}) => new Paragraph({ alignment: align, spacing, children });
const ck = (on) => on ? "☒" : "☐";

function cell(children, width, opts={}) {
  const { colspan=1, rowspan=1, valign=VerticalAlign.CENTER, bd=borders } = opts;
  return new TableCell({
    borders: bd, width: { size: width, type: WidthType.DXA },
    columnSpan: colspan, rowSpan: rowspan, verticalAlign: valign,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: Array.isArray(children) ? children : [children]
  });
}
function row(...cells) { return new TableRow({ cantSplit: true, children: cells }); }
function tbl(cols, rows) { return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths: cols, rows }); }
function sp(n=1) { return Array.from({length:n}, ()=>p([t("")])); }
const val = (v, sz=16) => t(v || "", sz, true);

// Варианты: выбранный — ☒ и жирным
function choice(options, chosen, sz=16) {
  const out = [];
  options.forEach((o, i) => {
    const on = Array.isArray(chosen) ? chosen.includes(o) : chosen === o;
    if (i) out.push(t("   ", sz));
    out.push(t(ck(on) + " " + o, sz, on));
  });
  return out;
}
// Варианты внутри строки: выбранный подчёркнут жирным, без чекбоксов (2 / 3 / 4 / 5)
function pick(options, chosen, sz=16) {
  const out = [];
  options.forEach((o, i) => {
    const on = chosen === o;
    if (i) out.push(t(" / ", sz));
    out.push(t(o, sz, on, on));
  });
  return out;
}
const yesNo = (yes, sz=16) => choice(["Да", "Нет"], yes ? "Да" : "Нет", sz);
const listOrNo = (arr, sz=16) => arr.length ? arr.map((x, i) => p([t(`${i+1}. ${x}`, sz)])) : [];

function sigTable() {
  return sigLines([
    { label: "Собственник/Доверенное лицо", fio: D.proxy || D.owner, note: OWNER_NOTE },
    { label: "Заинтересованные лица" },
    { label: "Осмотр осуществил", fio: EXPERT, expert: true },
  ], 16+BUMP);
}

function pageFooter() {
  return new Footer({ children: [ p([
    t("Лист №",16,true),
    new TextRun({ children: [PageNumber.CURRENT], bold: true, size: 16+BUMP, font: f }),
    t("    Акта осмотра №" + D.claimNo,16,true),
  ], AlignmentType.RIGHT) ] });
}

// ===== КОМПЛЕКТАЦИЯ =====
// [ключ, подпись, варианты, формат значения]
const KOMPL_ITEMS = [
  ["gear","",["АКПП","МКПП"]], ["lock","Ц.замок (с ДУ)"], ["headwash","Омыв. фар"], ["spoiler","Спойлер доп."],
  ["drive","Привод",["Пер","Зад","Пол"]], ["park","Датч.парк.",["Пер","Зад"]], ["lights","",["Ксенон","адапт","светод"]], ["steps","Подножки доп."],
  ["abs","ABS"], ["camera","Камера"], ["windows","Эл. ст. дверей",["2","4"]], ["guards","Дуги защитные"],
  ["airbag","Airbag",null,(v)=>`(${v===true?"__":v} шт.)`], ["sunroof","Люк",["Эл.","мех."]], ["handles","Ручки двер",["окр","не окр"]], ["alarm","Охр. система"],
  ["climate","",["Кондиц.","Климат"]], ["susp","Рег. подвеска"], ["trim","",["Кожа","велюр","ткань"]], ["sensors","Датч",["дождя","подог"]],
  ["mirrors","Зеркала",["эл.","обогр"]], ["gas","Газ. Бал. Обор."], ["alloy","Диски легк.сплав.",null,(v)=>`R${v===true?"__":v}`], ["tires","Шины",null,(v)=>v===true?"":v],
  ["mirrPaint","Зеркала",["окр","не окр"]], ["tow","Прицеп. уст-во"], ["fog","П/туман. фары"], ["liners","Подкрылки доп"],
  ["ads","Реклама наружная"], ["vinyl","Плёнка винил"],
];
const KS = 14;
function komplRuns([key, label, opts, fmt]) {
  const v = KOMPL[key];
  const runs = [t(ck(!!v) + " ", KS, !!v)];
  if (label) runs.push(t(label + ((opts || fmt) ? " " : ""), KS, !!v));
  if (opts) opts.forEach((o, i) => {
    const on = v === o || (Array.isArray(v) && v.includes(o));
    if (i) runs.push(t("/", KS));
    runs.push(t(o, KS, on, on));
  });
  if (fmt) runs.push(t(v ? fmt(v) : fmt(true), KS, !!v));
  return runs;
}
const komplRows = [];
for (let i = 0; i < KOMPL_ITEMS.length; i += 4) {
  const chunk = KOMPL_ITEMS.slice(i, i + 4);
  const cells = chunk.map(it => cell(p(komplRuns(it)), 2340));
  if (chunk.length === 2) cells.push(cell(p(komplRuns(["deflector","Дефлектор",["капота","окон пер.","окон зад."]])), 4680, {colspan:2}));
  komplRows.push(row(...cells));
}

// =================== ЛИСТ 1 ===================
const sheet1 = [
  p([t("ООО «Фаворит»",22,true)], AlignmentType.CENTER),
  p(choice(["Выезд", "Фото"], D.mode, 16), AlignmentType.RIGHT),
  // СК и полис — таблицей без рамок, чтобы ничего не «уезжало»
  tbl([1700,3000,2500,2160], [
    row(cell(p([t("Название СК",16,true)]),1700,{bd:noBorders}), cell(p([val(D.insurer || "Росгосстрах")]),3000,{bd:ulOnly}),
        cell(p([t("Страховой полис №",16,true)],AlignmentType.RIGHT),2500,{bd:noBorders}), cell(p([val(D.polis)]),2160,{bd:ulOnly})),
    row(cell(p([t("СК виновника",16,true)]),1700,{bd:noBorders}), cell(p([val(D.culpritSK)]),3000,{bd:ulOnly}),
        cell(p([t("Полис виновника №",16,true)],AlignmentType.RIGHT),2500,{bd:noBorders}), cell(p([val(D.culpritPolis)]),2160,{bd:ulOnly})),
  ]),
  ...sp(1),
  p([t("АКТ ОСМОТРА ТРАНСПОРТНОГО СРЕДСТВА",22,true)], AlignmentType.CENTER),
  p([t("№  " + D.claimNo,20,true)], AlignmentType.CENTER),
  ...sp(1),
  p(choice(["первичный","дополнительный","после восстановительного ремонта","виновника"], D.inspType, 16), AlignmentType.CENTER),
  ...sp(1),
  p([t("Основание для осмотра: Направление на осмотр № ",16,true), val(D.claimNo)]),
  p([t("Адрес осмотра: ",16,true), val(D.address)]),
  p([t("Осмотр проводился ",16), val(`«${D.day}» ${D.month} ${D.year4} г.`), t(" с ",16), val(D.start[0]), t(" час. ",16), val(D.start[1]), t(" мин. по ",16), val(D.end[0]), t(" час. ",16), val(D.end[1]), t(" мин.",16)]),
  ...sp(1),
  tbl([4680,4680], [
    row(cell(p([t("Осмотр ТС проводится:",16,true)]),4680), cell(p([t("ТС предоставлено на осмотр:",16,true)]),4680)),
    row(cell(p([t(ck(LIGHT==="day")+" в светлое время суток",16,LIGHT==="day")]),4680), cell(p([t(ck(CLEAN==="clean")+" в чистом виде",16,CLEAN==="clean")]),4680)),
    row(cell(p([t(ck(LIGHT==="art")+" в темное время суток с использованием искусственного освещения",16,LIGHT==="art")]),4680), cell(p([t(ck(CLEAN==="dirty")+" в грязном виде",16,CLEAN==="dirty")]),4680)),
    row(cell(p([t(ck(LIGHT==="dark")+" в темное время суток без использования искусственного освещения",16,LIGHT==="dark")]),4680), cell(p([t("")]),4680)),
    row(cell(p([t("Сведения о ТС",16,true)],AlignmentType.CENTER),9360,{colspan:2})),
  ]),
  tbl([3000,6360], [
    row(cell(p([t("Регистрационный знак:",16,true)]),3000), cell(p([val(D.plate)]),6360)),
    row(cell(p([t("VIN (№ кузова)",16,true)]),3000), cell(p([val(D.vin)]),6360)),
    row(cell(p([t("Марка, модель",16,true)]),3000), cell(p([val(D.model)]),6360)),
    row(cell(p([t("Тип кузова",16,true)]),3000), cell(p([val(D.bodyType), t("          Число дверей: ",16), ...pick(["2","3","4","5"], D.doors)]),6360)),
    row(cell(p([t("Дата начала эксплуатации:",16,true)]),3000), cell(p([val(D.startDate), t("          Год выпуска: ",16), val(D.year)]),6360)),
    row(cell(p([t("Показ. одометра",16,true)]),3000), cell(p([val(D.mileage), t(/^\d/.test(D.mileage) ? " км" : "",16,true)]),6360)),
    row(cell(p([t("Шасси (рама) №",16,true)]),3000), cell(p([val(D.chassis)]),6360)),
    row(cell(p([t("Цвет, вид покрытия",16,true)]),3000), cell(p([val(D.color), t("          ",16), ...pick(["акрил","метал.","перл."], D.paint)]),6360)),
    row(cell(p([t("Двигатель",16,true)]),3000), cell(p(pick(["Бенз.","диз.","гибрид."], D.fuel)),6360)),
    row(cell(p([t("Свид. о рег. / ПТС",16,true)]),3000), cell(p([t("Серия ",16), val(D.stsSeries), t("  № ",16), val(D.stsNumber), t("   /  ПТС: ",16), val(D.pts)]),6360)),
  ]),
  p([t("Комплектация",16,true)]),
  tbl([2340,2340,2340,2340], [
    ...komplRows,
    row(cell(p([t("Бронированный  ",KS), ...pick(["да","нет"], KOMPL.armored ? "да" : "нет", KS)]),9360,{colspan:4})),
    row(cell(p([t("Прочее: ",KS), val(KOMPL.other || "", KS)]),9360,{colspan:4})),
  ]),
  ...sp(1),
  sigTable(),
];

// =================== ЛИСТ 2 ===================
const sheet2 = [
  new Paragraph({ children: [new PageBreak()] }),
  sigLines([
    { label: "Собственник", fio: D.owner, note: OWNER_NOTE },
    { label: "Довер. лицо", fio: D.proxy, note: OWNER_MODE === "proxy" ? OWNER_NOTE : "" },
  ], 14+BUMP),
  ...sp(1),
  tbl([9360], [
    row(cell(p([t("Идентификационные характеристики и параметры ТС (номер кузова, шасси, двигателя, государственный регистрационный знак, цвет краски) соответствуют регистрационным документам",16,true)]),9360)),
    row(cell([p(yesNo(!D.vinMismatch)), ...(D.vinMismatch ? [p([t("Несоответствие: ",16), val(D.vinMismatch)])] : [])],9360)),
  ]),
  ...sp(1),
  p([t("Дополнительное оборудование и отступление от стандартной комплектации (в зоне аварийных повреждений)",16,true)]),
  p(yesNo(D.extraEquip.length > 0)),
  ...listOrNo(D.extraEquip),
  ...sp(1),
  tbl([9360], [
    row(cell(p([t("Справка из компетентных органов  ",16), ...choice(["предоставлена","не предоставлена"], {yes:"предоставлена",no:"не предоставлена"}[SPRAVKA])]),9360)),
    row(cell(p([t("Дата и время ДТП (события): ",16,true), val(D.eventDate)]),9360)),
    row(cell(p([t("Место ДТП (события): ",16,true), val(D.eventPlace)]),9360)),
  ]),
  ...sp(1),
  p([t("Дефекты эксплуатации транспортного средства - повреждения, не относящиеся к заявленному случаю, полученные до заявленного события:  ",16,true), ...yesNo(D.defects.length > 0)]),
  ...listOrNo(D.defects, 15),
  ...sp(1),
  p([t("Повреждения, находящиеся в зоне аварийных, принадлежность которых к заявленному событию определить не представляется возможным:  ",16,true), ...yesNo(D.unclear.length > 0)]),
  ...listOrNo(D.unclear, 15),
  ...sp(1),
  sigTable(),
];

// =================== ЛИСТ 3 — повреждения ===================
const colW = [2600, 4200, 1200, 1360];
function damRow(n, d) {
  return row(
    cell(p([t(`${n}. ${d.name}`,15,true)]), colW[0]),
    cell(p([t(d.vid || "",15)]), colW[1]),
    cell(p([t(d.mesto || "",15)]), colW[2]),
    cell(p([t(d.obem || "",15)]), colW[3]),
  );
}
const damHeader = new TableRow({ tableHeader: true, children: [
  cell(p([t("Наименование поврежденных деталей (узлов, агрегатов)",15,true)]),colW[0]),
  cell(p([t("Вид, характер",15,true)]),colW[1]),
  cell(p([t("Место",15,true)]),colW[2]),
  cell(p([t("Объем (степень) повреждения",15,true)]),colW[3]),
]});
const sheet3 = [
  new Paragraph({ children: [new PageBreak()] }),
  p([t("Повреждения, указанные в документах компетентных органов/Извещении о ДТП, и относящиеся к заявленному событию:",16,true)]),
  tbl(colW, [damHeader, ...damages.map((d, i) => damRow(i + 1, d))]),
  ...sp(1),
  p([t("Осмотр проводился с проведением демонтажных работ:  ",16), ...yesNo(DEMONTAZH)]),
  ...sp(1),
  sigTable(),
];

// =================== ЛИСТ 4 — заключение ===================
const ownerLine = (on) => p([t("Собственник (Владелец), доверенное лицо  ",16), on ? val(OWNER_NOTE || "", 16) : t("_________________________________",16)]);
const sheet4 = [
  new Paragraph({ children: [new PageBreak()] }),
  p([t("Выбрать один из следующих вариантов:",16,true)]),
  ...sp(1),
  p([t(ck(STOA==="stoa")+" Осмотр проведен в условиях СТОА. В ходе осмотра выявлены и зафиксированы все повреждения ТС",16,true)]),
  ownerLine(STOA==="stoa"),
  ...sp(1),
  p([t(ck(STOA==="nte")+" В ходе осмотра выявлены все видимые повреждения, требуется проведение независимой технической экспертизы (НТЭ) в условиях СТОА в порядке п.13 ст. 12 Закона об ОСАГО. На информирование о дате/времени/месте проведения НТЭ/проведение эвакуации ТС на СТОА для НТЭ по телефону/смс ________________ согласен:",16)]),
  ownerLine(STOA==="nte"),
  ...sp(1),
  p([t(ck(STOA==="refuse")+" В ходе осмотра выявлены все видимые повреждения. Владелец/доверенное лицо отказался от организации независимой технической экспертизы в условиях СТОА",16)]),
  ownerLine(STOA==="refuse"),
  ...sp(1),
  tbl([9360], [row(cell([
    p([t("Характер описанных повреждений в разделе «Повреждения, указанные в документах компетентных органов/Извещении о ДТП, и относящиеся к заявленному событию» даёт основание предварительно установить, что все они могут принадлежать к рассматриваемому ДТП. Итоговый объем повреждений и способ их устранения определяется при производстве расчета в соответствии с нормативами изготовителя ТС, технологическими особенностями конструкции или ремонта.",16,true)]),
    p(yesNo(MATCH)),
  ],9360))]),
  ...sp(1),
  p([t("Заключение (иное): ",16,true), val(D.conclusion || "Нет")]),
  ...sp(1),
  tbl([9360], [
    row(cell(p([t(ck(REPAIRABLE)+" ТС подлежит ремонту",16,REPAIRABLE)]),9360)),
    row(cell(p([t(ck(!REPAIRABLE)+" ТС не подлежит ремонту",16,!REPAIRABLE)]),9360)),
    row(cell(p([t("Информация о пробах и элементах ТС, взятых для исследования:",16,true)]),9360)),
    row(cell(p([t("☒ Пробы материала и элементов транспортного средства не брались.",16,true)]),9360)),
    row(cell(p([t("☐ Пробы материала и элементов транспортного средства взяты",16)]),9360)),
  ]),
  ...sp(1),
  p([t("Скрытые повреждения, возможно наличие в зоне основных повреждений:  ",16,true), ...yesNo(HIDDEN)]),
  p([t("Состояние ТС:  ",16,true), val(D.state)]),
  ...sp(1),
  p([t("Акт осмотра составлен в моём присутствии. По результатам осмотра подтверждаю, что все видимые повреждения, относящиеся к заявленному мною событию отражены в акте осмотра.",15)]),
  p([t("Подтверждаю, что в случае выявления повреждений ТС, не зафиксированных в настоящем Акте осмотра, в том числе ввиду представления ТС в грязном виде или проведения осмотра в условиях плохого освещения/видимости, я уведомлен об обязанности представить ТС на дополнительный осмотр Страховщику в чистом виде, согласовав в письменном виде со Страховщиком дату, время и место осмотра.",15)]),
  p([t("В случае непредставления ТС в чистом виде на дополнительный осмотр риск неблагоприятных последствий, связанных с невозможностью установить полный объем повреждений и их относимость к страховому случаю, возлагается на Потерпевшего, данные действия будут расцениваться как злоупотребление правом.",15)]),
  p([t("Я предупрежден о возможности наличия скрытых повреждений в зоне локализации, которые могут быть выявлены в ходе выполнения ремонтных работ, а также предупрежден об обязанности предъявить ТС на осмотр страховщику до начала работ по их устранению.",15)]),
  p([t("Я выражаю согласие, что факт передачи ТС в ремонт на СТОА (в случае принятия решения о возмещении ущерба в натуральной форме) оформляется актом приема-передачи ТС в ремонт и заверяется подписями потерпевшего и представителя СТОА.",15)]),
  ...sp(1),
  p([t(`Акт составлен на  ${KASKO ? 5 : 4}  листах`,16,true)], AlignmentType.CENTER),
  ...sp(1),
  sigTable(),
];

// =================== ЛИСТ 5 — Приложение (только КАСКО) ===================
const k5W = [700, 2165, 2165, 2165, 2165];
const k5mk = (v) => v === true ? "X" : (v ? String(v) : "");
const sheet5 = !KASKO ? [] : [
  new Paragraph({ children: [new PageBreak()] }),
  p([t("Приложение к Акту осмотра № " + D.claimNo + "  от  " + `${D.day}.${{"января":"01","февраля":"02","марта":"03","апреля":"04","мая":"05","июня":"06","июля":"07","августа":"08","сентября":"09","октября":"10","ноября":"11","декабря":"12"}[D.month] || D.month}.${D.year4}` + " г.",16,true)]),
  ...sp(1),
  tbl(k5W, [
    new TableRow({ tableHeader: true, children: ["№","Замена","Ремонт, час.","Окрас.","Диагн."].map((h, i) => cell(p([t(h,15,true)],AlignmentType.CENTER),k5W[i])) }),
    ...damages.map((d, i) => row(
      cell(p([t(String(i + 1),15)],AlignmentType.CENTER),k5W[0]),
      ...[d.zamena, d.remont, d.okraska, d.diagn].map((v, j) => cell(p([t(k5mk(v),15,true)],AlignmentType.CENTER),k5W[j+1])),
    )),
  ]),
  ...sp(1),
  sigLines([{ label: "Осмотр осуществил", fio: EXPERT, expert: true }], 16+BUMP),
];

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 700, right: 700, bottom: 900, left: 1000, footer: 400 } } },
    footers: { default: pageFooter() },
    children: [...sheet1, ...sheet2, ...sheet3, ...sheet4, ...sheet5],
  }]
});

return Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT_FILE, buf);
  console.log("OK");
});

};
