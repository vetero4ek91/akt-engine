const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, VerticalAlign, PageBreak,
  Footer, PageNumber, UnderlineType,
  ImageRun, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType
} = require('docx');
const fs = require('fs');

// Движок акта: пустая форма. Данные, подпись и ФИО эксперта передаются снаружи.
module.exports = function build(cfg) {
  const { D, KOMPL, UCHET_OK, LIGHT, CLEAN, WEATHER, ACT_PLACE, EVENTS, REPAIR, OTHERS, PHOTO, damages, OUT_FILE, SIGN_EXPERT, OWNER_MODE, OWNER_NOTE, SIG_RLE, EXPERT } = cfg;
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
    children: [new Paragraph({ keepNext: true, spacing: { before: 0, after: 0 }, children })],
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
const B   = { style: BorderStyle.SINGLE, size: 3, color: "000000" };
const brd = { top: B, bottom: B, left: B, right: B };
const NB  = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const nob = { top: NB, bottom: NB, left: NB, right: NB };

const f = "Times New Roman";
const t  = (text, sz=16, bd=false, ul=false) => new TextRun({ text, size: sz, bold: bd, font: f, underline: ul ? { type: UnderlineType.SINGLE } : undefined });
const p  = (children, align=AlignmentType.LEFT) => new Paragraph({ spacing: { before: 0, after: 0 }, alignment: align, children });
const ck = (on) => on ? "☒" : "☐";

function cell(children, width, opts={}) {
  const { cs=1, rs=1, va=VerticalAlign.CENTER, bd=brd } = opts;
  return new TableCell({
    borders: bd, width: { size: width, type: WidthType.DXA },
    columnSpan: cs, rowSpan: rs, verticalAlign: va,
    margins: { top: 30, bottom: 30, left: 80, right: 80 },
    children: Array.isArray(children) ? children : [children]
  });
}
function row(...cells) { return new TableRow({ cantSplit: true, children: cells }); }
function tbl(cols, rows) { return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths: cols, rows }); }
function sp() { return new Paragraph({ spacing: { before: 0, after: 50 }, children: [t("")] }); }
function line() { return p([t("_".repeat(100), 14)]); }

// Хвост акта держится одним блоком: не влезает — переезжает на след. лист целиком
const pk = (children, align=AlignmentType.LEFT) => new Paragraph({ keepNext: true, keepLines: true, spacing: { before: 0, after: 0 }, alignment: align, children });
function spk() { return new Paragraph({ keepNext: true, spacing: { before: 0, after: 50 }, children: [t("")] }); }
function linek() { return pk([t("_".repeat(100), 14)]); }

// Варианты одной строкой: выбранный — ☒ и жирным
function choice(options, chosen, sz=15) {
  const out = [];
  options.forEach((o, i) => {
    const on = Array.isArray(chosen) ? chosen.includes(o) : chosen === o;
    if (i) out.push(t("   ", sz));
    out.push(t(ck(on) + " " + o, sz, on));
  });
  return out;
}
// Значение вписанное жирным
const val = (v, sz=16) => t(v || "", sz, true);

// ===== КОМПЛЕКТАЦИЯ =====
const KOMPL_ITEMS = [
  ["alloy","Диски колес легкосплавные"], ["heatWS","Эл. подогрев ветр. стекла"], ["sunroof","Люк",["мех.","эл. прив."]], ["camera","Камера заднего вида"],
  ["xenon","Фары",["ксенон","биксенон"]], ["rain","Датчик",["дождя","света"]], ["trim","Обивка",["кожа","велюр","ткань"]], ["climate","",["Кондиц.","Климат контроль"]],
  ["headwash","Омыватели фар"], ["windows","Эл. привод стекол",["пер.","зад."]], ["seatheat","Эл. подогрев сидений"], ["gear","",["АКПП","МКПП"]],
  ["fog","Фары противотуманные"], ["mirrAdj","Эл. регулировка зеркал"], ["seatAdj","Эл. регул. сидений"], ["gas","Газовое оборудование"],
  ["moldings","Накладки дверей, арок"], ["mirrFold","Эл. привод складыв. зеркал"], ["park","Парктроник",["пер.","зад."]], ["tow","Тягово-сцепное устройство"],
];
function komplCell([key, label, opts]) {
  const v = KOMPL[key];
  const runs = [t(ck(!!v) + " ", 15, !!v)];
  if (label) runs.push(t(label + (opts ? " " : ""), 14, !!v));
  if (opts) opts.forEach((o, i) => {
    const on = v === o || (Array.isArray(v) && v.includes(o));
    if (i) runs.push(t(" / ", 14));
    runs.push(t(o, 14, on, on));
  });
  return cell(p(runs), 2340);
}
const komplRows = [];
for (let i = 0; i < KOMPL_ITEMS.length; i += 4) komplRows.push(row(...KOMPL_ITEMS.slice(i, i + 4).map(komplCell)));

// ===== ТАБЛИЦА ПОВРЕЖДЕНИЙ =====
const cW = [400, 6040, 800, 800, 800, 520];
function damHeader() {
  return new TableRow({ tableHeader: true, children: [
    cell(p([t("№ п/п",13,true)],AlignmentType.CENTER), cW[0]),
    cell([
      p([t("Наименование детали ",13,true), t("(узла, агрегата, доп. оборудования)",13), t(" с локализацией ",13,true), t("(передний, задний, левый, правый, верхний, нижний, наруж., внутр.).",13)]),
      p([t("Описание повреждений: ",13,true), t("вмятина, складка, изгиб, излом, разрыв, разбит, трещина, царапина, скол, отслоение, потертость, задир, перекос проема, смещение с монтажного места, потеря герметичности, уничтожено огнем, высокой температурой, продуктами горения, средствами пожаротушения.",13)]),
      p([t("Размер повреждений в см или в % от площади элемента.",13,true)]),
    ], cW[1]),
    cell(p([t("Замена",13,true)],AlignmentType.CENTER),  cW[2]),
    cell(p([t("Ремонт",13,true)],AlignmentType.CENTER),  cW[3]),
    cell(p([t("Окраска",13,true)],AlignmentType.CENTER), cW[4]),
    cell(p([t("Диагн.",13,true)],AlignmentType.CENTER),  cW[5]),
  ]});
}
function splitDam(text="") {
  const i = text.indexOf(" — ");
  return i === -1 ? { name: text, desc: "" } : { name: text.slice(0,i), desc: text.slice(i+3) };
}
function damRow(n, dm) {
  const { name, desc } = splitDam(dm.text);
  const mk = (on) => on ? "X" : "";
  return new TableRow({ cantSplit: true, height: { value: 360, rule: "atLeast" }, children: [
    cell(p([t(String(n),15)],AlignmentType.CENTER), cW[0]),
    cell(p([t(name,14,true), t(desc?(" — "+desc):"",14)]), cW[1]),
    cell(p([t(mk(dm.zamena),15)],AlignmentType.CENTER), cW[2]),
    cell(p([t(mk(dm.remont),15)],AlignmentType.CENTER), cW[3]),
    cell(p([t(mk(dm.okraska),15)],AlignmentType.CENTER), cW[4]),
    cell(p([t(mk(dm.diagn),15)],AlignmentType.CENTER), cW[5]),
  ]});
}
// список «Нет» или нумерованный
const listOrNo = (arr, sz=14) => arr.length ? arr.map((x, i) => p([t(`${i+1}. ${x}`, sz)])) : [p([t("Нет", sz, true)])];

// =================== СТРАНИЦА 1 ===================
const page1 = [
  p([t("LAT ASSISTANCE  /  ЛАТ АССИСТАНС", 20, true)], AlignmentType.CENTER),
  p([t("─────────────────────────────────────────────────────────────────────────────────────", 14)], AlignmentType.CENTER),

  tbl([4680,4680], [row(
    cell(p([t("Заказчик ", 16), val("Т-Страхование")]), 4680, {bd:nob}),
    cell(p([t("Номер полиса/убытка ", 16), val(D.claimNo)]), 4680, {bd:nob}),
  )]),
  p([t("Дата осмотра ", 16), val(D.dateWords)]),
  tbl([4680,4680], [row(
    cell(p([t("Начало осмотра ", 16), val(`«${D.start[0]}» час «${D.start[1]}» мин`)]), 4680, {bd:nob}),
    cell(p([t("Окончание осмотра ", 16), val(`«${D.end[0]}» час «${D.end[1]}» мин`)]), 4680, {bd:nob}),
  )]),
  p([t("Основание для осмотра ", 16), val(`Направление на осмотр № ${D.actNo} (осмотр ${D.inspType})`)]),
  p([t("Место осмотра: ", 16), val(D.address)]),
  sp(),

  p([t("АКТ ОСМОТРА ТРАНСПОРТНОГО СРЕДСТВА № " + D.actNo, 18, true)], AlignmentType.CENTER),
  sp(),

  tbl([3200,6160], [row(cell(p([t("Мною, экспертом – автотехником", 16)]), 3200), cell(p([val(EXPERT)]), 6160))]),
  tbl([3200,1600,1600,2960], [row(
    cell(p([t("Произведен осмотр ТС", 16)]), 3200),
    cell(p([t("Гос. рег. №", 16)]), 1600),
    cell(p([val(D.plate)]), 1600),
    cell(p([t("Пробег ", 16), val(D.mileage)]), 2960),
  )]),
  tbl([3000,3000,1680,1680], [row(
    cell(p([t("Марка, модель, модификация ТС", 16)]), 3000),
    cell(p([val(D.model)]), 3000),
    cell(p([t("Год вып./нач. экспл", 16)]), 1680),
    cell(p([val(D.year)]), 1680),
  )]),
  tbl([2200,1200,1200,1200,3560], [row(
    cell(p([t("Рег. документ:", 16)]), 2200),
    cell(p([t("Серия", 16)]), 1200),
    cell(p([val(D.stsSeries)]), 1200),
    cell(p([t("Номер", 16)]), 1200),
    cell(p([val(D.stsNumber)]), 3560),
  )]),
  tbl([2000,2400,1200,3760], [row(
    cell(p([t("№ кузова / рамы", 16)]), 2000),
    cell(p([val(D.bodyNo, 15)]), 2400),
    cell(p([t("VIN", 16)]), 1200),
    cell(p([val(D.vin)]), 3760),
  )]),
  tbl([1200,3600,1300,1300,800,1160], [row(
    cell(p([t("Привод:", 16)]), 1200),
    cell(p(choice(["Передний","Задний","Полный"], D.drive, 14)), 3600),
    cell(p([t("Мощность", 16)]), 1300),
    cell(p([val(D.power, 15)]), 1300),
    cell(p([t("Объем", 16)]), 800),
    cell(p([val(D.volume, 15)]), 1160),
  )]),
  tbl([1800,2600,1500,1960,1100,400], [row(
    cell(p([t("Двигатель: тип", 16)]), 1800),
    cell(p([val(D.engine, 15)]), 2600),
    cell(p([t("Тип кузова", 16)]), 1500),
    cell(p([val(D.bodyType)]), 1960),
    cell(p([t("Дверей", 16)]), 1100),
    cell(p([val(D.doors)],AlignmentType.CENTER), 400),
  )]),
  tbl([800,3200,1800,3560], [row(
    cell(p([t("Цвет", 16)]), 800),
    cell(p([val(D.color)]), 3200),
    cell(p([t("Тип ЛКП", 16)]), 1800),
    cell(p([val(D.paint, 15)]), 3560),
  )]),
  tbl([1800,7560], [
    row(cell(p([t("Принадлежащего", 16)]), 1800), cell(p([val(D.owner)]), 7560)),
    row(cell(p([t("", 12)]), 1800), cell(p([t("(Ф.И.О. владельца, почт. адрес, телефон)", 12)], AlignmentType.CENTER), 7560)),
    row(cell(p([t("Довер. лицо", 16)]), 1800), cell(p([val(D.proxy || "Нет")]), 7560)),
    row(cell(p([t("", 12)]), 1800), cell(p([t("(Ф.И.О. доверенного лица, почт. адрес, телефон)", 12)], AlignmentType.CENTER), 7560)),
  ]),
  p([t("Комплектация", 16, true)], AlignmentType.CENTER),
  tbl([2340,2340,2340,2340], [
    ...komplRows,
    row(cell(p([t("Прочее: ", 15), val(KOMPL.other || "", 15)]), 9360, {cs:4})),
    row(cell(p([t("Марка и размер шин: ", 15), val(D.tires, 15)]), 4680, {cs:2}), cell(p([t("Глубина протектора, мм: ", 15), val(D.tread, 15)]), 4680, {cs:2})),
  ]),
  sp(),

  p([t("Учетные данные ТС соответствуют фактическим — ", 15), t(ck(UCHET_OK), 16, UCHET_OK), t("      Учетные данные ТС не соответствуют фактическим — ", 15), t(ck(!UCHET_OK), 16, !UCHET_OK)]),
  sp(),

  tbl([4680,4680], [
    row(cell(p([t("Дата наступления страхового события", 15, true)]), 4680), cell(p([t("Тех. состояние АМТС либо его остатков", 15)]), 4680)),
    row(cell(p([val(D.eventDate, 15)]), 4680), cell(p([val(D.techState, 15)]), 4680)),
  ]),
  tbl([9360], [
    row(cell([p([t("АМТС имеет дефекты эксплуатации (коррозия, износ шин и т.п.)", 15)]), ...listOrNo(D.defects)], 9360)),
    row(cell([p([t("АМТС имеет повреждения, не относящиеся к данному случаю", 15)]), ...listOrNo(D.otherDamage)], 9360)),
  ]),
  sp(),

  p([t("При осмотре установлено:", 16, true)], AlignmentType.CENTER),
  // одна таблица ровно по числу повреждений, сама перетекает на след. лист
  tbl(cW, [damHeader(), ...damages.map((dm, i) => damRow(i + 1, dm))]),
];

// =================== ХВОСТ АКТА (одним блоком) ===================
// Крупнее, чем таблицы на 1 листе, чтобы легко читалось
const TS = 19, TS2 = 17;
const page2 = [
  spk(),
  pk([t("Примечания: ", TS, true), t(D.notes || "Нет", TS, true)]),
  spk(),

  pk([t("Акт составлен:  ", TS), ...choice(["по наружному осмотру", "по осмотру на СТО"], ACT_PLACE === "outside" ? "по наружному осмотру" : "по осмотру на СТО", TS)]),
  pk([t("Условия осмотра:  ", TS, true),
      ...choice(["естественное/светлое","искусственное","сумерки"], {day:"естественное/светлое",art:"искусственное",dusk:"сумерки"}[LIGHT], TS)]),
  pk([t("Чистота ТС:  ", TS, true), ...choice(["чистое","грязное"], {clean:"чистое",dirty:"грязное"}[CLEAN], TS),
      t("        Погода:  ", TS, true), ...choice(["ясно","пасмурно","осадки"], {clear:"ясно",cloudy:"пасмурно",rain:"осадки"}[WEATHER], TS)]),
  spk(),

  pk([t("Все перечисленные повреждения относятся:", TS, true)]),
  pk(choice(["К одному страховому событию", "К нескольким страховым событиям"], {one:"К одному страховому событию",many:"К нескольким страховым событиям"}[EVENTS], TS)),
  pk([t(ck(EVENTS==="study") + " Окончательное установление возможности/невозможности получения повреждений при заявленных обстоятельствах будет осуществлено в рамках дальнейшего исследования", TS, EVENTS==="study")]),
  spk(),

  pk([t("В ходе осмотра экспертом выявлены и зафиксированы все возможные и видимые повреждения ТС. Владелец ТС/Доверенное лицо уведомлен о необходимости обращения в Страховую Компанию для фиксации любых скрытых повреждений, выявленных в ходе самостоятельно организованного ремонта ТС или его дефектовки. Владельцу/Доверенному лицу предложено проведение дефектовки ТС в условиях СТОА по направлению Страховщика.", TS2)]),
  pk([t("В случае выявления скрытых повреждений, неуказанных в настоящем акте, обязуюсь предоставить страховщику возможность повторного осмотра ТС для определения объема таких повреждений.", TS2)]),
  spk(),

  pk([t("Информация о пробах и элементах ТС, взятых для исследования: ", TS), t("пробы не брались", TS, true)]),
  spk(),

  pk(choice(["ТС подлежит ремонту", "Ремонт ТС не рентабелен", "Для определения рентабельности ремонта требуется предварительный расчет"],
            {yes:"ТС подлежит ремонту",no:"Ремонт ТС не рентабелен",calc:"Для определения рентабельности ремонта требуется предварительный расчет"}[REPAIR], TS2)),
  spk(),

  pk([t("При осмотре присутствовали, с изложенным согласны:", TS, true)]),
  sigLines([{ label: "Владелец АМТС (доверенное лицо)", fio: D.proxy || D.owner, note: OWNER_NOTE }], TS, [3700, 3500, 3300]),
  pk([t("Другие заинтересованные лица:   ", TS), ...choice(["присутствовали", "не присутствовали"], OTHERS === "present" ? "присутствовали" : "не присутствовали", TS)]),
  spk(),
  sigLines([{ label: "Эксперт – автотехник", fio: EXPERT, expert: true }], TS, [3700, 3500, 3300]),
  spk(),
  pk([t("Проведена фотофиксация:   ", TS), ...choice(["да", "нет"], PHOTO ? "да" : "нет", TS), t("          Дата составления акта осмотра ", TS), val(D.date, TS)]),
  spk(),
  p([t("Особое мнение: ", TS), t(D.specialNote || "Нет", TS, true)]),
];

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 560, right: 600, bottom: 560, left: 800 }
      }
    },
    children: [...page1, ...page2]
  }]
});

return Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT_FILE, buf);
  console.log("OK");
});

};
