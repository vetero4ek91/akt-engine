const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, BorderStyle,
  WidthType, VerticalAlign, VerticalMergeType, HeightRule, ShadingType, PageOrientation,
  TableLayoutType, UnderlineType, PageNumber, ImageRun, HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom, TextWrappingType
} = require('docx');
const fs = require('fs');

// Движок «Акт о гибели, повреждении или утрате строений (квартир), домашнего и/или другого
// имущества» (ООО «Фаворит» / Росгосстрах, имущество физлиц). Пустая форма:
// данные, ФИО эксперта и подпись передаются снаружи.
//
// Листы (порядок вывода):
//   титул — всегда
//   D.buildings[] → «СТРОЕНИЕ», по листу на постройку (дом)
//   D.scheme      → «ОБЩАЯ СХЕМА МЕСТА ОСМОТРА» (дом — всегда)
//   D.flat        → «КВАРТИРА» (опционально, по умолчанию не выводится)
//   D.items[] / D.itemsTotal → «ДОМАШНЕЕ и/или другое ИМУЩЕСТВО»
//   D.rooms[]     → «ВНУТРЕННЯЯ ОТДЕЛКА И ИНЖЕНЕРНОЕ ОБОРУДОВАНИЕ», по листу на помещение
module.exports = function build(cfg) {
  const { D, SIGN_EXPERT, EXPERT, SIG_RLE, OUT_FILE } = cfg;

  // ---------- подпись (как в остальных движках) ----------
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
  // подпись: SIG_RLE (как в остальных движках) или готовый PNG (SIG_PNG_FILE)
  const SIG_ON = !!(SIGN_EXPERT && (SIG_RLE || cfg.SIG_PNG_FILE));
  const SIG_PNG = !SIG_ON ? null : (SIG_RLE ? sigPng() : fs.readFileSync(cfg.SIG_PNG_FILE));
  // графа подписи клиента: по умолчанию «С актом ознакомлен, акт не подписан»;
  // D.ownerNote = "" — пусто (подписывает ручкой), либо любой свой текст («Отказ» и т.п.)
  const OWNER_NOTE = D.ownerNote === undefined ? "С актом ознакомлен, акт не подписан" : D.ownerNote;
  const sigImg = () => new ImageRun({
    type: "png", data: SIG_PNG, transformation: { width: 58, height: 42 },
    floating: {
      horizontalPosition: { relative: HorizontalPositionRelativeFrom.CHARACTER, offset: 0 },
      verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: -120000 },
      wrap: { type: TextWrappingType.NONE }, allowOverlap: true, layoutInCell: true, zIndex: 10,
    },
  });

  // ---------- базовые помощники ----------
  const F = "Times New Roman";
  const GREEN = "D7E4BD";
  const t = (text, o = {}) => new TextRun({ text: String(text ?? ""), size: o.sz || 18, bold: !!o.b, italics: !!o.i, font: F,
    color: o.color, underline: o.u ? { type: UnderlineType.SINGLE } : undefined });
  const P = (runs, o = {}) => new Paragraph({ alignment: o.a || AlignmentType.LEFT,
    spacing: { before: o.before || 0, after: o.after || 0, line: o.line || 240 },
    border: o.border, children: Array.isArray(runs) ? runs : [runs] });
  const L = AlignmentType.LEFT, C = AlignmentType.CENTER, R = AlignmentType.RIGHT;
  const S1 = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const S2 = { style: BorderStyle.SINGLE, size: 10, color: "000000" };
  const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const DT = { style: BorderStyle.DOTTED, size: 4, color: "000000" };
  const GD = { style: BorderStyle.DOTTED, size: 2, color: "8C8C8C" };
  const ALL = { top: S1, bottom: S1, left: S1, right: S1 };
  const NONE = { top: NB, bottom: NB, left: NB, right: NB };
  const bd = (top, right, bottom, left) => ({ top, right, bottom, left });
  const cell = (children, o = {}) => new TableCell({
    children: (Array.isArray(children) ? children : [children]).map(x => typeof x === "string" ? P(t(x, o.r || {}), { a: o.a }) : x),
    width: o.w ? { size: o.w, type: WidthType.DXA } : undefined,
    columnSpan: o.span, verticalMerge: o.vm, borders: o.bd || ALL,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: o.va || VerticalAlign.CENTER,
    margins: o.m || { top: 10, bottom: 10, left: 60, right: 60 },
  });
  const table = (cols, rows, o = {}) => new Table({ width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: cols, layout: TableLayoutType.FIXED, rows, borders: o.borders });
  const row = (children, o = {}) => new TableRow({ children, cantSplit: true, tableHeader: !!o.header,
    height: o.h ? { value: o.h, rule: o.exact ? HeightRule.EXACT : HeightRule.ATLEAST } : undefined });
  const ck = (on) => on ? "☒" : "☐";
  const nz = (x, d = "") => (x === undefined || x === null || x === "") ? d : String(x);
  const ND = "не повреждено";

  // поле «подпись: значение» с точечной линией
  const field = (label, val, lw, vw, o = {}) => [
    cell(P([t(label, { b: true, sz: o.sz || 20 }), ...(o.hint ? [t(" " + o.hint, { sz: 11, b: true })] : [])]), { w: lw, bd: NONE, va: VerticalAlign.BOTTOM, span: o.lspan }),
    cell(P(t(nz(val), { sz: o.sz || 20 }), { a: C }), { w: vw, bd: bd(NB, NB, DT, NB), va: VerticalAlign.BOTTOM, span: o.vspan }),
  ];

  // клетки-«боксы» для даты/номера
  function boxes(str, groups, cw = 260) {
    const s = String(str || "").replace(/[^0-9A-Za-zА-Яа-я\-\/]/g, "");
    const cells = [], cols = [];
    let i = 0;
    groups.forEach((g, gi) => {
      if (gi) { cells.push(cell([P(t(""))], { w: 70, bd: NONE })); cols.push(70); }
      for (let k = 0; k < g; k++) {
        cells.push(cell([P(t(s[i++] || "", { b: true, sz: 22 }), { a: C })], { w: cw, bd: bd(S2, S2, S2, S2), m: { top: 0, bottom: 0, left: 0, right: 0 } }));
        cols.push(cw);
      }
    });
    return table(cols, [row(cells, { h: 320, exact: true })]);
  }
  const dateGroups = d => boxes(String(d || "").replace(/\./g, ""), [2, 2, 4]);
  const timeGroups = tm => boxes(String(tm || "").replace(/[:.]/g, ""), [2, 2]);
  // строка: ячейки без рамок, внутри которых текст/таблицы
  const layout = (cols, contents, o = {}) => table(cols, [row(contents.map((c, i) =>
    cell(Array.isArray(c) ? c : [c], { w: cols[i], bd: NONE, va: o.va || VerticalAlign.CENTER, m: { top: 0, bottom: 0, left: 40, right: 40 } })), { h: o.h })]);
  const hint = (txt) => P(t(txt, { sz: 13 }), { a: C });
  const lineVal = (val, o = {}) => P(t(nz(val, " "), { b: true, sz: 22 }), { border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000", space: 1 } }, before: o.before || 60 });

  // блок подписей внизу листа приложения
  function signBlock(totalW) {
    const lw = Math.round(totalW * 0.26), sw = Math.round(totalW * 0.17), dw = Math.round(totalW * 0.13), fw = totalW - lw - sw - dw;
    const gray = (x) => t(x, { i: true, sz: 16, color: "808080" });
    const sigP = SIG_ON ? P([sigImg()], { a: C }) : P(gray("подпись"), { a: C });
    const part = D.participant || {};
    return table([lw, fw, sw, dw], [
      row([cell(P(t("Осмотр произвел", { b: true, sz: 22 })), { w: lw, fill: GREEN }),
           cell(P(EXPERT ? t(EXPERT, { b: true, sz: 20 }) : gray("ФАМИЛИЯ ИМЯ ОТЧЕСТВО"), { a: C }), { w: fw }),
           cell(sigP, { w: sw }),
           cell(P(D.date ? t(D.date, { sz: 20 }) : gray("дата"), { a: C }), { w: dw })], { h: 400 }),
      row([cell(P(t("Выгодоприобретатель (доверенное лицо)", { b: true, sz: 22 })), { w: lw, fill: GREEN }),
           cell(P(part.fio ? t(part.fio, { b: true, sz: 20 }) : gray("ФАМИЛИЯ ИМЯ ОТЧЕСТВО"), { a: C }), { w: fw }),
           cell(P(OWNER_NOTE ? t(OWNER_NOTE, { sz: 14, b: true }) : gray("подпись"), { a: C }), { w: sw }),
           cell(P(D.date ? t(D.date, { sz: 20 }) : gray("дата"), { a: C }), { w: dw })], { h: 400 }),
    ]);
  }

  // таблица «элемент → несколько строк с точечными разделителями»
  // spec: { cols:[...], lead:(el, i) => [cells for vmerge-columns], lines: [[c1,c2,...]], after: [cells vmerged right] }
  function elementRows(cols, items, nLeadCols, nLineCols, trailCols = 0, rowH = 235) {
    const out = [];
    items.forEach((it, ii) => {
      const n = it.lines.length;
      it.lines.forEach((ln, k) => {
        const cells = [];
        const first = k === 0, last = k === n - 1;
        it.lead.forEach((content, ci) => cells.push(cell(first ? content : [P(t(""))],
          { w: cols[ci], vm: n > 1 ? (first ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE) : undefined, va: VerticalAlign.TOP })));
        ln.forEach((txt, li) => cells.push(cell([P(t(nz(txt), { sz: 17 }), { a: li === ln.length - 1 && nLineCols > 1 ? C : L })],
          { w: cols[nLeadCols + li], bd: bd(first ? S1 : NB, S1, last ? S1 : DT, S1), m: { top: 0, bottom: 0, left: 60, right: 60 } })));
        (it.trail || []).forEach((tr, ti) => {
          if (tr === null) return; // ячейку ведёт объединение уровнем выше
          cells.push(cell(first ? tr : [P(t(""))], { w: cols[nLeadCols + nLineCols + ti],
            vm: n > 1 ? (first ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE) : undefined }));
        });
        out.push({ cells, h: rowH });
      });
    });
    return out;
  }

  // ---------- сетка со схемой (общая схема / план помещения) ----------
  // shapes: {x,y,w,h, text, border:true|false, sz, b} в клетках
  function grid(cols, rows, cw, rh, shapes) {
    const own = Array.from({ length: rows }, () => Array(cols).fill(-1));
    const ok = [];
    const fits = (x, y, w, h) => { if (x < 0 || y < 0 || x + w > cols || y + h > rows) return false;
      for (let r = y; r < y + h; r++) for (let c = x; c < x + w; c++) if (own[r][c] >= 0) return false; return true; };
    shapes.forEach((s0, si) => {
      for (const s of [s0, ...(s0.alt || [])]) {
        const x = Math.round(s.x), y = Math.round(s.y), w = Math.max(1, Math.round(s.w)), h = Math.max(1, Math.round(s.h));
        if (!fits(x, y, w, h)) continue;
        for (let r = y; r < y + h; r++) for (let c = x; c < x + w; c++) own[r][c] = si;
        ok[si] = { ...s0, x, y, w, h }; return;
      }
    });
    const trs = [];
    for (let r = 0; r < rows; r++) {
      const cells = [];
      for (let c = 0; c < cols;) {
        const si = own[r][c];
        if (si < 0) { cells.push(new TableCell({ children: [new Paragraph({ children: [] })], width: { size: cw, type: WidthType.DXA },
          borders: { top: GD, bottom: GD, left: GD, right: GD }, margins: { top: 0, bottom: 0, left: 0, right: 0 } })); c++; continue; }
        const s = ok[si];
        const B = s.border === false ? GD : S2;
        const first = r === s.y;
        const txt = first || s.h === 1 ? s.text : "";
        cells.push(new TableCell({
          children: [new Paragraph({ alignment: s.a || C, spacing: { before: 0, after: 0, line: 200 }, children: txt ? String(txt).split("\n").flatMap((ln, i) => i ? [new TextRun({ break: 1, text: ln, size: s.sz || 14, bold: s.b !== false, font: F })] : [new TextRun({ text: ln, size: s.sz || 14, bold: s.b !== false, font: F })]) : [] })],
          width: { size: cw * s.w, type: WidthType.DXA }, columnSpan: s.w > 1 ? s.w : undefined,
          verticalMerge: s.h > 1 ? (first ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE) : undefined,
          verticalAlign: VerticalAlign.CENTER,
          borders: { top: B, bottom: B, left: B, right: B }, margins: { top: 0, bottom: 0, left: 0, right: 0 },
        }));
        c += s.w;
      }
      trs.push(new TableRow({ children: cells, height: { value: rh, rule: HeightRule.EXACT } }));
    }
    return table(Array(cols).fill(cw), trs);
  }
  const fmt = (n) => String(Math.round(n * 100) / 100).replace(".", ",");
  // объекты в метрах → фигуры в клетках с подписями размеров
  // objs: [{name, x, y, w, h}] (w — вдоль листа, h — вниз; метры), extra: [{text, x, y, w}] в метрах
  function placeObjects(objs, cols, rows, opt = {}) {
    if (!objs || !objs.length) return [];
    const all = [...objs, ...(opt.extra || []).map(e => ({ x: e.x, y: e.y, w: e.w || 4, h: 0.5 }))];
    const minX = Math.min(...all.map(o => o.x || 0)), minY = Math.min(...all.map(o => o.y || 0));
    const maxX = Math.max(...all.map(o => (o.x || 0) + o.w)), maxY = Math.max(...all.map(o => (o.y || 0) + o.h));
    const padL = 1, padT = 2, padR = opt.dimSide ? 4 : 1, padB = 1;
    const sc = opt.scale || Math.min((cols - padL - padR) / (maxX - minX), (rows - padT - padB) / (maxY - minY));
    const shapes = [], labels = [];
    const X = v => padL + (v - minX) * sc, Y = v => padT + (v - minY) * sc;
    objs.forEach(o => {
      const x = Math.round(X(o.x || 0)), y = Math.round(Y(o.y || 0));
      const w = Math.max(1, Math.round(o.w * sc)), h = Math.max(1, Math.round(o.h * sc));
      const label = o.name + (opt.dimsInside !== false && !opt.dimSide ? `\n${fmt(o.w)}×${fmt(o.h)} м` : "");
      shapes.push({ x, y, w, h, text: label, sz: opt.sz || 14 });
      if (opt.dimSide) {
        const my = y + Math.floor((h - 1) / 2);
        labels.push({ x, y: y - 1, w, h: 1, text: `${fmt(o.w)} м`, border: false, sz: 14, alt: [{ x, y: y + h, w, h: 1 }] });
        labels.push({ x: x + w, y: my, w: 3, h: 1, text: `${fmt(o.h)} м`, border: false, sz: 14,
          alt: [{ x: x - 3, y: my, w: 3, h: 1 }, { x: x + w, y: my, w: 2, h: 1 }, { x: x - 2, y: my, w: 2, h: 1 }] });
      }
    });
    shapes.push(...labels);
    (opt.extra || []).forEach(e => shapes.push({ x: Math.round(X(e.x)), y: Math.round(Y(e.y)), w: Math.max(1, Math.round((e.w || 4) * sc)), h: 1, text: e.text, border: false, sz: 14 }));
    return shapes;
  }

  // ---------- размеры страниц ----------
  const PORT = { width: 11906, height: 16838 };
  const LAND_W = 16838 - 2 * 600; // ширина полезной области на альбомном листе
  const PORT_W = 11906 - 1134 - 567;
  const landProps = { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 500, bottom: 400, left: 600, right: 600 } } };
  const portProps = { page: { size: PORT, margin: { top: 600, bottom: 500, left: 1134, right: 567 } } };
  const sections = [];

  // =================== ТИТУЛ ===================
  {
    const part = D.participant || {};
    const W = PORT_W;
    const causes = [
      ["pozhar", "Пожар"], ["vzryv", "Взрыв"],
      ["avaria", "Повреждение водой в результате аварии систем водоснабжения, отопления, канализации и пожаротушения"],
      ["sosedi", "Проникновение воды из соседних помещений"], ["krazha", "Кража отдельных элементов строения (квартиры)"], ["grabezh", "Грабеж"],
    ];
    const causes2 = [["razboy", "Разбой"], ["umysl", "Умышленное уничтожение (повреждение) строений (квартир) другими лицами."], ["stihiya", "Стихийные бедствия"], ["inoe", "Иное"]];
    const cz = (k, txt) => P([t(ck(D.cause === k) + " ", { sz: 22, b: D.cause === k }), t(txt, { sz: 17, b: D.cause === k })], { after: 50 });
    const inoeLines = [0, 1, 2].map(i => row([cell(P(t(i === 0 && D.cause === "inoe" ? nz(D.causeOther) : "", { b: true, sz: 17 })), { w: 4200 })], { h: 240 }));
    const half = Math.floor(W / 2);
    const body = [
      P(t("ООО «ФАВОРИТ»", { b: true, sz: 40 }), { a: C, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "7030A0", space: 4 } }, after: 160 }),
      layout([1400, 3200, 4300, W - 8900], [
        P(t("АКТ №", { b: true, sz: 26 }), { a: R }),
        boxes(D.actNo, [Math.max(9, String(D.actNo || "").length)], 300),
        dateGroups(D.date),
        P(t("от", { b: true, sz: 26 })),
      ]),
      layout([W - 3500, 1600, 1900], [P(t("")), P(t("Время осмотра", { b: true, sz: 22 }), { a: R }), timeGroups(D.time)]),
      P([t("О гибели, повреждении или утрате строений (квартир), домашнего и/или другого имущества принадлежащих", { b: true, sz: 22 })], { before: 120 }),
      lineVal(D.owner), hint("(Ф.И.О., наименование)"),
      P([t("и находящихся по адресу", { b: true, sz: 22 })], { before: 60 }),
      lineVal(D.address, { before: 0 }), hint("(наименование населенного пункта, волость, область, район, улица, дом, владение)"),
      layout([1700, 3000, W - 4700], [P(t("Экспертом", { b: true, sz: 22 })), dateGroups(D.date), lineVal(EXPERT, { before: 0 })], { va: VerticalAlign.BOTTOM }),
      hint("(Должность, Ф.И.О.)"),
      P(t("при участии Выгодоприобретателя, а также в отсутствие Выгодоприобретателя – совершеннолетнего члена его семьи (доверенного лица)", { b: true, sz: 22 }), { before: 60 }),
      lineVal(part.fio ? part.fio + (part.role ? `, ${part.role}` : "") : ""), hint("(Ф.И.О.)"),
      layout([3200, W - 3200], [P(t("представителя организации", { b: true, sz: 22 })), lineVal(D.orgRep || "—", { before: 0 })], { va: VerticalAlign.BOTTOM }),
      hint("(Наименование, Ф.И.О., должность)"),
      P(t(""), { after: 60 }),
      layout([3300, 3000, W - 6300], [P(t("на основании заявления от", { b: true, sz: 22 })), dateGroups(D.claimDate), P(t("составлен настоящий акт", { b: true, sz: 22 }))]),
      P(t(""), { after: 100 }),
      table([W], [row([cell(P(t("I. Общие сведения о произошедшем событии", { b: true, sz: 22 })), { w: W, bd: NONE, fill: "D9D9D9" })])]),
      P(t(""), { after: 100 }),
      table([half, W - half], [row([
        cell([P(t("Причина наступления события", { b: true, sz: 19 }), { after: 120 }), ...causes.map(([k, x]) => cz(k, x))], { w: half, va: VerticalAlign.TOP }),
        cell([P(t(""), { after: 60 }), ...causes2.map(([k, x]) => cz(k, x)), table([4200], inoeLines), P(t(""))], { w: W - half, va: VerticalAlign.TOP }),
      ])]),
      P(t(""), { after: 100 }),
      table([W], [row([cell([
        P(t("Время и дата наступления события", { b: true, sz: 19 })),
        layout([3000, 1000, 2500, W - 6500], [
          [P(t("Дата", { b: true, sz: 19 })), dateGroups(D.eventDate)], P(t("")),
          [P(t("Время", { b: true, sz: 19 })), timeGroups(D.eventTime)], P(t("")),
        ], { va: VerticalAlign.TOP }),
        P(t(""), { after: 60 }),
      ], { w: W })])]),
      P(t(""), { after: 100 }),
      table([2600, W - 2600], [
        row([cell(P(t("Характеристика события", { b: true, sz: 19 })), { w: 2600, va: VerticalAlign.TOP }), cell(P(t(nz(D.eventDescr), { sz: 19, b: true })), { w: W - 2600 })], { h: 280 }),
        ...(D.eventDescr && D.eventDescr.length > 90 ? [] : [row([cell(P(t("")), { w: W, span: 2 })], { h: 280 })]),
        row([cell(P(t("Меры, предпринятые по сохранности, спасению застрахованного имущества, а также по уменьшению ущерба", { b: true, sz: 19 })), { w: W, span: 2 })]),
        row([cell(P(t(nz(D.measures, "Не предпринимались"), { sz: 19, b: true })), { w: W, span: 2 })], { h: 280 }),
        row([cell(P(t("")), { w: W, span: 2 })], { h: 280 }),
      ]),
    ];
    sections.push({ properties: portProps, children: body });
  }

  // =================== СТРОЕНИЕ ===================
  const BLD_EL = [
    ["fundament", "Фундамент", "(надземная часть, подземная часть, наружная отделка надземной части, отмостка)", 3],
    ["steny", "Стены", "(стены, утеплитель, перегородки)", 3],
    ["naruzh", "Наружная отделка", "", 2],
    ["poly", "Полы", "(столбы, доски, лаги, гидро-, паро-, теплоизоляция/ утепление)", 3],
    ["perekr", "Перекрытие", "(балки, доски, лаги, гидро-, паро-, теплоизоляция/утепление)", 3],
    ["krysha", "Крыша", "(стропила, обрешетка, гидро-, паро-, теплоизоляция, утепление, кровля, карниз с обшивкой и снегозадержатели, водостоки, фронтон)", 4],
    ["prochee", "Прочее", "(крыльцо, лестницы, балкон, ступеньки)", 3],
  ];
  // элемент: { mat, dmg: [[описание, размер], ...], pct: "3%" } ; нет элемента → «не повреждено»
  function dmgLines(el, minN, pctLabel) {
    const lines = [];
    if (!el || ((!el.dmg || !el.dmg.length) && !el.pct)) lines.push([el && el.note ? el.note : ND, el && el.note ? "" : ND]);
    else {
      (el.dmg || []).forEach(d => lines.push([d[0], d[1]]));
      if (el.pct) lines.push([pctLabel, el.pct]);
    }
    while (lines.length < minN) lines.push(["", ""]);
    return lines;
  }
  const elName = (name, h) => [P(t(name, { b: true, sz: 20 })), ...(h ? [P(t(h, { i: true, sz: 15 }))] : [])];
  const P8_LABEL = (pfx) => pfx;
  function p8Row(p, W, n, nCols) {
    p = p || "none";
    const none = p === "none", part = p === "partial";
    const ot = typeof p === "object" ? p.otdelka : null, inz = typeof p === "object" ? p.inzh : null;
    const r = (x, b) => t(x, { sz: 19, b });
    return row([
      cell(P(t(String(n) + ".", { b: true, sz: 20 })), { w: 500, va: VerticalAlign.TOP }),
      cell([
        P([r("Внутренняя отделка (вкл. окна и двери) и инженерное оборудование:   ", true), r(ck(none) + " не повреждены", none), r("            "), r(ck(part) + " повреждены частично, описание в приложении", part)]),
        P([r(ck(!!ot) + " внутренняя отделка (вкл. окна и двери) значительно повреждена, общий процент обесценения составляет ", !!ot), t(ot ? ` ${ot} ` : "____________", { b: true, sz: 19, u: !!ot })]),
        P([r(ck(!!inz) + " инженерное оборудование значительно повреждено, общий процент обесценения составляет ", !!inz), t(inz ? ` ${inz} ` : "____________", { b: true, sz: 19, u: !!inz })]),
      ], { w: W - 500, span: nCols - 1 }),
    ]);
  }
  (D.buildings || []).forEach(b => {
    const W = LAND_W;
    const cols = [500, 3000, 1900, W - 500 - 3000 - 1900 - 3400, 3400];
    const items = BLD_EL.map(([k, name, h, minN], i) => {
      const el = (b.el || {})[k];
      return {
        lead: [[P(t(`${i + 1}.`, { b: true, sz: 20 }))], elName(name, h), [P(t(nz(el && el.mat), { sz: 18, b: true }), { a: C })]],
        lines: dmgLines(el, minN, `Общий процент повреждения элемента «${name.toLowerCase()}» в целом`),
      };
    });
    const er = elementRows(cols, items, 3, 2).map(x => row(x.cells, { h: x.h }));
    const hdr = row([
      cell(P(t("№", { b: true, sz: 20 }), { a: C }), { w: cols[0], fill: GREEN }),
      cell(P(t("Элементы строения", { b: true, sz: 20 }), { a: C }), { w: cols[1], fill: GREEN }),
      cell([P(t("Материал", { b: true, sz: 20 }), { a: C }), P(t("(фиксируется основной материал элемента)", { b: true, sz: 13 }), { a: C })], { w: cols[2], fill: GREEN }),
      cell(P(t("Описание повреждения", { b: true, sz: 20 }), { a: C }), { w: cols[3], fill: GREEN }),
      cell([P(t("Размер повреждения", { b: true, sz: 20 }), { a: C }), P(t("(в м2, м3, м/п, шт., либо % от общей площади)", { b: true, sz: 13 }), { a: C })], { w: cols[4], fill: GREEN }),
    ], { header: true });
    const hc = [3300, 3600, 3300, 1500, 2300, W - 14000];
    const head = table(hc, [
      row([...field("Наименование постройки (тип)", b.type, hc[0], hc[1]),
           ...field("Наименование объекта, к которому пристроено", b.attachedTo || "—", hc[2] + hc[3], hc[4] + hc[5], { lspan: 2, vspan: 2, hint: "(основное строение / доп. строение)" })]),
      row([...field("Хозяйственное назначение", b.purpose, hc[0], hc[1]), ...field("Количество дверей, вкл. входную", b.doors, hc[2], hc[3]), ...field("Количество окон", b.windows, hc[4], hc[5])]),
      row([...field("Размер (длина*ширина*высота)", b.size, hc[0], hc[1]), ...field("Год возведения", b.year, hc[2], hc[3]), ...field("Год последнего ремонта", b.repairYear, hc[4], hc[5])]),
    ]);
    sections.push({ properties: landProps, children: [
      P(t("СТРОЕНИЕ", { b: true, sz: 24 }), { a: R }), head, P(t(""), { after: 80 }),
      table(cols, [hdr, ...er, p8Row(b.p8, W, 8, 5)]),
      P(t(""), { after: 60 }), signBlock(W),
    ] });
  });

  // =================== ОБЩАЯ СХЕМА ===================
  if (D.scheme) {
    const W = PORT_W, cols = 40, cw = Math.floor(W / cols), rows = 54, rh = cw;
    const shapes = D.scheme.shapes || placeObjects(D.scheme.objects, cols, rows, { dimSide: true, extra: D.scheme.labels, sz: 15 });
    sections.push({ properties: portProps, children: [
      P(t("ОБЩАЯ СХЕМА МЕСТА ОСМОТРА", { b: true, sz: 24 }), { a: R, after: 120 }),
      grid(cols, rows, cw, rh, shapes),
      P(t(""), { after: 120 }), signBlock(W),
    ] });
  }

  // =================== КВАРТИРА (опционально) ===================
  if (D.flat) {
    const fl = D.flat, W = LAND_W;
    const cols = [500, 2500, 1900, W - 500 - 2500 - 1900 - 2400 - 2600, 2400, 2600];
    const EL = [["steny", "Стены"], ["peregorodki", "Перегородки"], ["potolok", "Потолочное перекрытие"], ["pol", "Половое перекрытие"]];
    const items = EL.map(([k, name], i) => {
      const el = (fl.el || {})[k];
      return { lead: [[P(t(`${i + 1}.`, { b: true, sz: 20 }))], [P(t(name, { b: true, sz: 20 }))], [P(t(nz(el && el.mat), { sz: 18, b: true }), { a: C })]],
        lines: dmgLines(el, 3, `Общий процент повреждения элемента «${name.toLowerCase()}» в целом`) };
    });
    const total = items.reduce((a, it) => a + it.lines.length, 0);
    let rIdx = 0;
    const er = elementRows(cols.slice(0, 5), items, 3, 2).map(x => {
      const c = [...x.cells, cell(rIdx === 0 ? [P(t(nz(fl.totalPct, "0%"), { b: true, sz: 24 }), { a: C })] : [P(t(""))],
        { w: cols[5], vm: rIdx === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE })];
      rIdx++; return row(c, { h: x.h });
    });
    const G = (x, s) => cell([P(t(x, { b: true, sz: 20 }), { a: C }), ...(s ? [P(t(s, { b: true, sz: 13 }), { a: C })] : [])], { fill: GREEN });
    const hdr = row([G("№"), G("Элементы квартиры"), G("Материал", "(фиксируется основной материал элемента)"), G("Описание повреждения"),
      G("Размер повреждения", "(в м2, м3, м/п, шт., либо % от общей площади)"), G("Общий % повреждения конструктивных элементов квартиры")], { header: true });
    const hc = [2600, 3300, 3400, 2000, W - 11300];
    const head = table(hc, [
      row([...field("Общая площадь", fl.area, hc[0], hc[1]), cell(P(t("")), { w: hc[2] + hc[3] + hc[4], span: 3, bd: NONE })]),
      row([...field("Количество комнат", fl.rooms, hc[0], hc[1]), ...field("Количество дверей, вкл. входную", fl.doors, hc[2], hc[3]), cell(P(t("")), { w: hc[4], bd: NONE })]),
      row([...field("Год возведения", fl.year, hc[0], hc[1]), ...field("Количество окон", fl.windows, hc[2], hc[3]), cell(P(t("")), { w: hc[4], bd: NONE })]),
      row([...field("Год последнего ремонта", fl.repairYear, hc[0], hc[1]), cell(P(t("")), { w: hc[2] + hc[3] + hc[4], span: 3, bd: NONE })]),
    ]);
    // общее имущество малоквартирного дома
    const oc = [500, 2600, 1800, W - 500 - 2600 - 1800 - 2000 - 1600 - 2000, 2000, 1600, 2000];
    const OEL = [["fundament", "Фундамент"], ["krysha", "Крыша"], ["naruzh", "Наружная отделка стен"]];
    const ob = fl.common || {};
    const oitems = OEL.map(([k, name], i) => {
      const el = ob[k];
      const lines = el ? (el.dmg || []).map(d => [d[0], d[1]]) : (fl.common ? [[ND, ND]] : []);
      while (lines.length < 2) lines.push(["", ""]);
      return { lead: [[P(t(`${i + 1}.`, { b: true, sz: 20 }))], [P(t(name, { b: true, sz: 20 }))], [P(t(nz(el && el.mat), { sz: 18, b: true }), { a: C })]],
        lines, trail: [[P(t(nz(el && el.pct), { b: true, sz: 20 }), { a: C })]] };
    });
    let oi = 0;
    const orows = elementRows(oc.slice(0, 6), oitems, 3, 2, 1).map(x => {
      const c = [...x.cells, cell(oi === 0 ? [P(t(nz(ob.totalPct), { b: true, sz: 22 }), { a: C })] : [P(t(""))], { w: oc[6], vm: oi === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE })];
      oi++; return row(c, { h: x.h });
    });
    const ohdr = row([G("№"), G("Общее имущество", "(при повреждении объекта квартира в малоквартирном доме)"), G("Материал", "(фиксируется основной материал элемента)"),
      G("Описание повреждения"), G("Размер повреждения", "(в м2, м3, м/п, шт., либо % от общей площади)"), G("% повреждения элемента"), G("Общий % повреждения")]);
    sections.push({ properties: landProps, children: [
      P(t("КВАРТИРА", { b: true, sz: 24 }), { a: R }), head, P(t(""), { after: 80 }),
      table(cols, [hdr, ...er, p8Row(fl.p5, W, 5, 6)]),
      P(t(""), { after: 80 }), table(oc, [ohdr, ...orows]),
      P(t(""), { after: 60 }), signBlock(W),
    ] });
  }

  // =================== ДОМАШНЕЕ ИМУЩЕСТВО ===================
  if ((D.items && D.items.length) || D.itemsTotal) {
    const W = LAND_W;
    const cols = [500, 4300, 1150, 1650, 2900, 1000, 900, W - 12400];
    const G = (x, o = {}) => cell([P(t(x, { b: true, sz: 16 }), { a: o.a || C })], { w: o.w, span: o.span, vm: o.vm, fill: GREEN, va: o.va || VerticalAlign.CENTER });
    const RS = VerticalMergeType.RESTART, CT = VerticalMergeType.CONTINUE;
    const reasons = ["причина не предъявления остатков:", "1. - несгораемые остатки имущества отсутствуют на месте пожара;",
      "2. - несгораемые остатки имущества под завалом конструкций / под пожарным мусором / под снежным покровом;",
      "3. - доступ в объект ограничен (опасность обрушения);", "4. - иное (расшифровать)"];
    const h1 = row([G("№ п/п", { w: cols[0], vm: RS }), G("Наименование и количество предметов с указанием размера, материала, цвета и т.п.", { w: cols[1], vm: RS }),
      G("Год приобретения", { w: cols[2], vm: RS }), G("Стоимость в новом состоянии с учетом действующих цен на дату наступления заявленного события", { w: cols[3], vm: RS }),
      G("Описание повреждений (разбито, поломано, украдено, …)", { w: cols[4], vm: RS }), G("% повреж-дения", { w: cols[5], vm: RS }),
      G("Отметка о предъявлении остатков", { w: cols[6] + cols[7], span: 2 })], { header: true });
    const h2 = row([...[0, 1, 2, 3, 4, 5].map(i => cell([P(t(""))], { w: cols[i], vm: CT, fill: GREEN })),
      G("да / нет", { w: cols[6] }), cell(reasons.map((x, i) => P(t(x, { b: i === 0, sz: 14 }))), { w: cols[7], fill: GREEN })], { header: true });
    const h3 = row(cols.map((w, i) => G(String(i + 1), { w })), { header: true });
    const r = (x, o = {}) => cell([P(t(nz(x), { sz: 17, b: o.b }), { a: o.a || C })], { w: o.w, span: o.span });
    let body;
    if (D.itemsTotal) {
      body = [row([r("1", { w: cols[0] }), r("Домашнее имущество уничтожено полностью, остатки предъявлены, годных остатков нет", { w: cols.slice(1, 6).reduce((a, b) => a + b), span: 5, b: true }),
        r("да", { w: cols[6] }), r("—", { w: cols[7] })], { h: 360 })];
    } else {
      body = D.items.map((it, i) => row([
        r(`${i + 1}.`, { w: cols[0] }), r(it.name, { w: cols[1], a: L }), r(it.year, { w: cols[2] }), r(it.price, { w: cols[3] }),
        r(it.descr, { w: cols[4], a: L }), r(it.pct, { w: cols[5] }),
        r(it.shown === false ? "нет" : "да", { w: cols[6] }), r(it.shown === false ? nz(it.reason, "4. - иное") : "—", { w: cols[7], a: it.shown === false ? L : C }),
      ], { h: 300 }));
    }
    sections.push({ properties: landProps, children: [
      P(t("ДОМАШНЕЕ и/или другое ИМУЩЕСТВО", { b: true, sz: 24 }), { a: R, after: 80 }),
      table(cols, [h1, h2, h3, ...body]),
      P(t(""), { after: 100 }), signBlock(W),
      new Paragraph({ alignment: R, spacing: { before: 60 }, children: [t("лист ", { sz: 18 }), new TextRun({ children: [PageNumber.CURRENT], size: 18, font: F }),
        t(" из ", { sz: 18 }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: F })] }),
    ] });
  }

  // =================== ВНУТРЕННЯЯ ОТДЕЛКА ===================
  const ROOM_EL = [["steny", "Стены", ""], ["potolok", "Потолок", ""], ["pol", "Пол", ""], ["okna", "Окна", "(в т.ч. откосы и подоконник)"],
    ["dveri", "Двери", "(в т.ч. коробка, полотно, наличники, фурнитура)"], ["inzh", "Инженерное оборудование", ""]];
  (D.rooms || []).forEach(rm => {
    const W = LAND_W;
    // шапка помещения + план-схема
    const iw = [4300, 4600], gapW = 250, planW = W - iw[0] - iw[1] - gapW;
    const info = [["Наименование помещения", rm.name], ["Наименование объекта (его части), в котором оно располагается", rm.location],
      ["Размер (длина*ширина*высота)", rm.size], ["Общая площадь (м2)", rm.area], ["Количество и размер дверных проемов (двери, арки и т.п.)", rm.doors],
      ["Количество и размер окон (в т.ч. технических проемов под вытяжку)", rm.windows], ["Количество и размер выступов, ниш, антресолей и т.п.", nz(rm.niches, "нет")],
      ["Количество и размер нестандартной отделки стен, пола, потолка (фальш-панели, плиточные фартуки и т.п.)", nz(rm.nonstd, "нет")], ["Год последнего ремонта", rm.repairYear]];
    const infoT = table(iw, info.map(([a, b]) => row([cell(P(t(a, { b: true, sz: 15 }), { line: 220 }), { w: iw[0], m: { top: 0, bottom: 0, left: 50, right: 50 } }), cell(P(t(nz(b), { sz: 17, b: true }), { line: 220 }), { w: iw[1], m: { top: 0, bottom: 0, left: 50, right: 50 } })])));
    const gc = 24, gcw = Math.floor(planW / gc), grows = 12, grh = 225;
    let pshapes = [];
    if (rm.plan && rm.plan.shapes) pshapes = rm.plan.shapes;
    else {
      const objs = rm.plan && rm.plan.objects ? rm.plan.objects : (() => {
        const m = String(rm.size || "").replace(/,/g, ".").split(/[*xх×]/).map(Number);
        return m[0] && m[1] ? [{ name: rm.name, x: 0, y: 0, w: m[0], h: m[1] }] : [];
      })();
      pshapes = placeObjects(objs, gc, grows, { dimSide: true, sz: 14 });
    }
    const planT = [P(t("План-схема помещения", { b: true, sz: 20 }), { a: R, after: 40 }), grid(gc, grows, gcw, grh, pshapes),
      P(t("(при сложной конфигурации помещения на схеме так же должны быть отмечены длины стен)", { i: true, sz: 15 }), { a: R })];
    const top = table([iw[0] + iw[1], gapW, planW], [row([
      cell([infoT], { w: iw[0] + iw[1], bd: NONE, va: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
      cell(P(t("")), { w: gapW, bd: NONE }),
      cell(planT, { w: planW, bd: NONE, va: VerticalAlign.TOP, m: { top: 0, bottom: 0, left: 0, right: 0 } }),
    ])]);
    const cols = [500, 3000, 2900, W - 500 - 3000 - 2900 - 2800, 2800];
    // элемент: { descr: "следы залития", items: [[материал, размер], ...] } ; нет → «не повреждено»
    const items = ROOM_EL.map(([k, name, h], i) => {
      const el = (rm.el || {})[k];
      const lines = el && el.items && el.items.length ? el.items.map(x => [x[0], x[1]]) : [["", el ? "" : ""]];
      while (lines.length < (k === "inzh" ? 4 : 3)) lines.push(["", ""]);
      return { lead: [[P(t(`${i + 1}.`, { b: true, sz: 20 }))], [P([t(name, { b: true, sz: 20 }), ...(h ? [t(" " + h, { sz: 15 })] : [])])],
        [P(t(el ? nz(el.descr) : ND, { sz: 17, b: !!el }), { a: C })]], lines };
    });
    const G = (x, s) => cell([P(t(x, { b: true, sz: 20 }), { a: C }), ...(s ? [P(t(s, { b: true, sz: 14 }), { a: C })] : [])], { fill: GREEN });
    const hdr = row([G("№"), G("Элементы"), G("Описание повреждения"), G("Наименование материала"), G("Размер повреждений", "(м2, м3, м/п, шт., %)")], { header: true });
    const er = elementRows(cols, items, 3, 2, 0, 205).map(x => row(x.cells, { h: x.h }));
    sections.push({ properties: landProps, children: [
      P(t("ВНУТРЕННЯЯ ОТДЕЛКА И ИНЖЕНЕРНОЕ ОБОРУДОВАНИЕ", { b: true, sz: 22 }), { a: R }),
      top, P(t(""), { after: 80 }), table(cols, [hdr, ...er]), P(t(""), { after: 40 }), signBlock(W),
    ] });
  });

  const doc = new Document({ styles: { default: { document: { run: { font: F } } } }, sections });
  return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT_FILE, buf); console.log("OK", OUT_FILE); });
};
