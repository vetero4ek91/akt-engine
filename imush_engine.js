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
//   D.scheme      → «ОБЩАЯ СХЕМА МЕСТА ОСМОТРА» (дом — всегда), сразу после титула; D.scheme.note — подпись под схемой
//   D.buildings[] → «СТРОЕНИЕ», по листу на постройку; b.total: true|"текст" — сгорело полностью (100% по всем элементам, п.8 100/100)
//   D.flat        → «КВАРТИРА» (опционально, по умолчанию не выводится)
//   D.items[] / D.itemsTotal → «ДОМАШНЕЕ и/или другое ИМУЩЕСТВО»
//   D.rooms[]     → «ВНУТРЕННЯЯ ОТДЕЛКА И ИНЖЕНЕРНОЕ ОБОРУДОВАНИЕ», по листу на помещение
//
// Значки на схемах (общая схема и план помещения), у любого объекта в scheme.objects / plan.objects:
//   openings: [{ type: "window"|"door", wall: "top"|"bottom"|"left"|"right", from: 0.5, w: 1.2,
//                open: "in"|"out", hinge: "start"|"end" }]   — from: м от левого (верх/низ) или верхнего (лев/прав) угла
//   stoves:   [{ x: 0.3, y: 0.3, w: 1, h: 1 }] или { x, y, w: 0.7, round: true } — м от левого верхнего угла объекта
//   hatch: true — весь объект заштрихован (сгорел/разрушен); zones: [{ x, y, w, h }] — зоны повреждения штриховкой
//   marks:    [{ x, y, type: "leak"|"fire", text: "прорыв подводки" }] — место протечки ⊙ / очаг пожара ⊗
// Для плана помещения без plan.objects (рисуется по rm.size) — rm.plan = { openings, stoves, zones, marks, hatch }.
// Общая схема: scheme.gaps: [{ a: "Жилой дом", b: "Баня", m: 6 }] — размерная линия между строениями (m не задано — по координатам);
//   scheme.north: true — стрелка «С» в правом верхнем углу. Под схемой/планом автоматически пишутся «Обозначения: ...».
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

  // ---------- значки на плане: окна, двери, печи ----------
  // PNG с прозрачным фоном, рисуется без внешних библиотек; items — в клетках сетки (см. placeObjects)
  function pngRGBA(W, H, px) {
    const zlib = require('zlib');
    const raw = Buffer.alloc((W * 4 + 1) * H);
    for (let y = 0; y < H; y++) px.copy(raw, y * (W * 4 + 1) + 1, y * W * 4, (y + 1) * W * 4);
    const T = []; for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; T[n] = c >>> 0; }
    const crc = b => { let c = 0xffffffff; for (const x of b) c = T[(c ^ x) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
    const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
    const ih = Buffer.alloc(13); ih.writeUInt32BE(W, 0); ih.writeUInt32BE(H, 4); ih[8] = 8; ih[9] = 6;
    return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk('IHDR', ih), chunk('IDAT', zlib.deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
  }
  function overlay(items, cols, rows, cw, rh, behind) {
    const K = 0.2; // пикселей на twip (≈3x от 96 dpi)
    const W = Math.ceil(cols * cw * K), H = Math.ceil(rows * rh * K);
    const px = Buffer.alloc(W * H * 4);
    const cx = v => v * cw * K, cy = v => v * rh * K;
    const SS = 3; // сглаживание: 3×3 подвыборки на пиксель
    // заливка многоугольника (even-odd) с подвыборкой, цвет [r,g,b]
    function fill(poly, col) {
      const ys = poly.map(p => p[1]), xs = poly.map(p => p[0]);
      const y0 = Math.max(0, Math.floor(Math.min(...ys))), y1 = Math.min(H - 1, Math.ceil(Math.max(...ys)));
      const x0 = Math.max(0, Math.floor(Math.min(...xs))), x1 = Math.min(W - 1, Math.ceil(Math.max(...xs)));
      if (x1 < x0 || y1 < y0) return;
      const cov = new Float32Array((x1 - x0 + 1) * (y1 - y0 + 1));
      for (let y = y0; y <= y1; y++) for (let sy = 0; sy < SS; sy++) {
        const Y = y + (sy + 0.5) / SS, xi = [];
        for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
          const [ax, ay] = poly[i], [bx, by] = poly[j];
          if ((ay > Y) !== (by > Y)) xi.push(ax + (Y - ay) * (bx - ax) / (by - ay));
        }
        xi.sort((a, b) => a - b);
        for (let k = 0; k + 1 < xi.length; k += 2)
          for (let x = Math.max(x0, Math.floor(xi[k])); x <= Math.min(x1, Math.floor(xi[k + 1])); x++)
            for (let sx = 0; sx < SS; sx++) { const X = x + (sx + 0.5) / SS; if (X >= xi[k] && X < xi[k + 1]) cov[(y - y0) * (x1 - x0 + 1) + x - x0] += 1 / (SS * SS); }
      }
      for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) {
        const a = cov[(y - y0) * (x1 - x0 + 1) + x - x0]; if (a <= 0) continue;
        const o = (y * W + x) * 4, A = px[o + 3] / 255, na = a + A * (1 - a);
        for (let c = 0; c < 3; c++) px[o + c] = Math.round((col[c] * a + px[o + c] * A * (1 - a)) / na);
        px[o + 3] = Math.round(na * 255);
      }
    }
    const BLK = [0, 0, 0], WHT = [255, 255, 255];
    const line = (a, b, w, col = BLK) => { const dx = b[0] - a[0], dy = b[1] - a[1], L = Math.hypot(dx, dy) || 1, nx = -dy / L * w / 2, ny = dx / L * w / 2;
      const ex = dx / L * w / 2, ey = dy / L * w / 2; // концы с выносом на полтолщины — без «щелей» в углах
      fill([[a[0] - ex + nx, a[1] - ey + ny], [b[0] + ex + nx, b[1] + ey + ny], [b[0] + ex - nx, b[1] + ey - ny], [a[0] - ex - nx, a[1] - ey - ny]], col); };
    const rect = (p, u, n, len, th) => [[p[0] - n[0] * th, p[1] - n[1] * th], [p[0] + u[0] * len - n[0] * th, p[1] + u[1] * len - n[1] * th],
      [p[0] + u[0] * len + n[0] * th, p[1] + u[1] * len + n[1] * th], [p[0] + n[0] * th, p[1] + n[1] * th]];
    const unit = Math.min(cw, rh) * K;          // размер клетки в пикселях
    const LW = Math.max(3, 25 * K);             // толщина линии ≈ рамка стены (1,25 pt)
    const TH = Math.max(LW * 1.8, unit * 0.2);    // полутолщина значка окна поперёк стены
    [...items].sort((a, b) => ({ hatch: 0, clear: 1 }[a.kind] ?? 2) - ({ hatch: 0, clear: 1 }[b.kind] ?? 2)).forEach(it => { // штриховка снизу
      if (it.kind === "window" || it.kind === "door") {
        const a = [cx(it.p1[0]), cy(it.p1[1])], b = [cx(it.p2[0]), cy(it.p2[1])];
        const len = Math.hypot(b[0] - a[0], b[1] - a[1]); if (len < 1) return;
        const u = [(b[0] - a[0]) / len, (b[1] - a[1]) / len], n = it.n; // n — нормаль внутрь объекта
        if (it.kind === "window") {
          fill(rect(a, u, n, len, TH), WHT);
          const o = rect(a, u, n, len, TH); o.forEach((p, i) => line(p, o[(i + 1) % 4], LW * 0.8));
          line(a, b, LW * 0.6);
        } else {
          fill(rect(a, u, n, len, LW * 1.6), WHT);               // разрыв в стене
          line([a[0] - n[0] * TH * 0.7, a[1] - n[1] * TH * 0.7], [a[0] + n[0] * TH * 0.7, a[1] + n[1] * TH * 0.7], LW); // откосы
          line([b[0] - n[0] * TH * 0.7, b[1] - n[1] * TH * 0.7], [b[0] + n[0] * TH * 0.7, b[1] + n[1] * TH * 0.7], LW);
          const hinge = it.hinge === "end" ? b : a, d = it.hinge === "end" ? [-u[0], -u[1]] : u;
          const s = it.open === "out" ? -1 : 1, ang = 60 * Math.PI / 180;
          const leaf = [hinge[0] + len * (Math.cos(ang) * d[0] + Math.sin(ang) * n[0] * s), hinge[1] + len * (Math.cos(ang) * d[1] + Math.sin(ang) * n[1] * s)];
          line(hinge, leaf, LW * 0.8);                             // полотно открыто под углом
        }
      } else if (it.kind === "hatch") {
        const x0 = cx(it.x), y0 = cy(it.y), x1 = cx(it.x + it.w), y1 = cy(it.y + it.h), sp = unit * (it.full ? 0.45 : 0.32);
        for (let c = x0 - (y1 - y0); c < x1; c += sp * Math.SQRT2) { // линии 45°: x = c + (y1 - y)
          const ya = Math.max(y0, y1 - (x1 - c)), yb = Math.min(y1, y1 - (x0 - c));
          if (yb - ya > 0.5) line([c + (y1 - ya), ya], [c + (y1 - yb), yb], LW * 0.45);
        }
        if (!it.full) { const r = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]; r.forEach((p, i) => line(p, r[(i + 1) % 4], LW * 0.5)); }
      } else if (it.kind === "clear") {
        fill([[cx(it.x), cy(it.y)], [cx(it.x + it.w), cy(it.y)], [cx(it.x + it.w), cy(it.y + it.h)], [cx(it.x), cy(it.y + it.h)]], WHT);
      } else if (it.kind === "mark") {
        const mx = cx(it.x), my = cy(it.y), r = unit * 0.42, N = 48;
        const pt = k => [mx + r * Math.cos(k), my + r * Math.sin(k)];
        fill(Array.from({ length: N }, (_, i) => pt(2 * Math.PI * i / N)), WHT);
        for (let i = 0; i < N; i++) line(pt(2 * Math.PI * i / N), pt(2 * Math.PI * (i + 1) / N), LW * 0.8);
        if (it.type === "fire") { const d = r * 0.7; line([mx - d, my - d], [mx + d, my + d], LW * 0.8); line([mx - d, my + d], [mx + d, my - d], LW * 0.8); }
        else fill(Array.from({ length: N }, (_, i) => [mx + r * 0.42 * Math.cos(2 * Math.PI * i / N), my + r * 0.42 * Math.sin(2 * Math.PI * i / N)]), BLK);
      } else if (it.kind === "dim") {
        const a = [cx(it.p1[0]), cy(it.p1[1])], b = [cx(it.p2[0]), cy(it.p2[1])], L = Math.hypot(b[0] - a[0], b[1] - a[1]); if (L < 2) return;
        const u = [(b[0] - a[0]) / L, (b[1] - a[1]) / L], n = [-u[1], u[0]], tk = unit * 0.25, ar = unit * 0.3;
        line(a, b, LW * 0.5);
        [[a, 1], [b, -1]].forEach(([p, s]) => { // стрелки на концах + засечки
          line(p, [p[0] + s * u[0] * ar + n[0] * ar * 0.4, p[1] + s * u[1] * ar + n[1] * ar * 0.4], LW * 0.5);
          line(p, [p[0] + s * u[0] * ar - n[0] * ar * 0.4, p[1] + s * u[1] * ar - n[1] * ar * 0.4], LW * 0.5);
          line([p[0] - n[0] * tk, p[1] - n[1] * tk], [p[0] + n[0] * tk, p[1] + n[1] * tk], LW * 0.5); });
      } else if (it.kind === "north") { // стрелка на север и буква «С» (дугой, без шрифтов)
        const x = cx(it.x), y0 = cy(it.y), u = unit, top = [x, y0 + u * 1.15], bot = [x, y0 + u * 2.6];
        line(top, bot, LW * 0.7);
        fill([top, [x - u * 0.28, top[1] + u * 0.55], [x + u * 0.28, top[1] + u * 0.55]], BLK);
        const r = u * 0.42, c = [x, y0 + u * 0.5], N = 24;
        for (let i = 0; i < N; i++) { const k0 = Math.PI * (0.25 + 1.5 * i / N), k1 = Math.PI * (0.25 + 1.5 * (i + 1) / N);
          line([c[0] + r * Math.cos(k0), c[1] - r * Math.sin(k0)], [c[0] + r * Math.cos(k1), c[1] - r * Math.sin(k1)], LW * 0.8); }
      } else if (it.kind === "stove") {
        const x0 = cx(it.x), y0 = cy(it.y), x1 = cx(it.x + it.w), y1 = cy(it.y + it.h);
        if (it.round) {
          const r = Math.min(x1 - x0, y1 - y0) / 2, mx = (x0 + x1) / 2, my = (y0 + y1) / 2, N = 64;
          const pt = k => [mx + r * Math.cos(k), my + r * Math.sin(k)];
          fill(Array.from({ length: N / 2 + 1 }, (_, i) => pt(-Math.PI / 4 + Math.PI * i / (N / 2))), BLK); // закрашенная половина
          for (let i = 0; i < N; i++) line(pt(2 * Math.PI * i / N), pt(2 * Math.PI * (i + 1) / N), LW * 0.8);
          line(pt(-Math.PI / 4), pt(3 * Math.PI / 4), LW * 0.8);
        } else {
          fill([[x0, y0], [x1, y0], [x1, y1]], BLK); // половина закрашена
          [[x0, y0, x1, y0], [x1, y0, x1, y1], [x1, y1, x0, y1], [x0, y1, x0, y0], [x0, y0, x1, y1]].forEach(([a, b, c, d]) => line([a, b], [c, d], LW * 0.8));
        }
      }
    });
    const png = pngRGBA(W, H, px);
    const wPt = cols * cw / 15, hPt = rows * rh / 15; // twip → пиксели 96 dpi
    return (offX) => new ImageRun({ type: "png", data: png, transformation: { width: wPt, height: hPt },
      floating: {
        horizontalPosition: { relative: HorizontalPositionRelativeFrom.CHARACTER, offset: Math.round(offX * 635) },
        verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
        wrap: { type: TextWrappingType.NONE }, allowOverlap: true, layoutInCell: true, zIndex: behind ? 1 : 5, behindDocument: !!behind,
      } });
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
    // значки (окна, двери, печи) — прозрачная картинка поверх сетки, якорь в первой пустой клетке верхнего ряда
    // штриховка и подложки — ПОД текстом (за таблицей), остальные значки — поверх рамок
    const low = it => it.kind === "hatch" || it.kind === "clear";
    const dr = shapes.draw || [], back = dr.filter(low), front = dr.filter(it => !low(it));
    const ovB = back.length ? overlay(back, cols, rows, cw, rh, true) : null, ovF = front.length ? overlay(front, cols, rows, cw, rh) : null;
    const ov = ovB || ovF ? (off) => [ovB && ovB(off), ovF && ovF(off)].filter(Boolean) : null;
    const anchorC = ov ? own[0].findIndex(v => v < 0) : -1;
    const trs = [];
    for (let r = 0; r < rows; r++) {
      const cells = [];
      for (let c = 0; c < cols;) {
        const si = own[r][c];
        if (si < 0) { cells.push(new TableCell({ children: [new Paragraph({ children: r === 0 && c === anchorC ? ov(-c * cw) : [] })], width: { size: cw, type: WidthType.DXA },
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
    const padL = opt.dimSide ? 3 : 1, padT = 2, padR = opt.dimSide ? 4 : 1, padB = 1; // слева запас под подпись высоты, если справа занято
    const sc = opt.scale || Math.min((cols - padL - padR) / (maxX - minX), (rows - padT - padB) / (maxY - minY));
    const shapes = [], labels = [];
    const X = v => padL + (v - minX) * sc, Y = v => padT + (v - minY) * sc;
    const draw = [], legend = {}, rects = {};
    objs.forEach(o => {
      const x = Math.round(X(o.x || 0)), y = Math.round(Y(o.y || 0));
      const w = Math.max(1, Math.round(o.w * sc)), h = Math.max(1, Math.round(o.h * sc));
      // проёмы: { type: "window"|"door", wall: "top"|"bottom"|"left"|"right", from: м от левого/верхнего угла, w: ширина, м,
      //           open: "in"|"out" (дверь, по умолч. внутрь), hinge: "start"|"end" (петли у начала или конца проёма) }
      const kx = w / o.w, ky = h / o.h;
      (o.openings || []).forEach(op => {
        const f = +String(op.from ?? 0).replace(",", "."), ow = +String(op.w ?? 0.9).replace(",", ".");
        const wall = op.wall || "bottom", horiz = wall === "top" || wall === "bottom";
        const at = horiz ? (wall === "top" ? y : y + h) : (wall === "left" ? x : x + w);
        const s0 = horiz ? x + f * kx : y + f * ky, s1 = horiz ? x + Math.min(o.w, f + ow) * kx : y + Math.min(o.h, f + ow) * ky;
        draw.push({ kind: op.type === "door" ? "door" : "window", open: op.open, hinge: op.hinge,
          p1: horiz ? [s0, at] : [at, s0], p2: horiz ? [s1, at] : [at, s1],
          n: { top: [0, 1], bottom: [0, -1], left: [1, 0], right: [-1, 0] }[wall] });
      });
      // печи: { x, y, w, h } в м от левого верхнего угла объекта; round: true — круглая (диаметр = w)
      (o.stoves || []).forEach(st => {
        const sw = +String(st.w ?? 1).replace(",", "."), sh = st.round ? sw : +String(st.h ?? sw).replace(",", ".");
        draw.push({ kind: "stove", round: !!st.round, x: x + (+String(st.x ?? 0).replace(",", ".")) * kx, y: y + (+String(st.y ?? 0).replace(",", ".")) * ky,
          w: st.round ? Math.min(sw * kx, sw * ky) : sw * kx, h: st.round ? Math.min(sw * kx, sw * ky) : sh * ky });
      });
      // зоны повреждения штриховкой: hatch: true — весь объект; zones: [{ x, y, w, h }] — м от левого верхнего угла объекта
      const num = v => +String(v ?? 0).replace(",", ".");
      if (o.hatch) { draw.push({ kind: "hatch", x, y, w, h, full: true }); legend.hatch = 1; }
      (o.zones || []).forEach(z => { draw.push({ kind: "hatch", x: x + num(z.x) * kx, y: y + num(z.y) * ky, w: num(z.w) * kx, h: num(z.h) * ky }); legend.hatch = 1; });
      // отметки: marks: [{ x, y, type: "leak"|"fire", text }] — место протечки / очаг пожара, м от левого верхнего угла объекта
      (o.marks || []).forEach(m => { const ty = m.type === "fire" ? "fire" : "leak";
        draw.push({ kind: "mark", type: ty, x: x + num(m.x) * kx, y: y + num(m.y) * ky });
        (legend[ty] = legend[ty] || []).push(m.text || ""); });
      rects[o.name] = { x, y, w, h, o };
      // подложка под название объекта, чтобы штриховка его не перечёркивала (размер подписи — прикидкой по символам)
      const hitC = z => num(z.x) <= o.w / 2 + 0.6 && num(z.x) + num(z.w) >= o.w / 2 - 0.6 && num(z.y) <= o.h / 2 + 0.4 && num(z.y) + num(z.h) >= o.h / 2 - 0.4;
      if ((o.hatch || (o.zones || []).some(hitC)) && o.name) { const tw = Math.min(w - 0.2, String(o.name).length * 0.36 + 0.4);
        draw.push({ kind: "clear", x: x + (w - tw) / 2, y: y + h / 2 - 0.45, w: tw, h: 0.9 }); }
      const label = o.name + (opt.dimsInside !== false && !opt.dimSide ? `\n${fmt(o.w)}×${fmt(o.h)} м` : "");
      shapes.push({ x, y, w, h, text: label, sz: opt.sz || 14 });
      if (opt.dimSide) {
        // строка для подписи высоты: ближе к середине, но мимо проёмов на боковых стенах
        const sideSpans = (o.openings || []).filter(op => op.wall === "left" || op.wall === "right").map(op => {
          const f = +String(op.from ?? 0).replace(",", ".") * ky, ow = +String(op.w ?? 0.9).replace(",", ".") * ky;
          return [y + f - (op.type === "door" ? ow * 0.6 : 0.3), y + f + ow + (op.type === "door" ? ow * 0.6 : 0.3)]; });
        const mid = y + Math.floor((h - 1) / 2), rowsBy = [mid];
        for (let d = 1; d < h; d++) rowsBy.push(mid - d, mid + d);
        const my = rowsBy.find(r => r >= y && r < y + h && !sideSpans.some(([a, b]) => r + 1 > a && r < b)) ?? mid;
        // подпись длины не должна налезать на значки на стене — сдвиг на клетку дальше
        const busy = side => (o.openings || []).some(op => (op.wall || "bottom") === side);
        const ty = busy("top") ? 2 : 1, by = busy("bottom") ? 1 : 0;
        labels.push({ x, y: y - ty, w, h: 1, text: `${fmt(o.w)} м`, border: false, sz: 14, alt: [{ x, y: y + h + by, w, h: 1 }, { x, y: y - 1, w, h: 1 }] });
        const rx = busy("right") ? 1 : 0, lx = busy("left") ? 1 : 0;
        labels.push({ x: x + w + rx, y: my, w: 3, h: 1, text: `${fmt(o.h)} м`, border: false, sz: 14,
          alt: [{ x: x - 3 - lx, y: my, w: 3, h: 1 }, { x: x + w + rx, y: my, w: 2, h: 1 }, { x: x - 2 - lx, y: my, w: 2, h: 1 }, { x: x + w, y: my, w: 3, h: 1 }] });
      }
    });
    // расстояния между строениями: gaps: [{ a: "Жилой дом", b: "Баня", m: "6" }] (m не задано — считается по координатам)
    (opt.gaps || []).forEach(g => {
      const A = rects[g.a], B = rects[g.b]; if (!A || !B) return;
      const oa = A.o, ob = B.o, ax0 = oa.x || 0, ay0 = oa.y || 0, bx0 = ob.x || 0, by0 = ob.y || 0;
      const ovX = Math.min(ax0 + oa.w, bx0 + ob.w) - Math.max(ax0, bx0), ovY = Math.min(ay0 + oa.h, by0 + ob.h) - Math.max(ay0, by0);
      let p1, p2, dm;
      if (ovX > 0 && !(ovY > 0)) { // друг над другом
        const mx = X(Math.max(ax0, bx0) + ovX / 2), up = ay0 < by0 ? A : B, dn = up === A ? B : A;
        p1 = [mx, up.y + up.h]; p2 = [mx, dn.y]; dm = Math.abs((ay0 < by0 ? by0 - ay0 - oa.h : ay0 - by0 - ob.h));
      } else if (ovY > 0 && !(ovX > 0)) { // рядом по горизонтали
        const my = Y(Math.max(ay0, by0) + ovY / 2), lf = ax0 < bx0 ? A : B, rt = lf === A ? B : A;
        p1 = [lf.x + lf.w, my]; p2 = [rt.x, my]; dm = Math.abs(ax0 < bx0 ? bx0 - ax0 - oa.w : ax0 - bx0 - ob.w);
      } else { // наискосок — между ближайшими углами
        const cx = (r, o, other) => (o.x || 0) + o.w <= (other.x || 0) ? r.x + r.w : r.x;
        const cy = (r, o, other) => (o.y || 0) + o.h <= (other.y || 0) ? r.y + r.h : r.y;
        p1 = [cx(A, oa, ob), cy(A, oa, ob)]; p2 = [cx(B, ob, oa), cy(B, ob, oa)];
        const dx = Math.max(0, bx0 - ax0 - oa.w, ax0 - bx0 - ob.w), dy = Math.max(0, by0 - ay0 - oa.h, ay0 - by0 - ob.h); dm = Math.hypot(dx, dy);
      }
      draw.push({ kind: "dim", p1, p2 });
      const txt = `${g.m !== undefined ? String(g.m).replace(".", ",") : fmt(dm)} м`;
      const mx = (p1[0] + p2[0]) / 2, my = (p1[1] + p2[1]) / 2, vert = Math.abs(p2[1] - p1[1]) > Math.abs(p2[0] - p1[0]);
      labels.push(vert ? { x: Math.round(mx) + 1, y: Math.floor(my), w: 3, h: 1, text: txt, border: false, sz: 14, alt: [{ x: Math.round(mx) - 4, y: Math.floor(my), w: 3, h: 1 }] }
                       : { x: Math.round(mx) - 1, y: Math.floor(my) - 1, w: 3, h: 1, text: txt, border: false, sz: 14, alt: [{ x: Math.round(mx) - 1, y: Math.floor(my) + 1, w: 3, h: 1 }] });
    });
    if (opt.north) draw.push({ kind: "north", x: cols - 1.5, y: 0.15 });
    shapes.push(...labels);
    (opt.extra || []).forEach(e => shapes.push({ x: Math.round(X(e.x)), y: Math.round(Y(e.y)), w: Math.max(1, Math.round((e.w || 4) * sc)), h: 1, text: e.text, border: false, sz: 14 }));
    shapes.draw = draw;
    // условные обозначения под схемой — только то, что реально нарисовано
    const lg = [];
    if (legend.hatch) lg.push("штриховка — зона повреждения");
    const lt = arr => arr.filter(Boolean).length ? ` (${arr.filter(Boolean).join("; ")})` : "";
    if (legend.leak) lg.push("⊙ — место протечки" + lt(legend.leak));
    if (legend.fire) lg.push("⊗ — очаг пожара" + lt(legend.fire));
    shapes.legend = lg.length ? "Обозначения: " + lg.join("; ") : "";
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
      P([t("ООО «ФАВОРИТ»", { b: true, sz: 40 }), ...(D.actNo ? [t("     " + D.actNo, { b: true, sz: 30, u: true })] : [])], { a: C, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "7030A0", space: 4 } }, after: 160 }),
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
    if (!el || ((!el.dmg || !el.dmg.length) && !el.pct)) lines.push([el && el.note ? el.note : ND, "—"]);
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
  const bldSecs = [];
  (D.buildings || []).forEach(b => {
    const W = LAND_W;
    const cols = [500, 3000, 1900, W - 500 - 3000 - 1900 - 3400, 3400];
    const items = BLD_EL.map(([k, name, h, minN], i) => {
      const el = (b.el || {})[k];
      return {
        lead: [[P(t(`${i + 1}.`, { b: true, sz: 20 }))], elName(name, h), [P(t(nz(el && el.mat), { sz: 18, b: true }), { a: C })]],
        lines: b.total
          ? (() => { const l = [[b.total === true ? "сгорело полностью, без годных остатков" : b.total, "100%"]]; while (l.length < minN) l.push(["", ""]); return l; })()
          : dmgLines(el, minN, `Общий процент повреждения элемента «${name.toLowerCase()}» в целом`),
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
    bldSecs.push({ properties: landProps, children: [
      P(t("СТРОЕНИЕ", { b: true, sz: 24 }), { a: R }), head, P(t(""), { after: 80 }),
      table(cols, [hdr, ...er, p8Row(b.p8 || (b.total ? { otdelka: "100%", inzh: "100%" } : undefined), W, 8, 5)]),
      P(t(""), { after: 60 }), signBlock(W),
    ] });
  });

  // =================== ОБЩАЯ СХЕМА ===================
  if (D.scheme) {
    const W = PORT_W, cols = 40, cw = Math.floor(W / cols), rows = 54, rh = cw;
    const shapes = D.scheme.shapes || placeObjects(D.scheme.objects, cols, rows, { dimSide: true, extra: D.scheme.labels, sz: 15, gaps: D.scheme.gaps, north: D.scheme.north });
    sections.push({ properties: portProps, children: [
      P(t("ОБЩАЯ СХЕМА МЕСТА ОСМОТРА", { b: true, sz: 24 }), { a: R, after: 120 }),
      grid(cols, rows, cw, rh, shapes),
      ...(shapes.legend ? [P(t(shapes.legend, { b: true, sz: 18 }), { before: 60 })] : []),
      ...(D.scheme.note ? [P(t(D.scheme.note, { b: true, sz: 20 }), { before: 80 })] : []),
      P(t(""), { after: 120 }), signBlock(W),
    ] });
  }
  sections.push(...bldSecs);

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
        r(it.shown === false ? "нет" : "да", { w: cols[6] }), r(it.shown === false ? nz(it.reason, D.cause === "pozhar" ? "1" : "4. - иное") : "—", { w: cols[7], a: it.shown === false ? L : C }),
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
        return m[0] && m[1] ? [{ name: rm.name, x: 0, y: 0, w: m[0], h: m[1], openings: rm.plan && rm.plan.openings, stoves: rm.plan && rm.plan.stoves,
          zones: rm.plan && rm.plan.zones, marks: rm.plan && rm.plan.marks, hatch: rm.plan && rm.plan.hatch }] : [];
      })();
      pshapes = placeObjects(objs, gc, grows, { dimSide: true, sz: 14 });
    }
    const planT = [P(t("План-схема помещения", { b: true, sz: 20 }), { a: R, after: 40 }), grid(gc, grows, gcw, grh, pshapes),
      ...(pshapes.legend ? [P(t(pshapes.legend, { sz: 13 }), { line: 200 })] : []),
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
      top, P(t(""), { after: pshapes.legend ? 0 : 80, line: pshapes.legend ? 120 : 240 }), table(cols, [hdr, ...er]), P(t(""), { after: 40 }), signBlock(W),
    ] });
  });

  const doc = new Document({ styles: { default: { document: { run: { font: F } } } }, sections });
  return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT_FILE, buf); console.log("OK", OUT_FILE); });
};
