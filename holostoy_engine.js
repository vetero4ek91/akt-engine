const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, VerticalAlign, UnderlineType,
  ImageRun, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType
} = require('docx');
const fs = require('fs');

// Движок «Протокол осмотра территории» (холостой выезд, ООО «Фаворит»). Пустая форма:
// данные, подпись и ФИО исполнителя передаются снаружи.
module.exports = function build(cfg) {
  const { D, SIGN_EXPERT, EXPERT, SIG_RLE, OUT_FILE } = cfg;
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

const f = "Times New Roman", SZ = 24;
const t = (text, o = {}) => new TextRun({ text, size: o.sz || SZ, bold: !!o.b, italics: !!o.i, strike: !!o.s, font: f,
  underline: o.u ? { type: UnderlineType.SINGLE } : undefined });
const v = (x, blank = "__________") => x ? t(x, { b: true, u: true }) : t(blank);   // вписанное — жирным с подчёркиванием
const P = (children, o = {}) => new Paragraph({ alignment: o.a || AlignmentType.JUSTIFIED,
  indent: o.noInd ? undefined : { firstLine: 567 }, spacing: { before: 0, after: o.after || 0, line: 300 }, children });
const C = AlignmentType.CENTER;
const sp = (n = 1) => Array.from({ length: n }, () => new Paragraph({ children: [t("")] }));
// «наличие/отсутствие»: ненужное зачёркнуто
const presence = (present) => [t("наличие", { s: !present, b: present }), t("/"), t("отсутствие", { s: present, b: !present })];
const INS = { "Росгосстрах": "ПАО СК «Росгосстрах»", "СОГАЗ": "АО «СОГАЗ»" }[D.insurer || "Росгосстрах"] || D.insurer;
const hm = (x) => x && x[0] ? [v(x[0]), t(" час. "), v(x[1]), t(" мин.")] : [t("____ час. ______ мин.")];
const w = D.witnesses || [];
const witness = (n) => { const x = w[n - 1] || {}; return P([t(`${n})   ФИО `), v(x.fio, "______________________________________"), t(", (тел. "), v(x.tel, "_______________"), t(");")]); };

const body = [
  new Paragraph({ alignment: AlignmentType.RIGHT, children: [t("Убыток № ", { b: true }), t(D.claimNo || "____________", { b: true })] }),
  new Paragraph({ alignment: C, children: [t("ООО «Фаворит»", { b: true, sz: 26 })] }),
  ...sp(1),
  new Paragraph({ alignment: C, children: [t("ПРОТОКОЛ", { b: true })] }),
  new Paragraph({ alignment: C, children: [t("осмотра территории", { b: true })] }),
  ...sp(1),
  new Table({ width: { size: 9638, type: WidthType.DXA }, columnWidths: [4819, 4819], borders: { top: {style:BorderStyle.NONE}, bottom: {style:BorderStyle.NONE}, left: {style:BorderStyle.NONE}, right: {style:BorderStyle.NONE}, insideHorizontal: {style:BorderStyle.NONE}, insideVertical: {style:BorderStyle.NONE} },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: 4819, type: WidthType.DXA }, children: [new Paragraph({ children: [t("г. "), v(D.city, "__________________")] })] }),
      new TableCell({ width: { size: 4819, type: WidthType.DXA }, children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: D.day ? [v(`«${D.day}» ${D.month} ${D.year4} г.`)] : [t("«____»_________________ 20__ г.")] })] }),
    ] })] }),
  ...sp(1),
  P([t("Осмотр начат в "), ...hm(D.start)]),
  P([t("Осмотр окончен в "), ...hm(D.end)], { after: 160 }),
  P([t("Исполнитель "), v(EXPERT), t(" (ФИО) по поручению " + INS + " прибыл для проведения осмотра транспортного средства по адресу "),
     v(D.address, "____________________________________________"), t(" и ко времени, согласованному с потерпевшим.")], { after: 160 }),
  P([t("В присутствии "), t("(данные представителя заказчика или иного свидетеля)", { i: true }), t(":")]),
  witness(1), witness(2),
  P([t("произвел осмотр территории по вышеуказанному адресу и установил "), ...presence(!!D.carPresent),
     t(" на указанной территории транспортного средства за государственным номером "), v(D.plate, "________________"),
     t(" по состоянию на момент проведения осмотра.")]),
  P([t("Дополнительно установил "), ...presence(!!D.ownerPresent),
     t(" на указанной территории собственника транспортного средства за государственным номером "), v(D.plate, "________________"),
     t(" или его представителя на момент проведения осмотра.")], { after: 160 }),
  P([t("Свидетелям/представителям заказчика также объявлено о применении технических средств при осмотре территории: цифрового фотоаппарата.")], { after: 160 }),
  P([t("К протоколу осмотра территории прилагается фототаблица (приложение N 1).")]),
  P([t("Протокол прочитан Исполнителем вслух. Записано правильно. Замечания к проведенной процедуре осмотра и к протоколу отсутствуют.")]),
  ...sp(2),
  sigLines([{ label: "Свидетель/представитель заказчика:", fio: (w[0] || {}).fio || "" }], SZ - 2),
  ...sp(1),
  sigLines([{ label: "Исполнитель:", fio: EXPERT, expert: true }], SZ - 2),
];

const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 850, bottom: 1000, left: 1418 } } }, children: body }] });
return Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT_FILE, buf); console.log("OK"); });
};
