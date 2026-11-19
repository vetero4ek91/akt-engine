// ВХОДНЫЕ ДАННЫЕ РЕАЛЬНОГО ОСМОТРА. ЗАПОЛНЯТЬ ТОЛЬКО ПО ДОКУМЕНТАМ И ФОТО.
// ДАННЫЕ НИЖЕ ПУСТЫЕ НАМЕРЕННО: генератор не содержит вымышленного ТС.
const D = {
 claimNo:'', insurer:'Росгосстрах', polis:'', culpritSK:'', culpritPolis:'',
 inspType:'первичный', mode:'выездной', address:'', date:'', start:'', end:'',
 plate:'', vin:'', model:'', bodyType:'', doors:'', startDate:'', year:'',
 mileage:'', chassis:'', color:'', paint:'', fuel:'', engineModel:'',
 power:'', powerKw:'', volume:'', stsSeries:'', stsNumber:'', owner:'',
 phone:'', proxy:'', eventDate:'', eventPlace:'', state:'', conclusion:'',
 documentsStatus:'неизвестно', identificationDiscrepancies:[], extraEquipment:[],
 defects:[], uncertainDamages:[],
 // Сторона основного удара: 'left' / 'right' / 'both' / 'undetermined'.
 // Два независимых фото/ориентира, подтверждающих сторону по ходу движения.
 impactSide:'undetermined', orientationEvidence:[]
};
// Согласованный список повреждений: [{detail:'',description:''}, ...]
const damages = [];
// КАСКО: по одному объекту на повреждение, ключи zamena/remont/okraska/diagn.
// При remont:true в графе «Ремонт, час» печатается X.
const KASKO_REPAIRS = [];
// present = подтверждённо есть; absent = подтверждённо нет; unknown = не установлено.
const EQUIPMENT = {};
// Поддержка совместимости со старой шапкой.
const KOMPL = {};
const CHECKS = {
 insurance:'КАСКО', inspection:'первичный', place:'выездной', cleanliness:'', stoa:false,
 vinOk:true, extraEquip:false, defects:null, uncertain:null, hidden:null,
 dismantling:false, repairable:null, match:null
};
const SIGN_EXPERT = true;
const OWNER_MODE = 'none';
// Конец блока данных.

// Защита от случайной сборки акта с пустыми/тестовыми данными.
// Для тестов укажи AKT_TEST=1. Реальный акт требует реквизитов и списка повреждений.
if (process.env.AKT_TEST !== '1') {
  const required = ['claimNo','insurer','polis','plate','vin','model','stsSeries','stsNumber','owner','date'];
  const missing=required.filter(k=>!String(D[k]||'').trim());
  if(missing.length)throw new Error('Не заполнены обязательные поля: '+missing.join(', '));
  if(!Array.isArray(damages)||!damages.length)throw new Error('Нет согласованного списка повреждений');
  if(String(D.vin).length!==17)throw new Error('VIN: требуется 17 символов');
  // Блокирующий контроль стороны: никакой автоматической догадки по ракурсу фотографа.
  // Проверяется до создания DOCX, включая описания, состояние и заключение.
  const side=String(D.impactSide||'undetermined').toLowerCase();
  if(!['left','right','both'].includes(side))
    throw new Error('Сторона удара не подтверждена: укажи impactSide по ходу движения автомобиля');
  if(!Array.isArray(D.orientationEvidence)||D.orientationEvidence.filter(x=>String(x||'').trim()).length<2)
    throw new Error('Сторона удара: нужны минимум два независимых фото/ориентира в orientationEvidence');
  if(new Set(D.orientationEvidence.map(x=>String(x||'').trim()).filter(Boolean)).size<2)
    throw new Error('Сторона удара: ориентиры должны быть независимыми');
  const sideTexts=[...damages.map((d,i)=>({where:'повреждение №'+(i+1),value:String(d.detail||'')+' '+String(d.description||'')})),
    {where:'состояние ТС',value:String(D.state||'')},{where:'заключение',value:String(D.conclusion||'')}];
  // В JS кириллица не всегда считается word-character, поэтому используем границы по не-буквам.
  const hasLeft=t=>/(^|[^А-Яа-яЁё])(?:лев(?:ый|ая|ое|ого|ому|ым|ом|ую|ой|ые|ых|ыми)|слева)(?=$|[^А-Яа-яЁё])/iu.test(t);
  const hasRight=t=>/(^|[^А-Яа-яЁё])(?:прав(?:ый|ая|ое|ого|ому|ым|ом|ую|ой|ые|ых|ими)|справа)(?=$|[^А-Яа-яЁё])/iu.test(t);
  for(const item of sideTexts){
    if(side==='left'&&hasRight(item.value))throw new Error('Противоречие стороне удара (левая): '+item.where);
    if(side==='right'&&hasLeft(item.value))throw new Error('Противоречие стороне удара (правая): '+item.where);
  }
  if(side==='both'&&!sideTexts.some(x=>hasLeft(x.value))||side==='both'&&!sideTexts.some(x=>hasRight(x.value)))
    throw new Error('Для удара с двух сторон в перечне должны быть подтверждены обе стороны');

  if(String(CHECKS.insurance||'').toUpperCase()==='КАСКО' && KASKO_REPAIRS.length!==damages.length)
    throw new Error('КАСКО: число ремонтных воздействий должно совпадать с повреждениями');
}

const fs = require('fs');
const os = require('os');
const path = require('path');
const cp = require('child_process');
const { EXPERT, SIG_RLE } = require('./podpis.js');
const BASE = path.join(__dirname, 'АО_18.11.2026_fin_с восст. рем (1).docx');
const SIG = path.join(__dirname, 'signature.png');
const DRAWING = fs.readFileSync(path.join(__dirname, 'drawing.txt'), 'utf8').replace('layoutInCell="0"','layoutInCell="1"').replace('<wp:posOffset>-300000</wp:posOffset>','<wp:posOffset>150000</wp:posOffset>');
const OUT_FILE = process.env.AKT_OUT || path.join(__dirname, 'Akt_osmotra.docx');
if (!fs.existsSync(BASE)) throw new Error('Отсутствует оригинальный DOCX бланк');
if (!SIG_RLE || EXPERT !== 'Андреев А.Г.') throw new Error('Не подключен podpis.js');
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'novyi-akt-'));
try {
  cp.execFileSync('unzip', ['-q', BASE, '-d', tmp]);
  const xmlPath = path.join(tmp, 'word/document.xml');
  let xml = fs.readFileSync(xmlPath, 'utf8');
  const escape = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const run = s => `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="1"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr><w:t xml:space="preserve">${escape(s)}</w:t></w:r>`;
  // Работать с ФИЗИЧЕСКИМИ ячейками DOCX, не с ячейками заголовков.
  // Внесённый текст идёт строго в отдельную графу, не поверх метки поля.
  const tblRE=/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g;
  const rowRE=/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g;
  const cellRE=/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g;
  function setCell(cell, value, opts={}) {
    if (value===undefined || value===null || String(value)==='') return cell;
    const pos=cell.lastIndexOf('</w:p>');
    if(pos<0)return cell.replace('</w:tc>','<w:p>'+run(String(value))+'</w:p></w:tc>');
    const v=opts.underline?`<w:r><w:rPr><w:u w:val="single"/><w:b w:val="0"/><w:sz w:val="16"/></w:rPr><w:t>${escape(value)}</w:t></w:r>`:run(String(value));
    return cell.slice(0,pos)+v+cell.slice(pos);
  }
  function replaceCellRow(row, idx, val, opts={}){
    const cells=[...row.matchAll(cellRE)].map(m=>m[0]);
    if(idx>=cells.length)throw new Error('Нет графы '+idx+' в строке ('+cells.length+' ячеек)');
    return row.replace(cells[idx],setCell(cells[idx],val,opts));
  }
  function editTable(index, callback){
    let ti=-1,done=false;
    xml=xml.replace(tblRE,t=>{ti++;if(ti!==index)return t;done=true;return callback(t);});
    if(!done)throw new Error('Не найдена таблица '+index);
  }
  function setTableRow(t,ri, edits){
    let i=-1,done=false;
    let ans=t.replace(rowRE,r=>{i++;if(i!==ri)return r;done=true;return edits.reduce((v,e)=>replaceCellRow(v,e[0],e[1],e[2]||{}),r)});
    if(!done)throw new Error('Нет строки '+ri);
    return ans;
  }
  // Шапка: название СК, полис строго в соседних пустых графах.
  editTable(1,t=>{
    t=setTableRow(t,0,[[1,D.insurer],[3,D.polis]]);
    return setTableRow(t,1,[[1,D.culpritSK],[3,D.culpritPolis]]);
  });
  // Первая страница: таблица сведений о ТС, исходная сетка сохранена.
  editTable(3,t=>{
    const edits={
      1:[[1,D.plate]],
      4:[[1,D.model]],
      5:[[1,D.bodyType],[2,D.doors ? ' '+D.doors : '',{underline:true}]],
      6:[[1,D.startDate],[3,D.year]],
      7:[[1,D.mileage ? D.mileage+' км' : '']],
      8:[[1,D.chassis]],
      9:[[1,[D.color,D.paint].filter(Boolean).join(', ')]],
      10:[[1,D.engineModel]],
      11:[[1,D.power ? D.power+' л.с.' : ''],[2,D.volume ? D.volume+' см³' : '']],
      12:[[1,D.fuel]],
      13:[[1,'     '+D.stsSeries],[2,'     '+D.stsNumber]],
      14:[[1,D.owner]],
    };
    for(const [r,ed] of Object.entries(edits))t=setTableRow(t,Number(r),ed);
    // Уменьшаем пустую графу собственника, чтобы блок подписи оставался на первом листе.
    let ix=-1;t=t.replace(rowRE,r=>{ix++;if(ix!==14)return r;return r.replace(/<w:trHeight[^>]*\/>/, '<w:trHeight w:val="410"/>');});
    // По одной букве/цифре VIN в отдельной физической ячейке, 17 символов.
    const vin=String(D.vin||'');
    if(vin && vin.length!==17)throw new Error('VIN должен содержать 17 символов');
    t=setTableRow(t,2,[...vin].map((letter,i)=>[i+1,letter]));
    return t;
  });
  // Печатный второй лист вместо мест для заполнения ручкой.
  // Каждая из трёх строк остаётся в ИСХОДНОЙ таблице, без наложений.
  function printedRun(text,bold=false){
    return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="${bold?1:0}"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
  }
  function setPrintedCell(cell,parts) {
    const pr=(cell.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)||[''])[0];
    const open=(cell.match(/^<w:tc(?:\s[^>]*)?>/)||['<w:tc>'])[0];
    const par='<w:p><w:pPr><w:spacing w:before="30" w:after="30"/></w:pPr>'+parts.map(x=>printedRun(x[0],!!x[1])).join('')+'</w:p>';
    return open+pr+par+'</w:tc>';
  }
  function rewriteSingleCellRow(row,parts){
    const c=[...row.matchAll(cellRE)].map(x=>x[0]);
    if(c.length!==1)throw new Error('Строка печатного блока должна быть единой ячейкой, сейчас '+c.length);
    return row.replace(c[0],setPrintedCell(c[0],parts));
  }
  editTable(8,t=>{
    let i=-1;
    return t.replace(rowRE,r=>{
      i++;
      if(i===0)return rewriteSingleCellRow(r,[['1. Направление на осмотр № ',false],[D.claimNo,true]]);
      if(i===1)return rewriteSingleCellRow(r,[['Адрес осмотра: ',false],[D.address,true]]);
      if(i===2)return rewriteSingleCellRow(r,[['Осмотр проводился: ',false],[D.date+' с '+D.start+' до '+D.end,true]]);
      return r;
    });
  });
  // Дата ДТП выделяется жирным, без изменения остальных строк документа.
  editTable(9,t=>{
    let i=-1;
    return t.replace(rowRE,r=>{
      if(++i!==1)return r;
      return rewriteSingleCellRow(r,[['Дата и время ДТП (события): ',false],[D.eventDate,true]]);
    });
  });
  // Документы компетентных органов: печатать статус, а не сохранять выбор для ручки.
  editTable(9,t=>{
    let ri=-1;
    return t.replace(rowRE,r=>{
      if(++ri!==0)return r;
      const cells=[...r.matchAll(cellRE)].map(m=>m[0]);
      if(cells.length!==1)throw new Error('Нет графы статуса документов');
      const status = D.documentsStatus === 'предоставлен' ? 'ПРЕДОСТАВЛЕН' :
                     D.documentsStatus === 'не предоставлен' ? 'НЕ ПРЕДОСТАВЛЕН' : 'СВЕДЕНИЙ НЕТ';
      return r.replace(cells[0],setPrintedCell(cells[0],[
        ['Документ из компетентных органов (Извещение о ДТП): ',false],[status,true]
      ]));
    });
  });
  // Остальные поля второй и четвёртой страниц: внесение после названия графы
  // только там, где бланк не выделяет самостоятельную соседнюю ячейку.
  function addAfterText(label,value){
    if(!value)return;
    let found=false;
    xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
      if(found||!p.includes(label))return p;
      found=true;
      return p.replace('</w:p>',run(' '+value)+'</w:p>');
    });
    if(!found)console.warn('Отсутствует поле:',label);
  }
  // Заменено печатным блоком выше.
  // Заменено печатным блоком выше.
  // Заменено печатным блоком выше.
  // Заменено печатным блоком выше.
  addAfterText('Место ДТП',D.eventPlace);
  addAfterText('Состояние ТС',String(D.state||'').replace(/^\s*ТС\s+/i,''));
  addAfterText('Заключение (иное):',D.conclusion);
  // Номер акта у заголовка: заменяем в параграфе, не меняя сетку VIN.
  // Номер ставим в строке, отведенной после заголовка.
  // Номер акта в заголовке, независимо от разбиения пробелов на Word runs.
  let titleNo=false;
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(titleNo || !p.includes('______________________________'))return p;
    titleNo=true;return p.replace('</w:p>',styledText(' '+D.claimNo,{pt:11,bold:true})+'</w:p>');
  });
  // Заголовок акта: напечатанный номер без рукописной линии.
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!p.includes(D.claimNo) || !p.includes('№') || !p.includes('________'))return p;
    return p.replace(/_{6,}/g,'');
  });
  // Повреждения: общая сквозная нумерация, без отдельного раздела.
  // Для переполнения таблицы повторяем страницу повреждений перед выводами.
  function asRowCells(tr){return [...tr.matchAll(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g)].map(m=>m[0]);}
  function changeCell(tr, cix, value){
    const a=asRowCells(tr), old=a[cix];
    if(!old)return tr;
    // Убираем старый номер и остальные текстовые runs в ячейке, оставляем оформление параграфа.
    let fresh=old.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g,'');
    fresh=fresh.replace(/<\/w:p>/,run(value)+'</w:p>');
    return tr.replace(old,fresh);
  }
  // Одна объединённая широкая графа вместо Вид/характер, Место, Объём.
  // Наименование детали — жирным; описание — обычным шрифтом.
  function runDamage(value, bold) {
    return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="${bold?1:0}"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr><w:t xml:space="preserve">${escape(value)}</w:t></w:r>`;
  }
  function wideCell(old, name, descr) {
    let cell=old.replace(/<w:tcW[^>]*\/>/, '<w:tcW w:w="3506" w:type="pct"/>');
    cell=cell.replace('</w:tcPr>','<w:gridSpan w:val="4"/></w:tcPr>');
    cell=cell.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,'');
    let body='<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>';
    if(name)body+=runDamage(name, true)+(descr?runDamage(' — '+descr, false):'');
    body+='</w:p>';
    return cell.replace('</w:tc>',body+'</w:tc>');
  }
  function consolidateRow(row, name, desc) {
    const a=asRowCells(row);
    if(a.length!==9)throw new Error('Неверная строка таблицы повреждений: '+a.length);
    const wide=wideCell(a[1],name,desc);
    return row.replace(a.slice(1,5).join(''),wide);
  }
  function damageRow(row, number, item){
    let name='', desc='';
    if(item) {
      if(Array.isArray(item)) {name=item[0]||''; desc=[item[1],item[2]&&('в '+item[2]),item[3]&&('на площади '+item[3])].filter(Boolean).join(', ');}
      else if(typeof item==='string') {const m=item.split(/\s+[—–-]\s+/);name=m.shift();desc=m.join(' — ');}
      else {name=item.detail||'';desc=item.description||item.text||'';}
    }
    let r=changeCell(row,0,String(number));
    return consolidateRow(r,name,desc);
  }
  let damageTableOriginal='', footerOriginal='';
  let allTables=[...xml.matchAll(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g)];
  const target=allTables.find(m=>(m[0].match(/<w:tr(?:\s[^>]*)?>/g)||[]).length===37);
  const sigTable=allTables.find(m=>m.index>target.index && m[0].includes('Осмотр осуществил'));
  if(!target||!sigTable)throw new Error('Не найдена таблица повреждений / подписи');
  damageTableOriginal=target[0]; footerOriginal=sigTable[0];
  // Маркер листа из исходного бланка, между таблицей повреждений и подписями.
  const between=xml.slice(target.index+damageTableOriginal.length,sigTable.index);
  const sheetMarker=[...between.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)].map(m=>m[0]).find(p=>p.includes('Лист №'));
  if(!sheetMarker) throw new Error('Не найден номер листа повреждений');
  const damageRows=[...damageTableOriginal.matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map(m=>m[0]);
  if(damageRows.length!==37)throw new Error('Неожиданная структура таблицы: '+damageRows.length);
  const prefix=damageTableOriginal.slice(0,damageTableOriginal.indexOf(damageRows[0]));
  const suffix=damageTableOriginal.slice(damageTableOriginal.lastIndexOf(damageRows[36])+damageRows[36].length);
  function makeTable(start, count, final){
    const numberRows=Array.from({length:count},(_,i)=>damageRow(damageRows[2+(i%15)],start+i+1,damages[start+i]));
    // Ровно 20 строк на лист. Пустые строки оставляем без жирного шрифта.
    while(numberRows.length<15)numberRows.push(damageRow(damageRows[2+(numberRows.length%15)],'',[]));
    return prefix+[damageRows[0],consolidateRow(damageRows[1],'Наименование повреждённой детали и описание повреждения',''),...numberRows,...final?[damageRows[36]]:[]].join('')+suffix;
  }
  const firstTable=makeTable(0,Math.min(15,damages.length),damages.length<=15);
  xml=xml.replace(damageTableOriginal,firstTable);
  if(damages.length>15){
    let more='';
    for(let start=15;start<damages.length;start+=15){
      const n=Math.min(15,damages.length-start),last=start+n>=damages.length;
      // Каждый блок с 21-го пункта вынесен на следующую страницу.
      more+='<w:p><w:r><w:br w:type="page"/></w:r></w:p>'+makeTable(start,n,last)+sheetMarker+footerOriginal;
    }
    // Дополнительные страницы вставляем между третьим листом и выводами.
    const at=xml.indexOf(footerOriginal,xml.indexOf(firstTable));
    if(at<0)throw new Error('Не найден блок подписи после повреждений');
    xml=xml.slice(0,at+footerOriginal.length)+more+xml.slice(at+footerOriginal.length);
  }

  // Точные печатные галочки: сохраняем одну отметку в каждой штатной позиции.
  // ActiveX и VML целиком заменяются обычным текстовым run, пригодным для PDF.
  const CHECK_FONT='DejaVu Sans';
  const checkboxRun = yes => `<w:r><w:rPr><w:rFonts w:ascii="${CHECK_FONT}" w:hAnsi="${CHECK_FONT}"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr><w:t>${yes?'☒':'☐'}</w:t></w:r>`;
  function checkboxInCell(cell, states) {
    let n=0;
    let out=cell.replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g,
       old => `<w:t>${states[n++] ? "☒":"☐"}</w:t>`);
    if(n!==states.length)console.warn('Количество чекбоксов отличается:',n,states.length);
    return out;
  }
  function markOptions(tableIndex,rowIndex,chosenColumns) {
    editTable(tableIndex,t=>{
      let i=-1;return t.replace(rowRE,r=>{
        i++;if(i!==rowIndex)return r;
        const cells=[...r.matchAll(cellRE)].map(m=>m[0]);
        for(let j=cells.length-1;j>=0;j--){
          const obj=(cells[j].match(/<w:object(?:\s[^>]*)?>/g)||[]).length;
          if(obj)r=r.replace(cells[j],checkboxInCell(cells[j],Array.from({length:obj},()=>chosenColumns.includes(j))));
        }
        return r;
      });
    });
  }
  function markBinary(tableIndex,yes) {
    editTable(tableIndex,t=>{
      let i=-1;return t.replace(rowRE,r=>{
        i++;if(i!==0)return r;
        const cc=[...r.matchAll(cellRE)].map(m=>m[0]);
        if(cc.length!==2)return r;
        return r.replace(cc[1],checkboxInCell(cc[1],[yes,!yes]));
      });
    });
  }
  markOptions(0,0,[(CHECKS.insurance || 'КАСКО').toUpperCase()==='ОСАГО'?0:1]);
  const inspectionChoices=['первичный','дополнительный','после восст. ремонта','доп. осмотр ПЗН'];
  const inspection=String(CHECKS.inspection||D.inspType||'первичный').toLowerCase();
  const inspectionIndex=inspection.includes('пзн')?3:inspection.includes('ремонт')?2:inspection.startsWith('доп')?1:0;
  markOptions(2,0,[inspectionIndex]);
  markOptions(2,1,[String(CHECKS.place||D.mode||'выездной').toLowerCase().includes('стационар')?0:1]);
  markOptions(2,2,[]);        // фото не вместо живого осмотра
  markBinary(5, !(D.identificationDiscrepancies||[]).length);  // соответствие данным
  markBinary(6, !!(D.extraEquipment||[]).length); // дополнительное оборудование
  // Дефекты и неопределённые повреждения без указаний эксперта НЕ отмечаем.
  // Пустые строки разделов сохраняем для последующего описания.
  markOptions(8,0,[0]);    // основание №1: направление на осмотр
  const cleanliness=String(CHECKS.cleanliness||'').toLowerCase();
  // Чистота расположена на строке 4 (чистое) и строке 5 (грязное).
  // СТОА — САМОСТОЯТЕЛЬНАЯ опция в строке 5, колонке 1.
  // Ранее 'чистое ТС' ошибочно выбирало СТОА.
  markOptions(8,4,cleanliness.includes('чист')?[2]:[]);
  markOptions(8,5,[0,...(CHECKS.stoa===true?[1]:[]),...(cleanliness.includes('гряз')?[2]:[])]);
  function markRowBoxes(tableIndex,rowIndex,states) {
    editTable(tableIndex,t=>{
      let i=-1;return t.replace(rowRE,r=>{
        if(++i!==rowIndex)return r;
        let j=0;
        const result=r.replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g,()=>`<w:t>${states[j++] ? '☒':'☐'}</w:t>`);
        if(j!==states.length)console.warn('Отметки строк',tableIndex,rowIndex,j,states.length);
        return result;
      });
    });
  }
  markRowBoxes(13,22,[false,true]); // Демонтаж: Нет, положение строки может быть последним после пересборки
  // Соответствие характеру ДТП отмечается ниже по содержимому таблицы.
  // Заключение: исходные ActiveX нестабильно индексируются; окончательные отметки ниже.
  // Не ставим автоматического вывода о ремонтопригодности.
  // Отметка скрытых повреждений задаётся CHECKS.hidden или остаётся пустой.
  // Оставшиеся объекты (ниже заключение/демонтаж) заменяем штатным
  // пустым флажком без VML. Ничего не накладываем поверх оригинального.
  xml=xml.replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g, `<w:t>☐</w:t>`);
  // Пункты четвертого листа: рядом с вариантами текста добавляем явные отметки.
  function prefixParagraphContaining(textToFind,mark) {
    let matched=false;
    xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
      if(matched||!p.includes(textToFind))return p;
      matched=true;
      const start=p.indexOf('</w:pPr>');
      return start>=0?p.slice(0,start+8)+checkboxRun(mark)+p.slice(start+8):p.replace('</w:p>',checkboxRun(mark)+'</w:p>');
    });
  }
  // отметка ремонтопригодности проставлена непосредственно в таблице
  // Ремонтные воздействия КАСКО: точные ячейки после объединённой графы.
  let globalRepairIndex=0;
  xml=xml.replace(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/g,t=>{
    if(!t.includes('Наименование повреждённой детали и описание повреждения'))return t;
    let ri=-1;return t.replace(rowRE,r=>{
      ri++;if(ri<2 || ri>=17 || globalRepairIndex>=damages.length)return r;
      const k=KASKO_REPAIRS[globalRepairIndex++]||{};
      // В ОСАГО ремонтные воздействия в таблице КАСКО не проставляются.
      if(String(CHECKS.insurance||'').toUpperCase()!=='КАСКО')return r;
      const cells=[...r.matchAll(cellRE)].map(m=>m[0]);
      for(let i=cells.length-1;i>=0;i--){
        const key={2:'zamena',3:'remont',4:'okraska',5:'diagn'}[i];
        if(!key || !k[key])continue;
        // По решению эксперта от 19.11.2026: в «Ремонт, час» ставится X
        // даже без расчёта нормо-часов. Никакие часы не выдумывать.
        r=r.replace(cells[i],setCell(cells[i],'X'));
      }
      return r;
    });
  });
  // Комплектация находится в ОТДЕЛЬНОЙ плавающей таблице №4, а не в «Прочее».
  // Подтвержденные характеристики добавляем прямо в соответствующие ячейки.
  editTable(4,t=>{
    // Комплектация заполняется ниже из EQUIPMENT. Не вписываем жёсткие значения.
    const info={};
    let ri=-1;return t.replace(rowRE,r=>{
      ri++;if(!info[ri])return r;
      for(const [idx,val] of Object.entries(info[ri]))r=replaceCellRow(r,Number(idx),'  • '+val);
      return r;
    });
  });
  // Комплектация: подтверждённо отсутствующие опции зачёркиваем в штатном списке.
  function strikeFeature(label) {
    let done=false;
    editTable(4,t=>t.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g,r=>{
      if(done || !r.includes('>'+label+'</w:t>'))return r;
      done=true;
      if(/<w:rPr>/.test(r))return r.replace('<w:rPr>','<w:rPr><w:strike w:val="1"/>');
      return r.replace(/<w:r(?:\s[^>]*)?>/,m=>m+'<w:rPr><w:strike w:val="1"/></w:rPr>');
    }));
    if(!done)console.warn('Не удалось зачеркнуть:',label);
  }
  // Отсутствующие позиции зачёркиваются только по EQUIPMENT.

  // Версия 3: компактная печатная сетка первого листа вместо полей для ручки.
  function styledText(s,{strike=false,bold=true,pt=9}={}) {
    return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="${bold?1:0}"/>${strike?'<w:strike w:val="1"/>':''}<w:sz w:val="${Math.round(pt*2)}"/><w:szCs w:val="${Math.round(pt*2)}"/></w:rPr><w:t xml:space="preserve">${escape(s)}</w:t></w:r>`;
  }
  function cleanCell(c, parts, {pad=40}={}) {
    const start=c.match(/^<w:tc(?:\s[^>]*)?>/)[0];
    const prop=(c.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)||[''])[0];
    const body=parts.map(p=>styledText(p.text,p)).join('');
    return start+prop+`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="${pad}"/></w:pPr>${body}</w:p></w:tc>`;
  }
  function rewriteRowCells(table, index, spec, height) {
    let ri=-1;
    return table.replace(rowRE,row=>{
      if(++ri!==index)return row;
      const cells=[...row.matchAll(cellRE)].map(m=>m[0]);
      for(const [idx,parts] of Object.entries(spec)) {
        const k=Number(idx);if(!cells[k])continue;
        row=row.replace(cells[k],cleanCell(cells[k],parts));
      }
      if(height) row=row.replace(/<w:trHeight[^>]*\/>/,`<w:trHeight w:val="${height}"/>`);
      return row;
    });
  }
  const v=(text,pt=9)=>[{text,bold:true,pt}];
  editTable(3,t=>{
    t=rewriteRowCells(t,5,{2:v('Дверей: '+D.doors)},225);
    t=rewriteRowCells(t,7,{1:v(D.mileage?String(D.mileage)+' км':'',9)},225);
    t=rewriteRowCells(t,9,{1:v(' '+[D.color,D.paint].filter(Boolean).join(', '))},255);
    t=rewriteRowCells(t,10,{1:v(D.engineModel?'Модель/№: '+D.engineModel:'Модель/№: не установлена',8)},225);
    const powerText = D.powerKw && D.power ? `Мощность: ${D.powerKw} кВт / ${D.power} л.с.` :
                     D.powerKw ? `Мощность: ${D.powerKw} кВт` :
                     D.power ? `Мощность: ${D.power} л.с.` : 'Мощность: не установлена';
    t=rewriteRowCells(t,11,{1:v(powerText,8),2:v(D.volume?'Объём: '+D.volume+' см³':'Объём: не установлен',8)},230);
    const fuel=String(D.fuel||'').toLowerCase();
    const fuelText=fuel.includes('диз')?'Дизель':fuel.includes('бенз')?'Бензин':fuel.includes('гибрид')?'Гибрид':fuel.includes('электр')?'Электрический':'Тип не установлен';
    const injection = fuelText==='Дизель'||fuelText==='Бензин'?'Впрыск':'';
    t=rewriteRowCells(t,12,{1:v(fuelText,8),2:v(injection,8),3:v(' ',8),4:v(' ',8)},225);
    t=rewriteRowCells(t,13,{0:v('СТС',9),1:v('Серия  '+D.stsSeries,9),2:v('№  '+D.stsNumber,9)},245);
    return t;
  });
  // Каждую опцию комплектации выводим внутри её штатной ячейки.
  // Отсутствующая опция перечёркивается целиком, неизвестная не вычёркивается.
  const present=(text)=>[{text,bold:true,pt:7.5}];
  const absent=(text)=>[{text,bold:false,strike:true,pt:7.5}];
  const unknown=(text)=>[{text,bold:false,pt:7.5}];
  const mix=(...a)=>a;
  // Значения комплектации берутся из редактируемого блока EQUIPMENT, не из модели.
  const EQUIPMENT_LABELS = {
    transmission_auto:'АКПП',transmission_manual:'МКПП',central_lock:'Центральный замок',
    drive:'Привод',parking:'Парктроники перед./зад.',abs:'ABS',camera:'Камера',
    airbags:'Подушки безопасности',sunroof:'Люк',ac:'Кондиционер/климат-контроль',
    suspension:'Регулируемая подвеска',mirrors:'Зеркала эл./обогрев',lpg:'Газобаллонное оборудование',
    mirror_color:'Зеркала окрашены',trailer:'Прицепное устройство',advertising:'Наружная реклама',
    vinyl:'Виниловая плёнка',deflectors:'Дефлекторы капота / окон',foglights:'Противотуманные фары',
    wheel_liners:'Доп. подкрылки',headlight_wash:'Омыватели фар',spoiler:'Доп. спойлер',
    headlights:'Фары (тип не установлен)',steps:'Дополнительные подножки',
    windows:'Электростеклоподъёмники',bars:'Дуги защитные',handles:'Ручки окрашены',
    security:'Охранная система (доп.)',upholstery:'Обивка салона',rain_sensor:'Датчик дождя / подогрев',
    wheels:'Тип дисков не установлен',tyres:'Размер шин не установлен',
    armor:'Бронированный: не установлено',other:'Прочее'
  };
  const feature = (k) => {
    const f=EQUIPMENT[k];
    const label=(f&&f.text&&f.text!==k) ? f.text : EQUIPMENT_LABELS[k];
    if(!label) throw new Error('Нет русского названия комплектации: '+k);
    if(!f || f.status==='unknown') return unknown(label);
    return f.status==='absent' ? absent(label) : present(label);
  };
  const eq={
    1:{0:mix(feature('transmission_auto')[0],{text:' / ',pt:7.5},feature('transmission_manual')[0]),1:feature('central_lock')},
    2:{0:feature('drive'),1:feature('parking')},
    3:{0:feature('abs'),1:feature('camera')},
    4:{0:feature('airbags'),1:feature('sunroof')},
    5:{0:feature('ac'),1:feature('suspension')},
    6:{0:feature('mirrors'),1:feature('lpg')},
    7:{0:feature('mirror_color'),1:feature('trailer')},
    8:{0:feature('advertising'),1:feature('vinyl')},
    9:{0:feature('deflectors')},
    10:{0:feature('foglights'),1:feature('wheel_liners')},
    11:{0:feature('headlight_wash'),1:feature('spoiler')},
    12:{0:feature('headlights'),1:feature('steps')},
    13:{0:feature('windows'),1:feature('bars')},
    14:{0:feature('handles'),1:feature('security')},
    15:{0:feature('upholstery'),1:feature('rain_sensor')},
    16:{0:feature('wheels'),1:feature('tyres')},
    17:{0:feature('armor')},
    18:{0:feature('other')}
  };
  editTable(4,t=>{
    t=t.replace('w:tblpY="281"','w:tblpY="10"');
    for(const [ix,setting] of Object.entries(eq)) t=rewriteRowCells(t,Number(ix),setting,210);
    // Заголовок и последняя строка сохраняют высоту, остальное поджимается за счёт удаления старых переносов.
    return t;
  });
  // Более читаемый номер убытка без увеличения всех реквизитов документа.
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!p.includes('Лист №') || !p.includes(D.claimNo))return p;
    return p.replace(new RegExp('(<w:t(?:\\s[^>]*)?>)'+D.claimNo+'(<\\/w:t>)','g'),(_,a,b)=>a+D.claimNo+b)
            .replace(/<w:sz w:val="(?:16|17|18)"\/>/g,'<w:sz w:val="21"/>');
  });

  // Номер убытка возле обозначения листа на каждой странице, включая дополнительные.
  let leafNo=0;
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!p.includes('Лист №'))return p;
    leafNo++;
    return p.replace('Лист №__ Акта осмотра №________________', (process.env.AKT_TEST === '1' ? 'ТЕСТОВЫЙ ОБРАЗЕЦ | ' : '')+'Лист №'+leafNo+' Акта осмотра №'+D.claimNo);
  });
  // Номера убытка в заголовках листов печатаем 10.5 пт полужирным.
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!p.includes('Лист №') || !p.includes(D.claimNo))return p;
    return p.replace(/<w:sz w:val="(?:16|17|18|21)"\/>/g,'<w:sz w:val="18"/>');
  });
  if(leafNo<4)throw new Error('Не найдены номера листов: '+leafNo);
  // Удаляем избыточные ПУСТЫЕ строки на листах 1 и 2,
  // освобождаем место для нижнего блока ФИО/подписи без масштабирования текста.
  function trimEmptyTable(tableIndex, n) {
    editTable(tableIndex,table=>{
      const rr=[...table.matchAll(rowRE)].map(m=>m[0]);
      for(let i=rr.length-1;i>=1 && n>0;i--){
        let text=[...rr[i].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(m=>m[1].trim()).join('');
        if(!text){table=table.replace(rr[i],'');n--;}
      }
      return table;
    });
  }
  trimEmptyTable(4,2);trimEmptyTable(5,2);
  trimEmptyTable(9,2);trimEmptyTable(10,2);trimEmptyTable(15,7);
  editTable(16,t=>t.replace(/<w:trHeight w:val="(\d+)"/g,(m,n)=>`<w:trHeight w:val="${Math.round(Number(n)*0.73)}"`));

  // Четвёртый лист начинается целым блоком заключения.
  const conclusionTable=[...xml.matchAll(tblRE)].find(m=>m[0].includes('Характер описанных повреждений'));
  if(conclusionTable) xml=xml.slice(0,conclusionTable.index)+'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'+xml.slice(conclusionTable.index);
  // Владелец информирован, но акт не подписан им.
  xml=xml.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g,tr=>{
    if(!tr.includes('Собственник/Доверенное лицо'))return tr;
    const cells=[...tr.matchAll(/<w:tc[ >][\s\S]*?<\/w:tc>/g)];
    if(cells.length<2)return tr;
    const c=cells[1][0];
    return tr.replace(c,c.replace(/<\/w:p>/,run('С актом ознакомлен, акт не подписан')+'</w:p>'));
  });

  // Подпись в графе эксперта каждой из четырёх страниц.
  let signed = 0;
  xml = xml.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g, tr => {
    if (!tr.includes('Осмотр осуществил')) return tr;
    const cells = [...tr.matchAll(/<w:tc[ >][\s\S]*?<\/w:tc>/g)];
    if (cells.length < 3) return tr;
    const fio = cells[cells.length - 2][0];
    const sign = cells[cells.length - 1][0];
    let r = tr.replace(fio, fio.replace(/<\/w:p>/, run(EXPERT) + '</w:p>'));
    r = r.replace(sign, sign.replace(/<\/w:p>/, `<w:r>${DRAWING}</w:r></w:p>`));
    signed++;
    return r;
  });
  if (SIGN_EXPERT && signed < 4) throw new Error(`Найдено блоков подписи ${signed} вместо 4`);
  if (!SIGN_EXPERT) xml = xml.replace(/<w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r>/g, '');
  // Вместо схлопывания таблиц до 35% высоты убираем лишние интервалы.
  // Шрифт исходного бланка приводим к умеренным 9 pt там, где он был чрезмерным.
  // Исходные заголовки не изменяем; введённые данные напечатаны 8.5 pt без жирного.
  xml=xml.replace(/<w:sz w:val="2[24]"\/>/g,'<w:sz w:val="18"/>');
  xml=xml.replace(/<w:szCs w:val="2[24]"\/>/g,'<w:szCs w:val="18"/>');
  // Текст о восстановительном ремонте на листе 4: не допускаем наложения строк.
  const p4 = xml.indexOf('Выводы по результатам осмотра по претензии');
  if(p4>=0){
    const a=xml.slice(0,p4), b=xml.slice(p4);
    xml=a+b;
  }
  // Четвёртый лист: слегка уменьшаем текст внутри заключения без изменения масштаба таблиц.
  const conclPos=xml.indexOf('Характер описанных повреждений');
  if(conclPos>=0) {
    let head=xml.slice(0,conclPos),tail=xml.slice(conclPos);
    tail=tail.replace(/<w:sz w:val="(?:18|19|20|21|22)"\/>/g,'<w:sz w:val="18"/>');
    tail=tail.replace(/<w:szCs w:val="(?:18|19|20|21|22)"\/>/g,'<w:szCs w:val="18"/>');
    xml=head+tail;
  }
  // На последнем листе подписи уже обозначены колонками таблицы.
  // Отдельная подпись под таблицей давала пустую пятую страницу.
  const paraMatches=[...xml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g)];
  const tailCaption=paraMatches.filter(m=>m[0].includes('ФИО') && m[0].includes('Подпись')).pop();
  if(tailCaption) xml=xml.slice(0,tailCaption.index)+xml.slice(tailCaption.index+tailCaption[0].length);
  // Новый вид подписей по старой форме: без рамок вокруг строк.
  // Сохраняем структурные ячейки для надёжной привязки ФИО и рисунка,
  // выводим лишь нижние линии в графах ФИО и подписи.
  function noFrameSignatures(table) {
    if(!table.includes('Собственник/Доверенное лицо') || !table.includes('Осмотр осуществил')) return table;
    const none='<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/>';
    if(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/.test(table)) table=table.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/,'<w:tblBorders>'+none+'</w:tblBorders>');
    else table=table.replace('</w:tblPr>','<w:tblBorders>'+none+'</w:tblBorders></w:tblPr>');
    return table.replace(rowRE,row=>{
      const cells=[...row.matchAll(cellRE)].map(m=>m[0]);
      if(cells.length<3)return row;
      for(let j=cells.length-1;j>=1;j--){
        let c=cells[j];
        c=c.replace(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/g,'');
        const bottom='<w:tcBorders><w:bottom w:val="single" w:sz="5" w:color="555555"/></w:tcBorders>';
        c=c.replace('</w:tcPr>',bottom+'</w:tcPr>');
        row=row.replace(cells[j],c);
      }
      return row;
    });
  }
  xml=xml.replace(tblRE,noFrameSignatures);
  // Четвёртый лист: место под заключение расширяется содержимым,
  // лишние пустые графы для ручного дописывания не нужны.
  xml=xml.replace(tblRE,t=>{
    if(!t.includes('Заключение (иное):') || !t.includes('Скрытые повреждения'))return t;
    const rr=[...t.matchAll(rowRE)].map(m=>m[0]);
    for(let i=rr.length-1;i>=0;i--){
      if(i<5)continue;
      const txt=[...rr[i].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(m=>m[1].trim()).join('');
      if(!txt)t=t.replace(rr[i],'');
    }
    return t;
  });
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!p.includes('Состояние ТС'))return p;
    return p.replace(/_{6,}/g,'');
  });
  // Подчёркивания ручного ввода вместо уже заполненных значений не нужны.
  // Правим только параграф номера акта и направлений, не стирая другие поля.
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
    if(!(p.includes(D.claimNo) && (p.includes('Направление на осмотр') || p.includes('АКТ ОСМОТРА'))))return p;
    return p.replace(/_{4,}/g,'');
  });
  // Переносим заголовок основания вместе со вторым листом.
  const basisAt=xml.indexOf('Основание для осмотра');
  if(basisAt>=0){
    const matches=[...xml.slice(0,basisAt).matchAll(/<w:p(?:\s[^>]*)?>/g)];
    const pAt=matches.length?matches[matches.length-1].index:-1;
    if(pAt>=0)xml=xml.slice(0,pAt)+'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'+xml.slice(pAt);
  }
  // Строго: лист с повреждениями ВСЕГДА начинается с новой страницы (лист 3).
  // Разрыв ставится до всей таблицы, а не внутри её первой строки.
  const damageHead=damages[0].detail;
  const dmgHeadAt=xml.indexOf(damageHead);
  if(dmgHeadAt<0) throw new Error('Не найден заголовок листа повреждений');
  const dmgTableAt=xml.lastIndexOf('<w:tbl>',dmgHeadAt);
  if(dmgTableAt<0) throw new Error('Не найдена таблица повреждений');
  // removed duplicated damage break

  // Динамические строки в двух разделах первого листа.
  // При отсутствии записей оставляем 1 пустую строку и исходную отметку.
  // При наличии записей автоматически увеличиваем высоту ячейки.
  function renderFlexibleSection(tableIndex, entries, title) {
    let did=false;
    editTable(tableIndex, table => {
      did=true;
      const rows=[...table.matchAll(rowRE)].map(x=>x[0]);
      if(rows.length<2) throw new Error('Нет строк в разделе '+title);
      let row=rows[1].replace(/<w:trHeight[^>]*\/>/g,'');
      const cells=[...row.matchAll(cellRE)].map(x=>x[0]);
      const c=cells[0];
      const tcPr=(c.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)||[''])[0];
      const opening=(c.match(/^<w:tc(?:\s[^>]*)?>/)||['<w:tc>'])[0];
      const line = (v,i) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="220" w:lineRule="auto"/></w:pPr>'
        +'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>'
        +'<w:t xml:space="preserve">'+escape((i+1)+'. '+String(v))+'</w:t></w:r></w:p>';
      const body=entries.length ? entries.map(line).join('') : '<w:p/>';
      row=row.replace(c,opening+tcPr+body+'</w:tc>');
      const start=table.indexOf(rows[0]);
      const end=table.lastIndexOf(rows[rows.length-1])+rows[rows.length-1].length;
      return table.slice(0,start)+rows[0]+row+table.slice(end);
    });
    if (!did) throw new Error('Не удалось изменить раздел '+title);
  }
  renderFlexibleSection(5, D.identificationDiscrepancies||[], 'Идентификация');
  renderFlexibleSection(6, D.extraEquipment||[], 'Доп. оборудование');

  // Убираем плавающую привязку блока ДТП: после него не должно оставаться
  // искусственного пустого промежутка, за который расплачивается подпись.
  xml=xml.replace(tblRE,t => t.includes('Дата и время ДТП (события)') ? t.replace(/<w:tblpPr[^>]*\/>/g,'') : t);
  // Две колонки для длинных перечней. Номера идут слева сверху вниз, затем справа.
  // Короткие перечни остаются обычными полноширинными параграфами.
  function balancedColumns(entries, fontHalfPoints, lineTwips) {
    const paragraph=(item, i)=>'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="'+lineTwips+'" w:lineRule="auto"/></w:pPr>'
       +'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="'+fontHalfPoints+'"/><w:szCs w:val="'+fontHalfPoints+'"/></w:rPr>'
       +'<w:t xml:space="preserve">'+escape((i+1)+'. '+String(item))+'</w:t></w:r></w:p>';
    if(entries.length<6) return entries.map(paragraph).join('');
    const estimate=v=>Math.max(1,Math.ceil((String(v).length+4)/51));
    let leftCount=1, best=Infinity;
    for(let k=1;k<entries.length;k++) {
      let l=entries.slice(0,k).reduce((a,b)=>a+estimate(b),0);
      let r=entries.slice(k).reduce((a,b)=>a+estimate(b),0);
      const cost=Math.abs(l-r)+(k===1||k===entries.length-1?1:0);
      if(cost<best){best=cost;leftCount=k;}
    }
    const sub= (parts,offset)=>parts.map((v,i)=>paragraph(v,i+offset)).join('');
    const colW=5510;
    const none='<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>';
    const cell=(parts,offset)=>'<w:tc><w:tcPr><w:tcW w:w="'+colW+'" w:type="dxa"/><w:tcMar><w:left w:w="60" w:type="dxa"/><w:right w:w="115" w:type="dxa"/></w:tcMar><w:tcBorders>'+none+'</w:tcBorders></w:tcPr>'+sub(parts,offset)+'</w:tc>';
    return '<w:tbl><w:tblPr><w:tblW w:w="11020" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>'+none+'</w:tblBorders></w:tblPr>'
     +'<w:tblGrid><w:gridCol w:w="5510"/><w:gridCol w:w="5510"/></w:tblGrid>'
     +'<w:tr>'+cell(entries.slice(0,leftCount),0)+cell(entries.slice(leftCount),leftCount)+'</w:tr></w:tbl><w:p/>';
  }
  // Стресс-тест: дефекты непосредственно в разделе на листе 2.
  // Никаких приложений, ссылки "см. приложение" и переносов номера ущерба.
  if (Array.isArray(D.defects) && D.defects.length) {
    const defectTitle = 'Дефекты эксплуатации транспортного средства';
    let done = false;
    const inlineParagraph = (value, bold=false) => '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="195" w:lineRule="auto"/><w:keepTogether/></w:pPr>'
      + '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b w:val="'+(bold?'1':'0')+'"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">'+escape(value)+'</w:t></w:r></w:p>';
    xml = xml.replace(tblRE,table=>{
      if (done || !table.includes(defectTitle)) return table;
      done = true;
      const rows=[...table.matchAll(rowRE)].map(m=>m[0]);
      if (rows.length<2) throw new Error('В разделе дефектов нет строк');
      // Ставим отметку «Да» для тестового списка.
      let header = rows[0];
      const hc=[...header.matchAll(cellRE)].map(m=>m[0]);
      if (hc.length>1) {
        const ch=hc[1].replace('☐', '☒');
        header=header.replace(hc[1],ch);
      }
      // Убираем фиксированную высоту исходной пустой строки и объединяем
      // физические ячейки описания, оставляя привычную табличную сетку.
      let details=rows[1].replace(/<w:trHeight[^>]*\/>/g,'');
      const cells=[...details.matchAll(cellRE)].map(m=>m[0]);
      if(cells.length<1)throw new Error('Нет ячейки для дефектов');
      const c=cells[0], tcPr=(c.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)||[''])[0];
      const ct=c.match(/^<w:tc(?:\s[^>]*)?>/)[0];
      const content=balancedColumns(D.defects, 18, 195);
      const replacement=ct+tcPr+content+'</w:tc>';
      details=details.replace(c,replacement);
      // Удаляем остальные пустые строки раздела, иначе второй лист переполняется.
      const first=table.indexOf(rows[0]);
      const last=table.lastIndexOf(rows[rows.length-1])+rows[rows.length-1].length;
      return table.slice(0,first)+header+details+table.slice(last);
    });
    if(!done)throw new Error('Не найден раздел дефектов');
    // У смежного раздела оставляем место под возможные неопределённые повреждения,
    // но не все семь пустых линеек. Иначе таблица выталкивает подпись на другой лист.
    let near=false;
    xml=xml.replace(tblRE,table=>{
      if(near || !table.includes('Повреждения, находящиеся в зоне аварийных'))return table;
      near=true;
      const rows=[...table.matchAll(rowRE)].map(m=>m[0]);
      for(let i=rows.length-1;i>=3;i--) table=table.replace(rows[i],'');
      return table;
    });
  }

  // Отдельный динамический список неопределённых повреждений.
  // Встраиваем в собственный раздел прямо на 2-м листе, как дефекты эксплуатации.
  // Пустой массив оставляет исходный раздел без галочек и заполнения.
  if (Array.isArray(D.uncertainDamages) && D.uncertainDamages.length) {
    let done=false;
    const pLine=(v,i)=>'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="185" w:lineRule="auto"/></w:pPr>'
     +'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
     +'<w:t xml:space="preserve">'+escape(String(i+1)+'. '+v)+'</w:t></w:r></w:p>';
    xml=xml.replace(tblRE,table=>{
      if(done || !table.includes('Повреждения, находящиеся в зоне аварийных'))return table;
      done=true;
      let rr=[...table.matchAll(rowRE)].map(m=>m[0]);
      if(rr.length<2)throw new Error('Не найдены строки неопределённых повреждений');
      let head=rr[0];
      const hCells=[...head.matchAll(cellRE)].map(m=>m[0]);
      if(hCells.length<2)throw new Error('Отсутствуют галочки неопределённых повреждений');
      const marked=hCells[1].replace(/<w:t>☐<\/w:t>/,'<w:t>☒</w:t>');
      head=head.replace(hCells[1],marked);
      let details=rr[1].replace(/<w:trHeight[^>]*\/>/g,'');
      const cc=[...details.matchAll(cellRE)].map(m=>m[0]);
      if(!cc.length)throw new Error('Нет строки для неопределённых повреждений');
      const c=cc[0], tcPr=(c.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)||[''])[0];
      const open=(c.match(/^<w:tc(?:\s[^>]*)?>/)||['<w:tc>'])[0];
      details=details.replace(c,open+tcPr+balancedColumns(D.uncertainDamages, 16, 185)+'</w:tc>');
      const a=table.indexOf(rr[0]);
      const b=table.lastIndexOf(rr[rr.length-1])+rr[rr.length-1].length;
      return table.slice(0,a)+head+details+table.slice(b);
    });
    if(!done)throw new Error('Раздел неопределённых повреждений не найден');
  }
  // Минимальный равномерный отступ слева и справа от галочек в блоках
  // выбора на первом листе. Сами ячейки и сетку бланка не перестраиваем.
  let ti=-1;
  xml=xml.replace(tblRE,table=>{
    ti++;
    if(ti!==0 && ti!==2)return table;
    return table.replace(/<w:t>\s*([☐☒])\s*<\/w:t>/g,(_m,sign)=>'<w:t xml:space="preserve">\u00a0'+sign+'\u00a0</w:t>');
  });
  // Заголовок таблицы ДТП нельзя оставлять внизу листа 2 без первой записи.
  // Keep with next на двух строках заголовка заставляет Word переносить их вместе.
  xml=xml.replace(tblRE, t=>{
    if(!t.includes('Наименование повреждённой детали') || !t.includes('КАСКО')) return t;
    const rr=[...t.matchAll(rowRE)].map(m=>m[0]);
    if(rr.length<3) return t;
    const keep=(row)=>row.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
      if(p.includes('<w:keepNext'))return p;
      if(p.includes('<w:pPr>'))return p.replace('<w:pPr>','<w:pPr><w:keepNext/>');
      return p.replace(/(<w:p(?:\s[^>]*)?>)/,'$1<w:pPr><w:keepNext/></w:pPr>');
    });
    t=t.replace(rr[0],keep(rr[0]));
    t=t.replace(rr[1],keep(rr[1]));
    return t;
  });
  // Разрыв перед целой таблицей ДТП, а не перед её первой ячейкой.
  const dmgStart=[...xml.matchAll(tblRE)].find(m=>m[0].includes('Наименование повреждённой детали и описание повреждения'));
  if(!dmgStart)throw new Error('Нет таблицы ДТП для разрыва страницы');
  xml=xml.slice(0,dmgStart.index)+'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'+xml.slice(dmgStart.index);
  // Убираем печатную инструкцию про варианты состояния, которые уже заполнены.
  xml=xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>p.includes('на ходу, не на ходу, требует ремонта')?p.replace(/\([^()]*на ходу, не на ходу[^()]*\)/g,''):p);
  // Отметка отсутствия демонтажных работ в первом (и единственном) блоке.
  xml=xml.replace(tblRE,t=>{
    if(!t.includes('Осмотр проводился с проведением демонтажных работ'))return t;
    let n=0;
    return t.replace(/<w:t>☐<\/w:t>/g, m=> {n++; return n===2?'<w:t>☒</w:t>':m;});
  });
  // На листе 4 убираем рукописные линии в разделах состояния и заключений.
  // Сохраняем обязательный текст о восстановительном ремонте и согласиях.
  const p4Begin=xml.indexOf('Характер описанных повреждений');
  if(p4Begin>=0) {
    const front=xml.slice(0,p4Begin), tail=xml.slice(p4Begin);
    xml=front+tail.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g,p=>{
      if(!p.includes('Состояние ТС') && !p.includes('Устранение недостатков восстановительного ремонта'))return p;
      return p.replace(/_{5,}/g,'');
    });
  }
  // Остаточные текстовые подчёркивания раздела недостатков ремонта удаляются,
  // включая разбитые по отдельным Word runs. Сам обязательный текст сохраняется.
  const conclusionStart=xml.indexOf('Характер описанных повреждений');
  if(conclusionStart>=0) {
    const head=xml.slice(0,conclusionStart), tail=xml.slice(conclusionStart);
    xml=head+tail.replace(/_{2,}/g,'');
  }
  // Итоговые отметки в разделе выводов после замены ActiveX.
  // Работаем по тексту строк, а не по нестабильным индексам Word-таблиц.
  function markConclusionRow(table, label, state) {
    if (state !== true && state !== false) return table;
    return table.replace(rowRE, row => {
      if (!row.includes(label)) return row;
      let n=0;
      return row.replace(/<w:t>☐<\/w:t>/g, m => {
        n++;
        return n===(state?1:2)?'<w:t>☒</w:t>':m;
      });
    });
  }
  xml=xml.replace(tblRE, table => {
    if(table.includes('Характер описанных повреждений')) {
      return markConclusionRow(table,'Характер описанных повреждений',CHECKS.match);
    }
    if(table.includes('Скрытые повреждения, возможно наличие')) {
      table=markConclusionRow(table,'ТС подлежит ремонту',CHECKS.repairable);
      table=markConclusionRow(table,'Скрытые повреждения, возможно наличие',CHECKS.hidden);
    }
    return table;
  });
  // Заголовок «КАСКО» над графами ремонтных воздействий сохраняется
  // в обоих режимах страхования; для ОСАГО графы остаются без отметок.
  xml=xml.replace('Наименование НЭ','ООО «Фаворит»');
  fs.writeFileSync(xmlPath, xml);
  const relpath = path.join(tmp, 'word/_rels/document.xml.rels');
  let rel = fs.readFileSync(relpath, 'utf8');
  rel = rel.replace('</Relationships>', '<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/signature.png"/></Relationships>');
  fs.writeFileSync(relpath,rel);
  let typesPath=path.join(tmp,'[Content_Types].xml');
  let types=fs.readFileSync(typesPath,'utf8');
  if(!types.includes('Extension="png"')) types=types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  fs.writeFileSync(typesPath,types);
  fs.mkdirSync(path.join(tmp,'word/media'), {recursive:true});
  fs.copyFileSync(SIG,path.join(tmp,'word/media/signature.png'));
  fs.rmSync(OUT_FILE,{force:true});
  cp.execFileSync('zip',['-qr',OUT_FILE,'.'],{cwd:tmp});
  console.log(OUT_FILE);
} finally { fs.rmSync(tmp,{recursive:true,force:true}); }
