// ТЕСТОВЫЕ СЦЕНАРИИ ИФЛ ФАВОРИТ. ВСЕ ДАННЫЕ ВЫМЫШЛЕННЫЕ.
// Использование: node test_imush_cases.js [zaliv|pozhar|total]
// Требуется: npm package docx, находящийся рядом imush_engine.js.
// Не меняет движок. Документ записывается в текущую рабочую папку.
const path = require('path');
const build = require('./imush_engine.js');
const kind = process.argv[2] || 'zaliv';
const root = {
  actNo: 'TEST-IFL-2026-001', date: '19.11.2026', time: '11:00',
  claimDate: '19.11.2026', owner: 'Иванов Иван Иванович (ТЕСТ)',
  address: 'ТЕСТОВЫЙ АДРЕС, г. Йошкар-Ола, ул. Примерная, д. 10',
  participant: {fio: 'Иванов Иван Иванович',role:'собственник'},
  eventDate: '18.11.2026', eventTime: '19:30',
  measures: 'ТЕСТ: предприняты меры по локализации повреждений',
};
const testCases = {
  zaliv: {
    ...root, actNo:'TEST-IFL-ZALIV-001', cause:'sosedi',
    eventDescr:'ТЕСТ: проникновение воды из вышерасположенной квартиры, следы залития кухни.',
    flat:{area:'54 м²',rooms:'2',doors:'4',windows:'3',year:'2010',repairYear:'2022',
      totalPct:'не определялся',
      el:{potolok:{mat:'Железобетонное перекрытие',dmg:[['ТЕСТ: следы увлажнения в проекции кухни','2,0 м²']]}}},
    rooms:[{
      name:'Кухня',location:'Квартира № 4, 2 этаж (ТЕСТ)',size:'4 × 3 × 2,7',area:'12',
      doors:'1 шт., 0,9 × 2,0 м',windows:'1 шт., 1,4 × 1,3 м',repairYear:'2022',
      plan:{
        openings:[
          {type:'window',wall:'top',from:1.1,w:1.4},
          {type:'door',wall:'right',from:0.5,w:0.9,open:'in',hinge:'start'}
        ],
        zones:[{x:0.8,y:0.1,w:2.0,h:1.0}],
        marks:[{x:1.6,y:0.4,type:'leak',text:'место проявления протечки на потолке'}]
      },
      el:{
        potolok:{descr:'Следы залития, вздутие и отслоение окраски у окна',items:[['Окрашенная поверхность','2,0 м²']]},
        steny:{descr:'Локальные следы увлажнения оконного откоса',items:[['Штукатурка, окраска','0,8 м²']]}
      }
    }]
  },
  pozhar: {
    ...root,actNo:'TEST-IFL-POZHAR-002',cause:'pozhar',
    eventDescr:'ТЕСТ: локальное возгорание в кухне жилого дома.',
    scheme:{north:true,note:'ТЕСТ: очаг пожара отмечен на плане строения',
      objects:[
        {name:'Жилой дом',x:0,y:0,w:8,h:6,
          openings:[{type:'window',wall:'top',from:2,w:1.4},
                    {type:'door',wall:'bottom',from:1.5,w:0.9,open:'in',hinge:'start'}],
          marks:[{x:3,y:2,type:'fire',text:'предполагаемый очаг'}]},
        {name:'Баня',x:11,y:1,w:4,h:3,
          openings:[{type:'door',wall:'bottom',from:1,w:0.8}]}
      ],
      gaps:[{a:'Жилой дом',b:'Баня',m:3}]},
    buildings:[{type:'Жилой дом',purpose:'Жилое',doors:'2',windows:'4',
      size:'8 × 6 × 3',year:'2005',repairYear:'2021',
      p8:'partial',
      el:{steny:{mat:'Дерево',dmg:[['ТЕСТ: следы термического воздействия','3 м²']]},
          perekr:{mat:'Дерево',dmg:[['ТЕСТ: локальное повреждение потолочного перекрытия','2 м²']]}}}]
  },
  total: {
    ...root,actNo:'TEST-IFL-TOTAL-003',cause:'pozhar',
    eventDescr:'ТЕСТ: полная гибель хозяйственной постройки в результате пожара.',
    scheme:{north:true,objects:[{name:'Сарай',x:0,y:0,w:6,h:4,hatch:true,
      openings:[{type:'door',wall:'bottom',from:2,w:0.9}],
      marks:[{x:3,y:2,type:'fire',text:'очаг'}]}]},
    buildings:[{type:'Сарай',purpose:'Хозяйственное',doors:'1',windows:'0',
      size:'6 × 4 × 2,5',year:'2012',repairYear:'—',total:true}]
  }
};
if(!Object.hasOwn(testCases,kind)) throw new Error('Допустимые сценарии: zaliv, pozhar, total');
const D=testCases[kind];
// Подпись отключена в тестовом примере, чтобы не хранить факсимиле в публичном репозитории.
build({D,EXPERT:'Андреев А.Г.',SIGN_EXPERT:false,
  OUT_FILE:path.resolve(process.cwd(), 'TEST_IFL_'+kind+'.docx')
}).catch(e=>{console.error(e);process.exitCode=1});
