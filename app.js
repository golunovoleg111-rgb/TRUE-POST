'use strict';
const buildNumber=new URL(document.currentScript.src).searchParams.get('v');
document.getElementById('buildNumber').textContent='Сборка №'+buildNumber;
const $=id=>document.getElementById(id), esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])), norm=v=>String(v??'').trim().toLowerCase().replace(/ё/g,'е').replace(/\s+/g,' ');
const blank={version:1,catalog:[],aliases:{},settings:{},format:{}};
let db=JSON.parse(localStorage.getItem('true-post-v1')||JSON.stringify(blank)),boxes=[];
function save(next){localStorage.setItem('true-post-v1',JSON.stringify(next));db=next;stats()}
let catalogIndex;
function stats(){catalogIndex=createCatalogIndex(db.catalog);catalogIndex.aliases=new Map([...Object.entries(warehouseProducts),...Object.entries(db.aliases)].map(([key,product])=>[productAliasKey(key),product]));catalogIndex.names=new Set([...catalogIndex.aliases.keys(),...db.catalog.map(x=>productAliasKey(x.product))]);showDirectory();$('products').innerHTML=[...new Set(db.catalog.map(x=>x.product))].map(v=>'<option value="'+esc(v)+'">').join('')}
function msg(v){
  $('status').textContent='';
  $('wbStatus').textContent='';
  const target=$('wbStatus').closest('[data-tab]').hidden?'status':'wbStatus';
  $(target).textContent=v;
}
function on(id,fn,event='click'){$(id).addEventListener(event,async e=>{e.preventDefault();msg('');try{await fn(e)}catch(err){msg(err.message)}})}
function fail(v){throw Error(v)}
function download(name,data,type){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([data],{type}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),30000)}
function tab(name){document.querySelectorAll('section').forEach(s=>s.hidden=s.dataset.tab!==name);document.querySelectorAll('nav button').forEach(b=>{b.classList.toggle('active',b.textContent===name);b.setAttribute('aria-current',b.textContent===name?'page':'false')});msg('')}
for(const s of document.querySelectorAll('section')){const b=document.createElement('button');b.textContent=s.dataset.tab;b.onclick=()=>tab(b.textContent);document.querySelector('nav').append(b)}tab('Сборочное задание');stats();
async function workbook(data){const w=new ExcelJS.Workbook();await w.xlsx.load(data);return w}
function grid(s){const rows=[];s.eachRow(r=>rows.push(r.values.slice(1).map((_,i)=>r.getCell(i+1).master.address===r.getCell(i+1).address?r.getCell(i+1).text.trim():'')));return rows}
async function fileBook(input){const f=input.files[0];if(!f)fail('Выберите файл XLSX.');return workbook(await f.arrayBuffer())}
async function writeBook(name,rows,assembly=false){
  const w=new ExcelJS.Workbook(),s=w.addWorksheet('Sheet1');
  let previousBox;
  for(const values of rows){
    const row=s.addRow(assembly?values.slice(0,6):values);
    if(!assembly)continue;
    if(values.length===1){
      s.mergeCells(row.number,1,row.number,4);
      row.font={bold:true};
      row.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD5E4D7'}};
      previousBox=undefined;
    }else if(row.number>1&&values[6]){
      // A note moves with the barcode cell when a row is inserted or moved.
      row.getCell(1).note='WB-box:'+values[6];
      if(previousBox&&previousBox!==values[6]){
        for(let i=1;i<=6;i++)row.getCell(i).border={top:{style:'thin',color:{argb:'FFB9C5B9'}}};
      }
      previousBox=values[6];
    }
  }
  const header=s.getRow(1);
  header.font={bold:true};
  if(assembly){
    header.alignment={wrapText:true,vertical:'middle'};
    header.height=29;
  }
  s.columns.forEach((c,i)=>c.width=(assembly?[17,13,18,17,8,22]:[24,20,24])[i]||22);
  s.views=[{state:'frozen',ySplit:1}];
  download(name,await w.xlsx.writeBuffer(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

on('catalog',async()=>{const w=await fileBook($('catalog'));function choose(i){const rows=grid(w.worksheets[i]);if(!rows.length)fail('Пустой лист');let head=rows.findIndex(r=>r.some(v=>/баркод|штрихкод|barcode/i.test(v)));if(head<0)head=0;const h=rows[head],keys=['barcode','product','color','size'],labels=['Баркод','Изделие / артикул','Цвет','Размер'],patterns=[/баркод|штрихкод|barcode/i,/изделие|артикул продавца|номенклатура|наименование/i,/цвет/i,/размер/i];$('mapping').innerHTML='<label>Лист<select id="sheetpick">'+w.worksheets.map((s,j)=>'<option '+(j===i?'selected':'')+'>'+esc(s.name)+'</option>').join('')+'</select></label>'+keys.map((k,j)=>'<label>'+labels[j]+'<select id="m'+k+'">'+(k==='color'?'<option value="-1" '+(!h.some(x=>/цвет/i.test(x))?'selected':'')+'>Цвет из артикула</option>':'')+h.map((v,n)=>'<option value="'+n+'" '+(n===h.findIndex(x=>patterns[j].test(x))?'selected':'')+'>'+esc(v)+'</option>').join('')+'</select></label>').join('')+'<button id="applycat">Сохранить перечень</button>';$('sheetpick').onchange=e=>choose(e.target.selectedIndex);on('applycat',()=>{const indices=keys.map(k=>+$('m'+k).value);if(new Set(indices).size!==4)fail('Выберите разные столбцы или режим «Цвет из артикула»');const catalog=rows.slice(head+1).filter(r=>r.some(Boolean)).map((r,i)=>{const item=Object.fromEntries(keys.map((k,j)=>[k,String(r[indices[j]]||'').trim()]));if(indices[2]===-1){const m=item.product.match(/^(.+)_([^_]+?)(_[пp])?$/i);if(!m)fail('Не удалось выделить цвет из артикула '+item.product);item.article=item.product;item.product=m[1].trim();item.color=m[2].trim();if(/^\*?\s*[пp]$/i.test(item.color)){const split=item.product.lastIndexOf('_');if(split>0){item.color=item.product.slice(split+1)+'_'+item.color;item.product=item.product.slice(0,split)}}}if(!Object.values(item).every(Boolean))fail('Неполная позиция в строке '+(head+i+2));return item});if(!catalog.length)fail('Перечень пуст');save({...db,catalog,format:{columns:indices}});$('mapping').replaceChildren();msg('Перечень сохранён')})}choose(0)},'change');
on('alias',e=>{const f=e.target.elements;save({...db,aliases:{...db.aliases,[norm(f.abbr.value)]:f.product.value.trim()}});e.target.reset()},'submit');
on('backup',()=>download('true-post.json',JSON.stringify(db,null,2),'application/json'));
on('restore',async()=>{const d=JSON.parse(await $('restore').files[0].text());if(d.version!==1||!Array.isArray(d.catalog)||!d.catalog.every(v=>v&&['barcode','product','color','size'].every(k=>typeof v[k]==='string'&&v[k]))||![d.aliases,d.settings,d.format].every(o=>o&&typeof o==='object'&&!Array.isArray(o))||!Object.values(d.aliases).every(v=>typeof v==='string'))fail('Несовместимая база');if(confirm('Заменить текущую базу данными из файла?'))save(d)},'change');
on('parse', async()=>{
  const button=$('parse'), status=$('parseStatus'), result=[];
  const sources=layoutSources().map(source=>({...source,fill:sourceFill(source)}));
  const urls=sources.filter(source=>source.url), files=sources.filter(source=>source.file);
  const counts=[];
  async function collect(buffer,source){
    const records=await readLayout(buffer,source.label,source.fill);
    result.push(...records);
    const color=fillOptions.find(([value])=>value===source.fill)?.[1]||'#'+source.fill;
    counts.push(source.label+' · '+color+' · коробок: '+records.length);
  }
  button.disabled=true; resetAssemblyBoxCodes(); $('assemblyBoxCodesPanel').hidden=true; boxes=[]; $('review').replaceChildren();
  const show=t=>status.textContent=t;
  try {
    if(!urls.length&&!files.length)fail('Вставьте ссылку или выберите XLSX.');
    $('layoutLoading').hidden=false;
    for(const [i,source] of urls.entries()){
      const url=new URL(source.url), id=url.pathname.match(/^\/spreadsheets\/d\/([\w-]+)/)?.[1];
      if(url.hostname!=='docs.google.com'||!id)fail('Нужна ссылка на Google Таблицу.');
      const gid=new URLSearchParams(url.hash.slice(1)).get('gid')||url.searchParams.get('gid');
      const link='https://docs.google.com/spreadsheets/d/'+id+'/export?format=xlsx'+(gid?'&gid='+encodeURIComponent(gid):'');
      show('Загружаю склад '+(i+1)+' из '+urls.length+'…');
      let buffer;
      try{
        const response=await fetch(link,{signal:AbortSignal.timeout(240000)});
        if(!response.ok)throw Error();
        const reader=response.body.getReader(),chunks=[];let bytes=0;
        for(;;){const {done,value}=await reader.read();if(done)break;chunks.push(value);bytes+=value.length;show('Склад '+(i+1)+' · загружено '+(bytes/1048576).toFixed(1)+' МБ…')}
        buffer=await new Blob(chunks).arrayBuffer();
        if(new Uint8Array(buffer)[0]!==80)throw Error();
      }catch{show('Браузер не смог скачать склад. Откройте XLSX по ссылке, выберите скачанный файл и уберите ссылку из поля. ');const a=document.createElement('a');a.href=link;a.textContent='Скачать расстановку';status.append(a);return}
      show('Ищу коробки выбранного цвета · '+source.label+'…');
      await collect(buffer,source);
    }
    for(const source of files){show('Читаю '+source.label+'…');await collect(await source.file.arrayBuffer(),source)}
    if(!result.length)fail('На выбранных листах нет коробок с указанной заливкой.\n'+counts.join('\n'));
    showReady=false;
    boxes=result.map(b=>{
      const items=decode(b.raw), incomplete=items.some(x=>!x.size||!x.qty);
      return {...b,items,incomplete,manual:Boolean(b.comment)||incomplete};
    });$('assemblyBoxCodesPanel').hidden=false;review();show('Найдено коробок: '+boxes.length+'\n'+counts.join('\n'));
  }catch(err){show(err.message)}finally{button.disabled=false;$('layoutLoading').hidden=true}
});
function review(){reviewContents();
on('assembly',()=>writeBook(assemblyBoxCodes?'Сборочный со ШК.xlsx':'Сборочный.xlsx',assemblyRows(),true));on('direct',()=>templateFile(flat()));on('share',async()=>{const rows=assemblyRows();const bytes=new Uint8Array(await new Response(new Blob([JSON.stringify(rows)]).stream().pipeThrough(new CompressionStream('gzip'))).arrayBuffer());const hash=btoa(Array.from(bytes,b=>String.fromCharCode(b)).join(''));const u=location.href.split('#')[0]+'#assembly='+hash;$('linkout').innerHTML='<label>Ссылка с копией задания для сборки<input readonly value="'+esc(u)+'"></label>';$('linkout').querySelector('input').select();try{await navigator.clipboard.writeText(u);msg('Ссылка скопирована')}catch{msg('Скопируйте выделенную ссылку')}})}
function flat(){const index=boxes.findIndex(b=>b.manual);if(index>=0)fail("Подтвердите содержимое коробки "+(index+1));return boxes.flatMap((b,i)=>b.items.map(x=>{const c=catalogIndex.barcodes.get(x.barcode);if(!c||!Number.isSafeInteger(x.qty)||x.qty<1)fail('Проверьте баркод и количество в коробке '+(i+1));return {...c,qty:x.qty,p:b.p,no:i+1,shk:''}}))}
const assemblyHeader=['Баркод товара','Кол-во товаров','ШК короба','Цвет','Размер','Изделие'];
function assemblyRows(){
  const sections=new Map();
  for(const item of flat()){
    const pallet=String(item.p).trim().replace(/^№\s*\?$/, '?');
    if(!sections.has(pallet))sections.set(pallet,[]);
    const code=assemblyBoxCodes?.get(String(item.no))||'';
    sections.get(pallet).push([item.barcode,item.qty,code,item.color,item.size,item.product,item.no]);
  }
  return [assemblyHeader,...Array.from(sections).flatMap(([pallet,rows])=>[[pallet],...rows])];
}
function readAssembly(rows){let p='';const out=[];for(const r of rows.slice(1)){if(!r.some(Boolean))continue;if(!r[1]&&r[0]){p=r[0];continue}const [barcode,qty,shk,color,size,product,no]=r;if(!barcode||!Number.isSafeInteger(+qty)||+qty<1)fail('Некорректная строка для баркода '+barcode);out.push({barcode:String(barcode),qty:+qty,shk:String(shk||'').trim(),p:String(p),color,size,product:product||catalogIndex.barcodes.get(String(barcode))?.product||'',no})}if(!out.length)fail('Нет товарных строк');return out}
let shared=null;
async function imported(){if(!$('assembled').files.length&&!shared&&assemblyBoxCodes)return readAssembly(assemblyRows());return readImportedSource()}
async function readImportedSource(){if(shared&&!$('assembled').files.length)return readAssembly(shared);const w=await fileBook($('assembled'));const sheet=w.worksheets[0],rows=grid(sheet);let i=0;sheet.eachRow(row=>{const values=rows[i++],note=row.getCell(1).note;const box=typeof note==='string'?note.match(/^WB-box:(\d+)$/)?.[1]:undefined;if(box&&!values[6])values[6]=box});if(!/баркод/i.test(rows[0]?.[0]))fail('Первая колонка должна содержать баркод');return readAssembly(rows)}
function checked(rows){const byBox=new Map(),byCode=new Map();for(const x of rows){const key=x.no?JSON.stringify([x.p,String(x.no)]):null;if(key&&x.shk){if(byBox.has(key)&&byBox.get(key)!==x.shk)fail('В коробке '+x.no+' разные ШК');byBox.set(key,x.shk)}}for(const x of rows){const key=x.no?JSON.stringify([x.p,String(x.no)]):null;x.shk=x.shk||(key&&byBox.get(key));if(!x.shk)fail('Заполните ШК короба для '+x.barcode);const owner=key||x.p;if(byCode.has(x.shk)&&byCode.get(x.shk)!==owner)fail('ШК '+x.shk+' назначен разным коробкам');byCode.set(x.shk,owner)}return rows}
async function templateFile(items){const m=new Map();for(const x of items)m.set(x.barcode,(m.get(x.barcode)||0)+x.qty);await writeBook('template.xlsx',[['Баркод','Количество'],...m])}
on('template',async()=>templateFile(await imported()));on('final',async()=>writeBook('shk-excel.xlsx',[['Баркод товара','Кол-во товаров','ШК короба'],...checked(await imported()).map(x=>[x.barcode,x.qty,x.shk])]));
on('cards',async()=>{const rows=checked(await imported()),pallets=new Map();for(const x of rows){if(!x.p)fail('В файле нет номеров палет');if(!pallets.has(x.p))pallets.set(x.p,new Map());const p=pallets.get(x.p);if(!p.has(x.shk))p.set(x.shk,[]);p.get(x.shk).push(x)}const sum=a=>a.reduce((n,x)=>n+x.qty,0);const html=[...pallets].map(([p,groups])=>'<details open><summary>Паллета № '+esc(p)+' <small>Коробов: '+groups.size+' · Товаров: '+sum([...groups.values()].flat())+' шт.</small></summary>'+[...groups].map(([code,a])=>'<article><input type="checkbox"><div><b>'+esc(code)+'</b> · '+(new Set(a.map(x=>x.barcode)).size===1?'МОНО':'МИКС')+'<br>Итого: '+sum(a)+' шт.</div><ul>'+a.map(x=>'<li>'+esc([x.product||x.barcode,x.color,'размер '+x.size].join(', '))+' — '+x.qty+' шт.</li>').join('')+'</ul></article>').join('')+'</details>').join('');download('Задание по палетам.html','<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>Задание по палетам</title><style>body{font:16px Arial;background:#f5f7fa;margin:3vw}details{background:white;border-left:5px solid #2673c5;border-radius:8px;margin:16px 0;overflow:hidden}summary{padding:18px;background:#dceaff;font-weight:bold;cursor:pointer}small{float:right;color:#205b49}article{display:flex;gap:20px;align-items:center;padding:14px;border-top:1px solid #ddd}article:has(:checked){opacity:.5}ul{margin:0}</style><h1>Задание по палетам</h1>'+html,'text/html;charset=utf-8')});
on('packing',e=>{const f=e.target.elements,count=+f.count.value,warehouse=f.warehouse.value.trim(),supply=f.supply.value.trim(),date=f.deliveryDate.value;if(!warehouse||!supply||!date||count<1||count>999)fail('Проверьте поля');const deliveryDate=date.split('-').reverse().join('-');const win=window.open('','_blank');if(!win)fail('Разрешите открытие окна печати');win.document.write('<!doctype html><meta charset=utf-8><title>Упаковочные листы</title><style>@page{size:A4;margin:10mm 12mm}body{margin:0;font:18pt Arial}article{break-after:page}article:last-child{break-after:auto}h1{font-size:22pt;margin:0 0 3mm}table{border-collapse:collapse;width:100%;table-layout:fixed}td{border:1px solid;padding:3.5mm 4mm}td+td{font-weight:bold}</style>'+Array.from({length:count},(_,i)=>'<article><h1>Упаковочный лист</h1><table>'+[['Номер палеты',i+1],['Кол-во паллет в поставке',count],['Номер поставки',supply],['Кол-во коробов на данной паллете',16],['Склад назначения',warehouse],['Тип поставки','Короб'],['Наименование поставщика, ИНН','ИП Гладкая А. В.'],['Дата поставки',deliveryDate]].map(([a,b])=>'<tr><td>'+a+'</td><td>'+esc(b)+'</td></tr>').join('')+'</table></article>').join(''));win.document.close();win.focus();win.print();save({...db,settings:{warehouse}})},'submit');$('packing').elements.warehouse.value=db.settings.warehouse||'';
if(location.hash.startsWith('#assembly='))(async()=>{const bytes=Uint8Array.from(atob(location.hash.slice(10)),c=>c.charCodeAt(0));const rows=JSON.parse(await new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'))).text());readAssembly(rows);shared=rows;tab('Файлы для WB');const el=document.createElement('div');el.innerHTML='<h2>Сборочная таблица</h2><button id="getshared">Скачать сборочный XLSX</button><table>'+rows.map(r=>'<tr>'+r.slice(0,6).map(v=>'<td>'+esc(v)+'</td>').join('')+'</tr>').join('')+'</table>';$('sharedSource').append(el);on('getshared',()=>writeBook('Сборочный.xlsx',rows,true))})().catch(()=>msg('Некорректная ссылка на сборку'));

