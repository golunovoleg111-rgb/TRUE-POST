let assemblyBoxCodes=null;
let assemblyCodeRequest=0;
let pendingAssemblyBatch=null;

async function loadBoxCodes(file) {
  let rows;
  if (/\.xlsx$/i.test(file.name)) {
    const sheet=(await workbook(await file.arrayBuffer())).worksheets[0];
    if(!sheet)fail('В файле со ШК нет листов.');
    rows=[];
    const header=/^(?:шк|штрихкод(?:ы)?|баркод(?:ы)?)\s+короб(?:а|ов|ки|ок)$/i;
    let headerRow=0, codeColumn=0;
    sheet.eachRow(row=>{
      if(headerRow)return;
      row.eachCell(cell=>{
        if(!codeColumn && header.test(norm(cell.text))){
          headerRow=row.number;
          codeColumn=cell.col;
        }
      });
    });
    function codeText(cell){
      const value=cell.value;
      if(typeof value==='number'&&!Number.isSafeInteger(value))fail('Длинные ШК в Excel нужно хранить как текст.');
      const width=/^0+$/.test(cell.numFmt||'')?cell.numFmt.length:0;
      return typeof value==='number'&&width?String(value).padStart(width,'0'):cell.text.trim();
    }
    sheet.eachRow(row=>{
      if(codeColumn){
        if(row.number>headerRow)rows.push([codeText(row.getCell(codeColumn))]);
        return;
      }
      const values=[];
      row.eachCell(cell=>values.push(codeText(cell)));
      rows.push(values);
    });
  } else {
    rows=(await file.text()).split(/\r?\n/).map(line=>[line.trim().replace(/^"(\d+)"$/,'$1')]);
  }
  const codes=[];
  const seen=new Set();
  let hasHeader=false;
  for(const row of rows) {
    const values=row.filter(Boolean);
    if (!values.length)continue;
    const [code]=values;
    if(!hasHeader&&!codes.length&&values.length===1&&/^(?:шк(?:\s+короб(?:а|ов)?)?|штрихкод(?:ы)?(?:\s+короб(?:а|ов)?)?|баркод(?:ы)?(?:\s+короб(?:а|ов)?)?)$/i.test(code)) {
      hasHeader=true;
      continue;
    }
    if(values.length!==1||!/^\d+$/.test(code))fail('Укажите один числовой ШК в каждой строке списка.');
    if(seen.has(code))fail('ШК '+code+' повторяется в списке.');
    seen.add(code);
    codes.push(code);
  }
  if(!codes.length)fail('В списке нет ШК коробов.');
  return codes;
}

function placeBoxCodes(rows,codes) {
  const groups=new Map();
  const keyOf=row=>JSON.stringify([row.p,String(row.no)]);
  for(const row of rows) {
    if(!row.no)fail('Для авторасстановки нужен сборочный XLSX, скачанный с сайта. В нём сохранён состав каждой коробки.');
    const key=keyOf(row);
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(row);
  }
  if(groups.size!==codes.length)fail('В сборке коробок: '+groups.size+'. В списке ШК: '+codes.length+'. Количество должно совпадать.');
  const byBox=new Map();
  let mixCount=0;
  for(const [key,group] of groups) {
    const code=codes[byBox.size];
    if(group.some(row=>row.shk&&row.shk!==code))fail('У коробки на палете '+group[0].p+' уже указан другой ШК. Возьмите исходный сборочный файл или проверьте порядок списка.');
    byBox.set(key,code);
    if(group.length>1)mixCount++;
  }
  const placed=rows.map(row=>({...row,shk:byBox.get(keyOf(row))}));
  return {placed,mixCount};
}

function resetAssemblyBoxCodes() {
  assemblyCodeRequest++;
  assemblyBoxCodes=null;
  clearAssemblyBatch();
  const input=$('assemblyBoxCodes');
  input.value='';
  const filename=input.parentElement.querySelector('.file-caption span');
  if(filename)filename.textContent='';
  $('assemblyBoxCodesStatus').textContent='';
  const downloadButton=$('assembly');
  if(downloadButton)downloadButton.textContent='Скачать сборочный XLSX';
}

function clearAssemblyBatch() {
  pendingAssemblyBatch=null;
  $('assemblyBatch').hidden=true;
  $('assemblyBatchStart').value='';
  $('assemblyBatchPreview').textContent='';
}

function applyAssemblyCodes(rows,codes,previous=new Map()) {
  const {placed,mixCount}=placeBoxCodes(rows,codes);
  const next=new Map(previous);
  const owners=new Map([...next].map(([box,code])=>[code,box]));
  for(const row of placed){
    const box=String(row.no), existing=next.get(box), owner=owners.get(row.shk);
    if(existing&&existing!==row.shk)fail('У коробки №'+box+' уже есть другой ШК. Проверьте начальный номер.');
    if(owner&&owner!==box)fail('ШК '+row.shk+' уже назначен коробке №'+owner+'.');
    owners.set(row.shk,box);
    next.set(box,row.shk);
  }
  assemblyBoxCodes=next;
  $('assembly').textContent='Скачать сборочный со ШК';
  $('assemblyBoxCodesStatus').textContent='Заполнено коробок: '+next.size+' из '+boxes.length+'. '+(next.size===boxes.length?'Все ШК проставлены.':'Осталось заполнить: '+(boxes.length-next.size)+'.')+' В этом списке миксов: '+mixCount+'.';
  clearAssemblyBatch();
  saveDraft();
}

function selectedBatchRange() {
  if(!pendingAssemblyBatch)return null;
  const start=Number($('assemblyBatchStart').value);
  const end=start+pendingAssemblyBatch.codes.length-1;
  if(!Number.isInteger(start)||start<1||end>pendingAssemblyBatch.total)return null;
  return {start,end};
}

$('assemblyBatchStart').addEventListener('input',()=>{
  const range=selectedBatchRange();
  $('assemblyBatchApply').disabled=!range;
  $('assemblyBatchPreview').textContent=range?'Коробки №'+range.start+'–'+range.end+' по порядку сборки. Палеты: '+boxes[range.start-1].p+' → '+boxes[range.end-1].p+'.':'Укажите начальный номер коробки в допустимом диапазоне.';
});
$('assemblyBatch').addEventListener('submit',e=>{
  e.preventDefault();
  const range=selectedBatchRange();
  if(!range)return;
  const {rows,codes}=pendingAssemblyBatch;
  const batchRows=rows.filter(row=>Number(row.no)>=range.start&&Number(row.no)<=range.end);
  try {applyAssemblyCodes(batchRows,codes,assemblyBoxCodes||new Map())}
  catch(err){$('assemblyBoxCodesStatus').textContent=err.message}
});

$('assemblyBoxCodes').addEventListener('change',async e=>{
  const file=e.target.files[0];
  if(!file)return;
  const request=++assemblyCodeRequest;
  const status=$('assemblyBoxCodesStatus');
  clearAssemblyBatch();
  status.textContent='Расставляю ШК…';
  try {
    // Box numbering follows the same pallet order as the export.
    const rows=flat();
    const codes=await loadBoxCodes(file);
    if(request!==assemblyCodeRequest)return;
    const total=boxes.length;
    if(codes.length>total)fail('В сборке '+total+' коробок, а в списке '+codes.length+' ШК. Уберите лишние коды.');
    if(codes.length===total){applyAssemblyCodes(rows,codes);return}
    pendingAssemblyBatch={rows,codes,total};
    $('assemblyBatchStart').max=total-codes.length+1;
    $('assemblyBatchApply').disabled=true;
    $('assemblyBatch').hidden=false;
    status.textContent='В сборке '+total+' коробок, в списке '+codes.length+' ШК. Выберите, с какой коробки начать.';
  }catch(err){
    if(request!==assemblyCodeRequest)return;
    status.textContent=err.message+(assemblyBoxCodes?' Сохранена предыдущая расстановка ШК.':'');
    e.target.value='';
  }
});

restoreDraft();
