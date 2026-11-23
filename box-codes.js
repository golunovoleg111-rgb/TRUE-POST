let assignedBoxRows=null;
let boxCodeRequest=0;
let assemblyBoxCodes=null;
let assemblyCodeRequest=0;

function resetBoxCodes() {
  boxCodeRequest++;
  assignedBoxRows=null;
  $('boxCodes').value='';
  const filename=$('boxCodes').parentElement.querySelector('.file-caption span');
  if(filename)filename.textContent='';
  $('boxCodesStatus').textContent='';
  $('downloadAssigned').hidden=true;
}

async function loadBoxCodes(file) {
  let rows;
  if (/\.xlsx$/i.test(file.name)) {
    const sheet=(await workbook(await file.arrayBuffer())).worksheets[0];
    if(!sheet)fail('В файле со ШК нет листов.');
    rows=[];
    sheet.eachRow(row=>{
      const values=[];
      row.eachCell(cell=>{
        const value=cell.value;
        if(typeof value==='number'&&!Number.isSafeInteger(value))fail('Длинные ШК в Excel нужно хранить как текст.');
        const width=/^0+$/.test(cell.numFmt||'')?cell.numFmt.length:0;
        values.push(typeof value==='number'&&width?String(value).padStart(width,'0'):cell.text.trim());
      });
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

$('assembled').addEventListener('change',resetBoxCodes);
$('review').addEventListener('input',e=>{
  if(!e.target.dataset.k)return;
  if(!$('assembled').files.length&&!shared)resetBoxCodes();
  resetAssemblyBoxCodes();
});
$('review').addEventListener('click',e=>{
  if(!e.target.matches('[data-add],[data-remove]'))return;
  if(!$('assembled').files.length&&!shared)resetBoxCodes();
  resetAssemblyBoxCodes();
});
$('boxCodes').addEventListener('change',async e=>{
  const file=e.target.files[0];
  if(!file)return;
  const request=++boxCodeRequest;
  const status=$('boxCodesStatus');
  status.textContent='Расставляю ШК…';
  try {
    const fromCurrent=!$('assembled').files.length&&!shared&&boxes.length;
    const [rows,codes]=await Promise.all([fromCurrent?readAssembly(assemblyRows(false)):readImportedSource(),loadBoxCodes(file)]);
    const {placed,mixCount}=placeBoxCodes(rows,codes);
    if(request!==boxCodeRequest)return;
    assignedBoxRows=placed;
    $('downloadAssigned').hidden=false;
    status.textContent='Готово. Заполнено коробок: '+codes.length+', из них миксов: '+mixCount+'.';
  }catch(err){
    if(request!==boxCodeRequest)return;
    status.textContent=err.message+(assignedBoxRows?' Сохранена предыдущая расстановка ШК.':'');
    $('boxCodes').value='';
  }
});

on('downloadAssigned',()=>{
  if(!assignedBoxRows)return;
  const rows=[assemblyHeader];
  let previous;
  for(const item of assignedBoxRows) {
    if(item.p!==previous){rows.push([item.p]);previous=item.p}
    rows.push([item.barcode,item.qty,item.shk,item.color,item.size,item.product,item.no]);
  }
  return writeBook('Сборочный со ШК.xlsx',rows,true);
});

function resetAssemblyBoxCodes() {
  assemblyCodeRequest++;
  assemblyBoxCodes=null;
  const input=$('assemblyBoxCodes');
  input.value='';
  const filename=input.parentElement.querySelector('.file-caption span');
  if(filename)filename.textContent='';
  $('assemblyBoxCodesStatus').textContent='';
  const downloadButton=$('assembly');
  if(downloadButton)downloadButton.textContent='Скачать сборочный XLSX';
}

$('assemblyBoxCodes').addEventListener('change',async e=>{
  const file=e.target.files[0];
  if(!file)return;
  const request=++assemblyCodeRequest;
  const status=$('assemblyBoxCodesStatus');
  status.textContent='Расставляю ШК…';
  try {
    const rows=readAssembly(assemblyRows(false));
    const codes=await loadBoxCodes(file);
    if(request!==assemblyCodeRequest)return;
    const {placed,mixCount}=placeBoxCodes(rows,codes);
    assemblyBoxCodes=new Map(placed.map(row=>[String(row.no),row.shk]));
    $('assembly').textContent='Скачать сборочный со ШК';
    status.textContent='Готово. Заполнено коробок: '+codes.length+', из них миксов: '+mixCount+'.';
  }catch(err){
    if(request!==assemblyCodeRequest)return;
    status.textContent=err.message+(assemblyBoxCodes?' Сохранена предыдущая расстановка ШК.':'');
    e.target.value='';
  }
});
