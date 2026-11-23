const fillOptions = [
  ['FFFF00','Жёлтый'], ['FF0000','Красный'], ['990000','Тёмно-красный'],
  ['00FF00','Зелёный'], ['0000FF','Синий'], ['00FFFF','Бирюзовый'],
  ['FF9900','Оранжевый'], ['9900FF','Фиолетовый'], ['FF00FF','Розовый'],
  ['CCCCCC','Серый'], ['custom','Другой цвет (HEX)'],
];
const sourceFills = new Map();

function layoutSources() {
  const urls=$('urls').value.split(/\n/).map(value=>value.trim()).filter(Boolean);
  const files=[...$('layouts').files];
  const sources=urls.map((url,i)=>({url,label:'Ссылка '+(i+1),detail:url,key:'url:'+url}));
  for(const file of files) sources.push({file,label:file.name,key:'file:'+JSON.stringify([file.name,file.size,file.lastModified])});
  const counts=new Map();
  return sources.map(source=>{
    const count=(counts.get(source.key)||0)+1;
    counts.set(source.key,count);
    return {...source,key:source.key+':'+count};
  });
}

function sourceFill(source) {
  const saved=sourceFills.get(source.key);
  return saved?.choice==='custom'?saved.custom.slice(1).toUpperCase():saved?.choice||'FFFF00';
}

function renderLayoutSources() {
  const host=$('layoutColors');
  host.replaceChildren();
  for(const source of layoutSources()){
    const row=document.createElement('div');
    row.className='layout-color-row';
    const title=document.createElement('div');
    title.className='layout-source-name';
    title.textContent=source.label;
    if(source.detail){const detail=document.createElement('small');detail.textContent=source.detail;title.append(detail)}
    const label=document.createElement('label');
    label.textContent='Цвет заливки';
    const select=document.createElement('select');
    select.setAttribute('aria-label','Цвет заливки · '+source.label);
    for(const [value,name] of fillOptions)select.add(new Option(name,value));
    const state=sourceFills.get(source.key)||{choice:'FFFF00',custom:'#ffff00'};
    select.value=state.choice;
    const picker=document.createElement('input');
    picker.type='color';
    picker.value=state.custom;
    picker.setAttribute('aria-label','Свой цвет · '+source.label);
    picker.title='Выберите цвет или введите HEX в палитре браузера';
    picker.hidden=select.value!=='custom';
    select.onchange=()=>{
      state.choice=select.value;
      picker.hidden=state.choice!=='custom';
      sourceFills.set(source.key,state);
    };
    picker.oninput=()=>{state.custom=picker.value;sourceFills.set(source.key,state)};
    label.append(select);
    row.append(title,label,picker);
    host.append(row);
  }
}

$('urls').addEventListener('input',renderLayoutSources);
$('layouts').addEventListener('change',renderLayoutSources);
renderLayoutSources();
