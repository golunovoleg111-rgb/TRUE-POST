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
    title.title=source.detail||source.label;
    const state=sourceFills.get(source.key)||{choice:'FFFF00',custom:'#ffff00'};
    const dropdown=document.createElement('details');
    dropdown.className='fill-dropdown';
    const summary=document.createElement('summary');
    const palette=document.createElement('div');
    palette.className='fill-palette';
    const makeSwatch=value=>{
      const swatch=document.createElement('span');
      swatch.className='fill-swatch';
      swatch.style.backgroundColor=value==='custom'?state.custom:'#'+value;
      return swatch;
    };
    const refresh=()=>{
      const name=fillOptions.find(([value])=>value===state.choice)[1];
      summary.replaceChildren(makeSwatch(state.choice),document.createTextNode(name));
      summary.setAttribute('aria-label','Цвет заливки · '+source.label+' · '+name);
      palette.querySelectorAll('button').forEach(btn=>btn.setAttribute('aria-pressed',String(btn.dataset.fill===state.choice)));
    };
    const picker=document.createElement('input');
    picker.type='color';
    picker.value=state.custom;
    picker.setAttribute('aria-label','Свой цвет · '+source.label);
    picker.title='Выберите цвет или введите HEX в палитре браузера';
    picker.hidden=state.choice!=='custom';
    for(const [value,name] of fillOptions){
      const item=document.createElement('button');
      item.type='button';
      item.dataset.fill=value;
      item.append(makeSwatch(value),document.createTextNode(name));
      item.onclick=()=>{
        state.choice=value;
        sourceFills.set(source.key,state);
        picker.hidden=value!=='custom';
        refresh();
        dropdown.open=false;
        summary.focus();
      };
      palette.append(item);
    }
    picker.oninput=()=>{state.custom=picker.value;sourceFills.set(source.key,state);refresh()};
    dropdown.addEventListener('keydown',e=>{
      if(e.key==='Escape'&&dropdown.open){dropdown.open=false;summary.focus()}
    });
    dropdown.addEventListener('focusout',e=>{if(e.relatedTarget&&!dropdown.contains(e.relatedTarget))dropdown.open=false});
    dropdown.append(summary,palette);
    refresh();
    row.append(title,dropdown,picker);
    host.append(row);
  }
}

document.addEventListener('click',e=>{
  document.querySelectorAll('.fill-dropdown[open]').forEach(menu=>{if(!menu.contains(e.target))menu.open=false});
});
$('urls').addEventListener('input',renderLayoutSources);
$('layouts').addEventListener('change',renderLayoutSources);
renderLayoutSources();
