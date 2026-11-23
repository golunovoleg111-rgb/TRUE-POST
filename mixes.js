const colorAliases = new Map([
  ['т/з','темно-зеленый'], ['т/с','темно-синий'], ['с/с','светло-серый'],
  ['с/б','светло-бежевый'], ['т/б','темно-бежевый'],
  ['бел','белый'], ['крас','красный'], ['мол','молочный'],
  ['роз','розовый'], ['чер','черный'], ['черное','черный'],
  ['синие','синий'], ['голубые','голубой'],
  ['фисташка','фисташковый'],
  ['коринчевый','коричневый'],
]);

const warehouseProducts = {
  'квж№1':'21_К_Вельвет',
  'квж№2':'29_К_Вельвет№2',
  'кзж№1':'216_К_Замша_№1',
  'кбж№1':'211_КБрюки_жилет',
  'твид№2':'227_К_Твид_№2',
  'влк':'24_К_Велюр',
  'квж_пю':'218_К_Вельвет_ПЮ',
  'б№5':'1_Б_№5_Теплые',
  'спорт':'228_К_спорт',
};
const productAliasKey=value=>norm(value).replace(/\s*№\s*/g,'№');
function mappedProduct(product) {
  const key=productAliasKey(product);
  return catalogIndex.aliases.get(key) || product;
}

function colorKey(value) {
  let color = norm(value).replace(/[‐‑–—]/g,'-');
  const stripe = /(?:[\s_*-]+[пp]|[\s_-]+в[\s_-]+пол(?:ос|со)ку)$/u;
  const striped = stripe.test(color);
  color = color.replace(stripe,'').replace(/\s*\/\s*/g,'/').replace(/[\s_-]+/g,'-');
  color = colorAliases.get(color) || color;
  return color + (striped ? '_п' : '');
}

function decode(raw) {
  let product='', color='', pending=false;
  const items=[];
  const add=(size='',qty=0)=>items.push({product,color,size,qty,barcode:size?catalogIndex.lookup(mappedProduct(product),color,size):''});
  const setColor=value=>{if(pending)add();color=value;pending=true};
  const quantityPattern=/^(\d+(?:\s*[-–/]\s*\d+)?)\s*[рp]?\s*[-–—]\s*(\d+)\s*(?:шт\.?)?$/i;
  function consume(line){
    const quantity=line.match(quantityPattern);
    if(quantity){add(quantity[1],Number(quantity[2]));pending=false;return}
    const inline=line.match(/^(.+?)\s+(\d+\s*[рp]?\s*[-–—]\s*\d+\s*(?:шт\.?)?)$/i);
    if(inline){setColor(inline[1]);consume(inline[2]);return}
    setColor(line);
  }
  for(const line of raw.split(/[\r\n;]+/).map(s=>s.trim().replace(/\s*\(микс[^)]*\)/ig,'').trim()).filter(Boolean)){
    if(/^микс(?:\s+с\s+комментарием)?$/i.test(line))continue;
    const words=line.split(/\s+/);
    let head='';
    for(let count=words.length;count>0;count--){
      const key=productAliasKey(words.slice(0,count).join(' '));
      if(catalogIndex.names.has(key)){head=words.splice(0,count).join(' ');break}
    }
    if(!product&&!head&&words.length>1&&/^[а-я]+(?:№\d+|\d+)$/i.test(words[0]))head=words.shift();
    if(head){
      if(pending)add();product=head;color='';pending=false;
      if(words.length)consume(words.join(' '));
      continue;
    }
    if(!product){product=line;continue}
    consume(line);
  }
  if(pending||!items.length)add();
  return items;
}

function validItem(item){return catalogIndex.barcodes.has(item.barcode)&&Number.isSafeInteger(item.qty)&&item.qty>0}
function needsCheck(box){return box.manual||!box.items.every(validItem)}
let showReady=false;
function reviewContents(){
  const incomplete=boxes.filter(needsCheck).length;
  const controls='<div class="actions"><button id="assembly">'+(assemblyBoxCodes?'Скачать сборочный со ШК':'Скачать сборочный XLSX')+'</button><button id="share">Создать ссылку на таблицу</button><button id="direct">Скачать Template</button></div><div id="linkout"></div>';
  const html=boxes.map((b,i)=>{
    if(!showReady&&!needsCheck(b))return '';
    const rows=b.items.map((x,j)=>{
      const product=mappedProduct(x.product);
      const choices=db.catalog.filter(v=>norm(v.product)===norm(product)&&colorKey(v.color)===colorKey(x.color)).sort((a,b)=>String(a.size).localeCompare(String(b.size),'ru',{numeric:true}));
      const attrs=' data-b="'+i+'" data-j="'+j+'" ';
      const selected=choices.some(v=>v.barcode===x.barcode);
      const picker=choices.length?'<select aria-label="Размер и баркод" '+attrs+'data-k="barcode"><option value="">Выберите размер и баркод</option>'+choices.map(v=>'<option value="'+esc(v.barcode)+'" '+(v.barcode===x.barcode?'selected':'')+'>'+esc(v.size)+' — '+esc(v.barcode)+'</option>').join('')+(x.barcode&&!selected?'<option selected>'+esc(x.barcode)+'</option>':'')+'</select>':'';
      return '<div class="mix-row"><strong>'+esc([x.product,x.color,x.size].filter(Boolean).join(' · '))+'</strong>'+picker+'<label>Баркод<input list="codes" '+attrs+'data-k="barcode" value="'+esc(x.barcode)+'"></label><label>Количество<input type="number" min="1" step="1" '+attrs+'data-k="qty" value="'+(x.qty||'')+'"></label>'+(b.items.length>1?'<button class="minor" data-remove="'+i+','+j+'">Убрать позицию</button>':'')+'</div>';
    }).join('');
    return '<details '+(needsCheck(b)?'open':'')+'><summary>Палета №'+esc(b.p)+' · коробка '+(i+1)+' · '+esc(b.raw.replace(/\n/g,' · '))+'</summary>'+(b.manual?'<p>'+(b.comment?'Указан микс с комментарием. ':b.incomplete?'В ячейке нет полного состава. Проверьте комментарий в расстановке. ':'')+'Укажите содержимое коробки вручную.</p>':'')+(b.comment?'<pre>'+esc(b.comment)+'</pre>':'')+rows+'<div class="actions"><button class="minor" data-add="'+i+'">Добавить позицию</button><button data-confirm="'+i+'">Подтвердить коробку</button></div></details>';
  }).join('');
  const host=$('review');
  host.innerHTML='<h2>Проверка сборки</h2><p>Всего коробок: '+boxes.length+' · Готово: '+(boxes.length-incomplete)+' · Требуют проверки: '+incomplete+'</p><button class="minor" id="showReady">'+(showReady?'Скрыть готовые':'Показать готовые')+'</button>'+controls+'<datalist id="codes"></datalist>'+html;
  $('showReady').onclick=()=>{showReady=!showReady;review()};
  host.onfocusin=e=>{if(e.target.matches('input[list="codes"]')&&!$('codes').children.length)$('codes').innerHTML=db.catalog.map(v=>'<option value="'+esc(v.barcode)+'">'+esc([v.product,v.color,v.size].join(' · '))+'</option>').join('')};
  host.oninput=e=>{const d=e.target.dataset;if(!d.k)return;const box=boxes[d.b],item=box.items[d.j];item[d.k]=d.k==='qty'?Number(e.target.value):e.target.value.trim();box.manual=true;
    if(d.k==='barcode'){const row=e.target.closest('.mix-row');row.querySelectorAll('[data-k="barcode"]').forEach(el=>{if(el!==e.target)el.value=item.barcode});const found=catalogIndex.barcodes.get(item.barcode);if(found)item.size=found.size}
  };
  host.onclick=e=>{const d=e.target.dataset;
    if(d.add!==undefined){const b=boxes[d.add],last=b.items.at(-1);b.items.push({...last,size:'',qty:0,barcode:''});b.manual=true;review()}
    if(d.remove!==undefined){const [i,j]=d.remove.split(',');boxes[i].items.splice(+j,1);boxes[i].manual=true;review()}
    if(d.confirm!==undefined){const b=boxes[d.confirm];if(!b.items.every(validItem)){msg('Проверьте баркоды и количество в коробке '+(+d.confirm+1));return}b.manual=false;msg('');review()}
  };
}
