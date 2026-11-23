function createCatalogIndex(items) {
  const variants = new Map(), barcodes = new Map();
  // A suffix П denotes a separate stock item. Only its separator varies.
  function key(product, color, size) {
    return JSON.stringify([norm(product), colorKey(color), norm(size)]);
  }
  for (const item of items) {
    const k = key(item.product, item.color, item.size);
    if (!variants.has(k)) variants.set(k, item.barcode);
    else if (variants.get(k) !== item.barcode) variants.set(k, '');
    barcodes.set(item.barcode, item);
  }
  return {barcodes, lookup:(product,color,size)=>variants.get(key(product,color,size)) || ''};
}
function showDirectory() {
  const search = $('dirSearch'), view = $('dirView');
  let page = 0;
  function draw() {
    const aliases = view.value === 'aliases';
    const records = aliases ? Object.entries(db.aliases).sort(([a],[b])=>a.localeCompare(b,'ru')) : db.catalog.map(x=>[x.product,x.color,x.size,x.barcode]);
    const query = norm(search.value);
    const filtered = records.filter(row=>row.some(x=>norm(x).includes(query)));
    page = Math.max(0,Math.min(page,Math.ceil(filtered.length/20)-1));
    const titles = aliases ? ['Аббревиатура','Изделие',''] : ['Изделие','Цвет','Размер','Баркод'];
    $('dirTable').innerHTML='<thead><tr>'+titles.map(x=>'<th>'+x+'</th>').join('')+'</tr></thead><tbody>'+filtered.slice(page*20,page*20+20).map(row=>'<tr>'+row.map(x=>'<td>'+esc(x)+'</td>').join('')+(aliases?'<td><button type="button" class="minor" data-edit="'+esc(row[0])+'">Изменить</button></td>':'')+'</tr>').join('')+'</tbody>';
    $('dirCount').textContent=filtered.length?(page*20+1)+'–'+Math.min(page*20+20,filtered.length)+' из '+filtered.length:'Совпадений нет';
    $('prevPage').disabled=page===0; $('nextPage').disabled=(page+1)*20>=filtered.length;
  }
  search.oninput=view.onchange=()=>{page=0;draw()};
  $('prevPage').onclick=()=>{page--;draw()}; $('nextPage').onclick=()=>{page++;draw()};
  $('dirTable').onclick=e=>{const b=e.target.closest('[data-edit]');if(b){const f=$('alias').elements;f.abbr.value=b.dataset.edit;f.product.value=db.aliases[b.dataset.edit];f.product.focus()}};
  draw();
}

