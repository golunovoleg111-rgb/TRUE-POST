const parseXml = text => new DOMParser().parseFromString(text,'application/xml');
async function readLayout(buffer, label) {
  const zip=await JSZip.loadAsync(buffer);
  async function xml(path){const f=zip.file(path); if(!f)throw Error('В XLSX отсутствует '+path);return parseXml(await f.async('string'))}
  const book=await xml('xl/workbook.xml'), rels=await xml('xl/_rels/workbook.xml.rels');
  const sheets=Array.from(book.getElementsByTagNameNS('*','sheet'));
  let sheet=sheets.find(s=>/^расстановка(?:\s+1)?$/i.test(s.getAttribute('name')))||sheets[0];
  if(sheets.length>1){const n=prompt('Лист расстановки для '+label+'\n'+sheets.map(s=>s.getAttribute('name')).join('\n'),sheet.getAttribute('name'));sheet=sheets.find(s=>s.getAttribute('name')===n);if(!sheet)throw Error('Лист не выбран')}
  const target=Array.from(rels.documentElement.children).find(r=>r.getAttribute('Id')===sheet.getAttribute('r:id')).getAttribute('Target');
  const path=new URL(target,'https://xlsx.local/xl/workbook.xml').pathname.slice(1);
  const styles=await xml('xl/styles.xml');
  const children=name=>Array.from(styles.getElementsByTagNameNS('*',name)[0].children);
  const fills=children('fills').map(f=>{const color=f.getElementsByTagNameNS('*','fgColor')[0];return color?.getAttribute('rgb')?.slice(-6).toUpperCase()==='FFFF00'||['5','13'].includes(color?.getAttribute('indexed'))});
  const yellow=children('cellXfs').map(f=>fills[Number(f.getAttribute('fillId'))]);
  const strings=zip.file('xl/sharedStrings.xml')?Array.from((await xml('xl/sharedStrings.xml')).getElementsByTagNameNS('*','si'),s=>Array.from(s.getElementsByTagNameNS('*','t'),t=>t.textContent).join('')):[];
  const raw=await zip.file(path).async('string'), headers=[], boxes=[];
  const attr=(s,k)=>s.match(new RegExp('\\s'+k+'="([^"]*)"'))?.[1];
  for(const m of raw.matchAll(/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)){
    if(!m[2])continue;
    const a=attr(m[1],'r'), type=attr(m[1],'t'),v=m[2].match(/<(?:\w+:)?v>([^<]*)</)?.[1];
    const value=String(type==='s'?strings[+v]||'':type==='inlineStr'?parseXml('<root>'+m[2]+'</root>').documentElement.textContent:v||'').trim();
    if(!value||!a)continue;
    const i=a.match(/^([A-Z]+)(\d+)$/), col=[...i[1]].reduce((n,c)=>26*n+c.charCodeAt(0)-64,0), row=+i[2];
    const h=value.match(/^(?:пал+ет[аы]?\s*)?№\s*(\d+)(?:\s*(?:[-–].*|\([^)]*\)))?$/i);
    if(h||/^(?:б\/н(?:\s|$)|№?\s*\?$)/i.test(value)){headers.push({row,col,n:h?.[1]||value});continue}
    if(yellow[+(attr(m[1],'s')||0)])boxes.push({row,col,a,value});
  }
  return boxes.map(({row,col,a,value})=>{const h=headers.filter(x=>x.row<row&&x.col<=col&&col<x.col+4).at(-1);return {p:h?.n||'Уточнить ('+a+')',id:label+'!'+a,raw:value}})
}
