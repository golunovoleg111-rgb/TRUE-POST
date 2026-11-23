const paths={box:'M3 7l9-4 9 4-9 4z M3 7v10l9 4 9-4V7 M12 11v10',file:'M6 3h8l4 4v14H6z M14 3v5h4 M9 13h6 M9 17h6',print:'M7 8V3h10v5 M6 17H3V9h18v8h-3 M7 14h10v7H7z',book:'M4 4h6l2 2 2-2h6v15h-6l-2 2-2-2H4z M12 6v15',upload:'M12 16V3 M7 8l5-5 5 5 M4 15v6h16v-6',arrow:'M8 4l8 8-8 8'};
function icon(name){return '<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round" stroke-linecap="round"><path d="'+paths[name]+'"/></svg>'}
document.querySelectorAll('nav button').forEach((b,i)=>{b.title=b.textContent;b.insertAdjacentHTML('afterbegin',icon(['box','file','print','book'][i]))});
document.querySelectorAll('input[type=file]').forEach(input=>{
 const wrapper=document.createElement('span');wrapper.className='file-control';input.before(wrapper);wrapper.append(input);
 const caption=document.createElement('span');caption.className='file-caption';caption.innerHTML=icon('upload')+'<strong>Выбрать '+(input.multiple?'файлы':'файл')+'</strong><span></span>';wrapper.append(caption);
 input.addEventListener('change',()=>caption.lastElementChild.textContent=Array.from(input.files,f=>f.name).join(', '));
});


document.querySelectorAll('[data-icon]').forEach(el=>{el.innerHTML=icon(el.dataset.icon)});
