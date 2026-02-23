// js/ui.js
export const $ = (sel)=>document.querySelector(sel);
// Element helper with children support.
// Usage: el('div', {className:'row'}, child1, child2, 'text')
export const el = (tag, props={}, ...children)=>{
  const e = document.createElement(tag);
  if(props && typeof props === 'object') Object.assign(e, props);
  for(const ch of children){
    if(ch===null || ch===undefined) continue;
    if(typeof ch === 'string' || typeof ch === 'number') e.appendChild(document.createTextNode(String(ch)));
    else e.appendChild(ch);
  }
  return e;
};
export function showMenu(anchorBtn, menuEl){
  const rect = anchorBtn.getBoundingClientRect();
  menuEl.style.display="block";
  menuEl.style.top = (rect.bottom + 6) + "px";
  menuEl.style.left = (rect.left) + "px";
}
export function hideMenu(menuEl){ menuEl.style.display="none"; }
export function downloadBlob(blob, filename){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download=filename;
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
}
export function rand1to9(){ return Math.floor(Math.random()*9)+1; }
export async function confirmDeleteNumber(title){
  const n=rand1to9();
  const v=prompt(`${title}\n\nПідтвердження видалення.\nВведіть число: ${n}`);
  if(v===null) return false;
  return String(v).trim()===String(n);
}
export function modalOpen({title, bodyNodes=[], actions=[]}){
  const backdrop = $("#modalBackdrop");
  const t = $("#modalTitle");
  const b = $("#modalBody");
  const a = $("#modalActions");
  t.textContent = title;
  b.innerHTML="";
  a.innerHTML="";
  for(const n of bodyNodes) b.appendChild(n);
  for(const act of actions) a.appendChild(act);
  backdrop.style.display="flex";
  return new Promise((resolve)=>{
    const close=(value)=>{
      backdrop.style.display="none";
      backdrop.onclick=null;
      resolve(value);
    };
    backdrop.onclick=(e)=>{ if(e.target===backdrop) close({type:"cancel"}); };
    actions.forEach(btn=>{
      const val = btn.dataset.value;
      btn.addEventListener("click", ()=>close({type:val}));
    });
  });
}
export function btn(text, value, cls="btn"){
  const b=document.createElement("button");
  b.className=cls;
  b.textContent=text;
  b.dataset.value=value;
  return b;
}
