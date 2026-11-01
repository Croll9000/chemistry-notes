/* Notes Packet & Slide Builder
   Reads builder/data.json (made from the website's pages each time the site is built) and makes
   a guided-notes packet (.docx), a teacher key (.docx) and slides (.pptx) in the browser. */
(function(){
"use strict";
/* the teacher code is checked against this SHA-256 hash (a deterrent, not real security) */
const PASS_SHA="22b098870a5404367538814101edc9dd95aaa108c8193699bda0dce58c3a1296";
const $=s=>document.querySelector(s), $$=(s,r=document)=>[...r.querySelectorAll(s)];
const C={teal:"40BAD2",tealDeep:"1F8FA8",tealInk:"0E5F72",purple:"7030A0",orange:"E07B00",ink:"1B2328",body:"4A5257",row:"E7F6F9",row2:"BBE6EF",box:"FFF8DC",boxLine:"E2C766",red:"C40000",key:"C40000"};
let DATA=null, GLOSS=[];
const sel=new Set();               /* "slug|secId" */
const MAN=new Set(), PRE=new Set(); /* token keys that are blanks */
const O={blank:"pre",lc:"page",ans:"key",intro:true,enr:true};
const IMGCACHE={};

/* ---------------- access ---------------- */
async function sha(t){ const b=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(t.trim().toLowerCase())); return [...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,"0")).join(""); }
async function unlock(code){ let h; try{ h=await sha(code); }catch(e){ $("#gerr").textContent="This browser could not check the code. Open the page from its https:// address (not a saved file) and try again."; return; } if(h!==PASS_SHA){ $("#gerr").textContent="That code is not right."; return; } try{ sessionStorage.setItem("pb-ok","1"); }catch(e){} start(); }
$("#go").addEventListener("click",()=>unlock($("#code").value)); $("#code").addEventListener("keydown",e=>{ if(e.key==="Enter") unlock($("#code").value); });
try{ if(sessionStorage.getItem("pb-ok")==="1") start(); }catch(e){}

/* ---------------- data ---------------- */
async function start(){ $("#gate").hidden=true; $("#app").hidden=false; status("Loading the notes…");
  try{ DATA=await (await fetch("builder/data.json",{cache:"no-cache"})).json(); }catch(e){ status("Could not load the notes data (builder/data.json).",true); return; }
  prepare(); tree(); wire(); render(); status(`Notes from ${DATA.units.length} units, updated ${new Date(DATA.built).toLocaleDateString()}.`); }

/* split styled runs into word tokens; every word gets a key so it can be blanked */
function tokenize(runs,ctx){ const out=[]; (runs||[]).forEach(r=>{ const st=Object.assign({},r); delete st.x; String(r.x).split(/(\s+)/).forEach(w=>{ if(w==="") return; const sp=/^\s+$/.test(w); out.push({w,st,sp,k:sp?null:ctx.slug+"|"+ctx.sec+"|"+(ctx.n++)}); }); }); return out; }
function prepare(){ const gl=new Set();
  Object.entries(DATA.pages).forEach(([slug,pg])=>{ pg.sections.forEach(sec=>{ const ctx={slug,sec:sec.id,n:0};
    const walk=bs=>bs.forEach(b=>{ if(b.t==="p"||b.t==="box") b.tk=tokenize(b.runs,ctx); else if(b.t==="list") b.items.forEach(it=>{ it.tk=tokenize(it.runs,ctx); walk(it.sub); }); else if(b.t==="table") b.rows.forEach(r=>r.forEach(c=>c.tk=tokenize(c.runs,ctx))); });
    walk(sec.blocks);
    /* glossary: every defined term on the site */
    const gw=bs=>bs.forEach(b=>{ if(b.t==="list") b.items.forEach(it=>{ let cur=[]; it.tk.forEach(t=>{ if(t.st.def) cur.push(t.w); else if(!t.sp&&cur.length){ gl.add(cur.join(" ").toLowerCase()); cur=[]; } }); if(cur.length) gl.add(cur.join(" ").toLowerCase()); gw(it.sub); }); }); gw(sec.blocks); }); });
  GLOSS=[...gl].map(g=>g.replace(/[^a-z0-9 -]/g,"").trim()).filter(g=>g.replace(/[^a-z]/g,"").length>=5).sort((a,b)=>b.length-a.length);
  Object.entries(DATA.pages).forEach(([slug,pg])=>{ const terms=new Set(); pg.sections.forEach(sec=>{ const w=bs=>bs.forEach(b=>{ if(b.t==="list") b.items.forEach(it=>{ let cur=[]; it.tk.forEach(t=>{ if(t.st.def) cur.push(t.w); else if(!t.sp&&cur.length){ terms.add(cur.join(" ")); cur=[]; } }); if(cur.length) terms.add(cur.join(" ")); w(it.sub); }); }); w(sec.blocks); });
    const pt=[...terms].map(g=>g.toLowerCase().replace(/[^a-z0-9 -]/g,"").trim()).filter(g=>g.replace(/[^a-z]/g,"").length>=5).sort((a,b)=>b.length-a.length);
    pg.sections.forEach(sec=>preChoose(sec,pt)); }); }
/* pre-chosen blanks: the defined terms first, then other terms defined on the same page (or multi-word terms from other pages), at most 4 blanks per section */
function preChoose(sec,pageTerms){ const groups=[], usedTerms=new Set(); const items=[];
  const walk=bs=>bs.forEach(b=>{ if(b.t==="list") b.items.forEach(it=>{ items.push(it.tk); walk(it.sub); }); else if(b.t==="p"&&b.style==="body") items.push(b.tk); }); walk(sec.blocks);
  const clean=w=>w.toLowerCase().replace(/^[^a-z0-9]+|[^a-z0-9]+$/g,"");
  items.forEach(tk=>{ let cur=[]; tk.forEach(t=>{ if(t.sp) return; if(t.st.def) cur.push(t); else if(cur.length){ groups.push(cur.map(x=>x.k)); usedTerms.add(cur.map(x=>clean(x.w)).join(" ")); cur=[]; } }); if(cur.length){ groups.push(cur.map(x=>x.k)); usedTerms.add(cur.map(x=>clean(x.w)).join(" ")); } });
  const cands=pageTerms.concat(GLOSS.filter(g=>g.includes(" "))).filter((g,i,a)=>a.indexOf(g)===i);
  if(groups.length<2){ for(const tk of items){ if(groups.length>=3) break; const words=tk.filter(t=>!t.sp&&!t.st.sub&&!t.st.sup); const low=words.map(t=>clean(t.w));
      for(const g of cands){ if(usedTerms.has(g)) continue; const gw=g.split(" "); let hit=false; for(let i=0;i+gw.length<=low.length;i++){ if(gw.every((w,j)=>low[i+j]===w)&&!words.slice(i,i+gw.length).some(t=>t.st.def||groups.some(G=>G.includes(t.k)))){ groups.push(words.slice(i,i+gw.length).map(t=>t.k)); usedTerms.add(g); hit=true; break; } } if(hit||groups.length>=3) break; } } }
  groups.slice(0,4).forEach(G=>G.forEach(k=>PRE.add(k))); }
const blanks=()=>O.blank==="pre"?PRE:O.blank==="man"?MAN:new Set();

/* ---------------- the checklist ---------------- */
function tree(){ const t=$("#tree"); t.innerHTML=DATA.units.map(u=>{ const pages=u.pages.filter(p=>DATA.pages[p.slug]&&DATA.pages[p.slug].sections.length); if(!pages.length) return "";
    return `<details class="u"><summary><input type="checkbox" data-u="${u.unit}"> ${u.name}</summary>${pages.map(p=>`<details class="pg"><summary><input type="checkbox" data-p="${p.slug}"> ${p.title}${p.enr?' <span class="tag">ENRICHMENT</span>':""}</summary>${DATA.pages[p.slug].sections.map(s=>`<label class="sec"><input type="checkbox" data-s="${p.slug}|${s.id}"> ${esc(s.title)}${s.enr?' <span class="tag">ENR</span>':""}</label>`).join("")}</details>`).join("")}</details>`; }).join("");
  t.addEventListener("click",e=>{ if(e.target.matches("summary input")) e.stopPropagation(); },true);
  t.addEventListener("change",e=>{ const c=e.target; if(c.dataset.u){ DATA.units.find(u=>u.unit==c.dataset.u).pages.forEach(p=>setPage(p.slug,c.checked)); }
    else if(c.dataset.p) setPage(c.dataset.p,c.checked); else if(c.dataset.s){ c.checked?sel.add(c.dataset.s):sel.delete(c.dataset.s); }
    syncTree(); render(); }); syncTree(); }
function setPage(slug,on){ const pg=DATA.pages[slug]; if(!pg) return; pg.sections.forEach(s=>{ const k=slug+"|"+s.id; if(on&&(O.enr||!s.enr)) sel.add(k); else if(!on) sel.delete(k); }); }
function syncTree(){ $$("#tree [data-s]").forEach(c=>c.checked=sel.has(c.dataset.s));
  $$("#tree [data-p]").forEach(c=>{ const ss=DATA.pages[c.dataset.p].sections.map(s=>c.dataset.p+"|"+s.id), n=ss.filter(k=>sel.has(k)).length; c.checked=n>0&&n===ss.length; c.indeterminate=n>0&&n<ss.length; });
  $$("#tree [data-u]").forEach(c=>{ const u=DATA.units.find(x=>x.unit==c.dataset.u); const ss=u.pages.flatMap(p=>(DATA.pages[p.slug]?DATA.pages[p.slug].sections:[]).map(s=>p.slug+"|"+s.id)), n=ss.filter(k=>sel.has(k)).length; c.checked=n>0&&n===ss.length; c.indeterminate=n>0&&n<ss.length; });
  const pages=new Set([...sel].map(k=>k.split("|")[0])); $("#count").textContent=sel.size? `${sel.size} section${sel.size>1?"s":""} from ${pages.size} page${pages.size>1?"s":""}` : "Nothing chosen yet.";
  if(!$("#title").dataset.touched) $("#title").value=autoTitle(); }
function autoTitle(){ const pages=order().map(x=>x.page.title); const u=[...new Set(order().map(x=>x.unit.short))]; if(!pages.length) return ""; return pages.length===1? pages[0] : u.length===1? `${u[0]}: ${[...new Set(pages)].join(", ")}`.slice(0,90) : "Chemistry Notes"; }
function wire(){ const seg=(id,key,after)=>$$(`#${id} button`).forEach(b=>b.addEventListener("click",()=>{ O[key]=b.dataset.v; $$(`#${id} button`).forEach(x=>x.setAttribute("aria-pressed",String(x===b))); after&&after(); render(); }));
  seg("o-blank","blank",()=>{ $("#seed").hidden=O.blank!=="man"; }); seg("o-lc","lc"); seg("o-ans","ans");
  $("#seed").addEventListener("click",()=>{ PRE.forEach(k=>MAN.add(k)); render(); });
  $("#intro").addEventListener("change",e=>{ O.intro=e.target.checked; render(); }); $("#enr").addEventListener("change",e=>{ O.enr=e.target.checked; });
  $("#title").addEventListener("input",e=>{ e.target.dataset.touched="1"; });
  $("#pv").addEventListener("click",e=>{ if(O.blank!=="man") return; const w=e.target.closest(".w"); if(!w||!w.dataset.k) return; MAN.has(w.dataset.k)?MAN.delete(w.dataset.k):MAN.add(w.dataset.k); w.classList.toggle("bl"); });
  $("#d-stu").addEventListener("click",()=>run(()=>makeDocx(false),"student packet"));
  $("#d-key").addEventListener("click",()=>run(()=>makeDocx(true),"teacher key"));
  $("#d-ppt").addEventListener("click",()=>run(makePptx,"slides")); }
async function run(fn,what){ if(!sel.size){ status("Check off some notes first.",true); return; } $$(".dl button").forEach(b=>b.disabled=true); status(`Making the ${what}…`);
  try{ await fn(); status(`✔ The ${what} downloaded.`); }catch(e){ console.error(e); status(`Something went wrong making the ${what}: ${e.message}`,true); }
  $$(".dl button").forEach(b=>b.disabled=false); }
function status(t,err){ const s=$("#status"); s.textContent=t; s.className="status"+(err?" err":""); }
const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"})[c]);

/* the chosen sections in site order, grouped by page */
function order(){ const out=[]; DATA.units.forEach(u=>u.pages.forEach(p=>{ const pg=DATA.pages[p.slug]; if(!pg) return; const secs=pg.sections.filter(s=>sel.has(p.slug+"|"+s.id)); if(secs.length) out.push({unit:u,page:p,pg,secs}); })); return out; }

/* ---------------- the document model shared by the preview, the packet and the slides ---------------- */
/* tokens → runs, with consecutive blanked words merged into one blank */
function toRuns(tk,B){ const out=[]; let i=0; while(i<tk.length){ const t=tk[i];
    if(!t.sp&&B.has(t.k)){ let j=i, words=[t.w], keys=[t.k]; while(j+2<tk.length&&tk[j+1].sp&&!tk[j+2].sp&&B.has(tk[j+2].k)){ words.push(tk[j+2].w); keys.push(tk[j+2].k); j+=2; }
      const m=words.join(" ").match(/^([(“‘"]*)(.*?)([.,;:!?)”’"]*)$/); if(m[1]) out.push({x:m[1],st:{}}); out.push({x:m[2],blank:true,st:t.st,keys}); if(m[3]) out.push({x:m[3],st:{}}); i=j+1; continue; }
    const L=out[out.length-1]; if(L&&!L.blank&&same(L.st,t.st)) L.x+=t.w; else out.push({x:t.w,st:t.st}); i++; }
  return out; }
const same=(a,b)=>["b","i","u","s","sub","sup","c","def"].every(k=>(a[k]||null)===(b[k]||null));
function model(){ const B=blanks(), M=[], lcs=[]; let listId=0;
  order().forEach(({unit,page,pg,secs})=>{ M.push({k:"page",text:page.title,unit:unit.name,enr:page.enr});
    secs.forEach(sec=>{ M.push({k:"sec",text:sec.title,enr:sec.enr,page:page.title});
      const walk=(bs,lvl,parentOrdered)=>bs.forEach(b=>{
        if(b.t==="p"){ if(b.style==="intro"&&!O.intro) return; M.push({k:"p",style:b.style,runs:toRuns(b.tk,B),tk:b.tk}); }
        else if(b.t==="box") M.push({k:"box",kind:b.kind,runs:toRuns(b.tk,B),tk:b.tk});
        else if(b.t==="table") M.push({k:"table",rows:b.rows.map(r=>r.map(c=>({runs:toRuns(c.tk,B),tk:c.tk,th:c.th,cs:c.cs,rs:c.rs})))});
        else if(b.t==="img") imgs(pg,b.snap).forEach(x=>M.push(x));
        else if(b.t==="list"){ const id=++listId; b.items.forEach(it=>{ M.push({k:"li",level:lvl,ordered:b.ordered,list:id,runs:toRuns(it.tk,B),tk:it.tk}); walk(it.sub,lvl+1,b.ordered); }); } });
      walk(sec.blocks,0,false); });
    if(pg.lc&&pg.lc.length&&O.lc!=="none"){ const L={k:"lc",page:page.title,items:pg.lc,pg}; if(O.lc==="page") M.push(L); else lcs.push(L); } });
  if(O.lc==="end"&&lcs.length){ M.push({k:"page",text:"Learning Check",unit:""}); lcs.forEach(L=>M.push(L)); }
  return M; }
function imgs(pg,id){ const s=pg.snaps[id]; if(!s||s.failed) return [];
  if(s.files) return s.files.map((f,i)=>({k:"img",file:f,w:s.sizes[i][0],h:s.sizes[i][1],label:s.labels&&s.labels[i]}));
  return [{k:"img",file:s.file,w:s.size[0],h:s.size[1]}]; }

/* ---------------- preview ---------------- */
function render(){ const pv=$("#pv"); $("#pvwrap").classList.toggle("manual",O.blank==="man");
  $("#hint").innerHTML= O.blank==="man"? "<b>Choose my own:</b> click any word in the notes to make it a blank (click again to undo). Blanks are shaded." : O.blank==="pre"? "<b>Pre-chosen blanks</b> are shaded: the defined terms and a few key words in each section. Switch to “Choose my own” to pick them yourself." : "No blanks: the packet has the full notes.";
  if(!sel.size){ pv.innerHTML=`<div class="empty">Check off some notes to see a preview.</div>`; return; }
  const B=blanks(); const W=tk=>(tk||[]).map(t=>{ const inner=styleHTML(esc(t.w),t.st); return t.sp? inner : `<span class="w${B.has(t.k)?" bl":""}" data-k="${t.k}">${inner}</span>`; }).join("");
  let h="", open=[]; const close=(to)=>{ while(open.length>to){ h+=`</${open.pop()}>`; } };
  model().forEach(m=>{ if(m.k!=="li") close(0);
    if(m.k==="page") h+=`<div class="pv-page">${esc(m.text)}${m.unit?` <span style="font-size:13px;color:#4a5257;font-weight:600">${esc(m.unit)}</span>`:""}</div>`;
    else if(m.k==="sec") h+=`<div class="pv-sec${m.enr?" enr":""}">${esc(m.text)}${m.enr?" (Enrichment)":""}</div>`;
    else if(m.k==="p") h+=`<p class="${m.style}">${W(m.tk)}</p>`;
    else if(m.k==="box") h+=`<div class="box">${W(m.tk)}</div>`;
    else if(m.k==="table") h+=`<table>${m.rows.map(r=>`<tr>${r.map(c=>`<${c.th?"th":"td"}${c.cs?` colspan="${c.cs}"`:""}${c.rs?` rowspan="${c.rs}"`:""}>${W(c.tk)}</${c.th?"th":"td"}>`).join("")}</tr>`).join("")}</table>`;
    else if(m.k==="img") h+=`<figure><img data-file="${m.file}" alt="">${m.label?`<figcaption>Example: ${esc(m.label)}</figcaption>`:""}</figure>`;
    else if(m.k==="li"){ const want=m.level+1; close(want); while(open.length<want){ const t=m.ordered&&open.length===m.level?"ol":"ul"; h+=`<${t}>`; open.push(t); } h+=`<li>${W(m.tk)}</li>`; }
    else if(m.k==="lc"){ h+=`<div class="pv-sec lc">Learning Check: ${esc(m.page)}</div>`; m.items.forEach(it=>{ h+= it.t==="cluster"? clusterHTML(it,m.pg) : qHTML(it,m.pg); }); } });
  close(0); pv.innerHTML=h; hydrate(); }
function styleHTML(x,st){ if(st.sub) x=`<sub>${x}</sub>`; if(st.sup) x=`<sup>${x}</sup>`; if(st.b||st.def) x=`<b>${x}</b>`; if(st.i) x=`<i>${x}</i>`; if(st.u||st.def) x=`<u>${x}</u>`; if(st.s) x=`<s>${x}</s>`; if(st.c) x=`<span style="color:${st.c}">${x}</span>`; return x; }
const R2H=runs=>(runs||[]).map(r=>styleHTML(esc(r.x),r)).join("").replace(/\n/g,"<br>");
const imgTag=(pg,id)=>imgs(pg,id).map(m=>`<figure><img data-file="${m.file}" alt=""></figure>`).join("");
function qHTML(q,pg){ let h=`<div class="q"><b>${q.n}.</b> ${q.stem?R2H(q.stem):imgTag(pg,q.stemImg)}${q.fig?imgTag(pg,q.fig):""}${q.table?tableHTML(q.table):q.tableImg?imgTag(pg,q.tableImg):""}`;
  if(q.choices) h+=q.choices.map((c,i)=>`<div class="ch">(${i+1}) ${c.runs?R2H(c.runs):imgTag(pg,c.img)}</div>`).join("");
  else if(q.parts) h+=q.parts.map(p=>`<div class="ch">${esc(p)}: ______________</div>`).join(""); else h+=`<div class="ch">Answer: ______________ ${q.unit?esc(q.unit):""}</div>`;
  return h+`<div class="ans">Answer: ${R2H(q.answer)}</div></div>`; }
const tableHTML=t=>`<table>${t.rows.map(r=>`<tr>${r.map(c=>`<${c.th?"th":"td"}${c.cs?` colspan="${c.cs}"`:""}>${R2H(c.runs)}</${c.th?"th":"td"}>`).join("")}</tr>`).join("")}</table>`;
function clusterHTML(c,pg){ return `<p class="clh">${R2H(c.head)}</p><p class="clt">${R2H(c.title)}</p>${c.intro?`<p>${R2H(c.intro)}</p>`:imgTag(pg,c.introImg)}${c.cap?`<p style="text-align:center"><b>${R2H(c.cap)}</b></p>`:""}${c.table?tableHTML(c.table):""}${c.fig?imgTag(pg,c.fig):""}`+c.qs.map(q=>qHTML(q,pg)).join(""); }

/* ---------------- pictures ---------------- */
/* pictures are stored one bundle per website page (builder/pics/<page>.json: file name → base64 PNG) */
const BUNDLES={};
const slugOf=f=>f.replace(/-s\d+(-\d+)?\.png$/,"");
function bundle(slug){ return BUNDLES[slug]||(BUNDLES[slug]=fetch("builder/pics/"+slug+".json").then(r=>r.json())); }
async function imgB64(file){ const b=await bundle(slugOf(file)); return "data:image/png;base64,"+b[file]; }
async function imgData(file){ if(IMGCACHE[file]) return IMGCACHE[file]; const b=await bundle(slugOf(file)); const bin=atob(b[file]); const u=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i); IMGCACHE[file]=u.buffer; return u.buffer; }
function hydrate(){ $$("#pv img[data-file]").forEach(im=>{ const f=im.dataset.file; imgB64(f).then(src=>{ if(im.dataset.file===f) im.src=src; }).catch(()=>{}); }); }
function fname(ext,suffix){ return (($("#title").value||"Chemistry Notes").replace(/[\\/:*?"<>|]+/g,"").trim()||"Chemistry Notes")+(suffix?" - "+suffix:"")+"."+ext; }
function save(blob,name){ const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); },4000); }

/* ---------------- Word packet ---------------- */
async function makeDocx(key){ const D=window.docx; const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,LevelFormat,BorderStyle,Footer,PageNumber,TabStopType}=D;
  const FONT="Arial", showAns=key||O.ans==="both";
  const tr=(r,extra={})=>{ const st=r.st||r; const o={text:r.x,font:FONT,bold:!!(st.b||st.def),italics:!!st.i,strike:!!st.s,subScript:!!st.sub,superScript:!!st.sup,color:st.c?st.c.replace("#",""):undefined,underline:(st.u||st.def)?{}:undefined}; return new TextRun(Object.assign(o,extra)); };
  const runsOf=(runs,size)=>{ const out=[]; (runs||[]).forEach(r=>{ if(r.blank){ if(key) out.push(tr(r,{bold:true,underline:{},highlight:"yellow",size})); else out.push(new TextRun({text:"\u00A0".repeat(Math.max(16,Math.round(r.x.length*2.6))),underline:{},font:FONT,size})); return; }
      String(r.x).split("\n").forEach((part,i)=>{ if(i) out.push(new TextRun({break:1,font:FONT})); if(part) out.push(tr(Object.assign({},r,{x:part}),{size})); }); }); return out; };
  /* 1.5 line spacing; while KEEP is on, paragraphs stay on the same page as the next one (a section or question is not split) */
  let KEEP=false;
  const P=(children,o={})=>{ const sp=Object.assign({after:80,line:360,lineRule:"auto"},o.spacing||{}); const k=KEEP&&o.keepNext!==false? {keepNext:true,keepLines:true} : {}; return new Paragraph(Object.assign({children},k,o,{spacing:sp})); };
  const endKeep=()=>{ if(KEEP){ KEEP=false; children.push(P([],{spacing:{after:0,line:240}})); } };
  const estH=(items)=>items.reduce((h,m)=>{ const len=r=>(r||[]).reduce((a,x)=>a+String(x.x).length,0);
    if(m.k==="p"||m.k==="li") return h+Math.ceil(len(m.runs)/92)*0.26+0.06; if(m.k==="box") return h+Math.ceil(len(m.runs)/88)*0.26+0.3;
    if(m.k==="table") return h+m.rows.length*0.32; if(m.k==="img"){ const w=m.w*0.62, hh=m.h*0.62, k=Math.min(1,640/w,420/hh); return h+hh*k/96+0.3; } return h; },0.5);
  const children=[]; const PAGEW=640;
  const addImg=async(m,maxW=PAGEW,maxH=420)=>{ const data=await imgData(m.file); let w=m.w*0.62, h=m.h*0.62; const k=Math.min(1,maxW/w,maxH/h); w*=k; h*=k;
    if(m.label) children.push(P([new TextRun({text:"Example: "+m.label,italics:true,font:FONT,size:18,color:C.body})],{alignment:AlignmentType.CENTER,spacing:{before:120,after:40}}));
    children.push(P([new ImageRun({type:"png",data,transformation:{width:Math.round(w),height:Math.round(h)}})],{alignment:AlignmentType.CENTER,spacing:{before:60,after:120,line:240}})); };
  const tableOf=(rows,size=20)=>{ const n=Math.max(...rows.map(r=>r.reduce((a,c)=>a+(c.cs||1),0)));
    return new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:rows.map(r=>new TableRow({children:r.map(c=>new TableCell({columnSpan:c.cs,rowSpan:c.rs,shading:c.th?{type:ShadingType.CLEAR,fill:C.row2,color:"auto"}:undefined,margins:{top:40,bottom:40,left:80,right:80},children:[P(runsOf(c.runs,size),{alignment:AlignmentType.CENTER,spacing:{after:0}})]}))}))}); };
  const boxOf=(runs,fill)=>new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:[new TableCell({shading:{type:ShadingType.CLEAR,fill,color:"auto"},margins:{top:80,bottom:80,left:140,right:140},borders:{top:{style:BorderStyle.SINGLE,size:8,color:C.boxLine},bottom:{style:BorderStyle.SINGLE,size:8,color:C.boxLine},left:{style:BorderStyle.SINGLE,size:8,color:C.boxLine},right:{style:BorderStyle.SINGLE,size:8,color:C.boxLine}},children:[P(runsOf(runs),{spacing:{after:0}})]})]})]});
  /* title block */
  children.push(P([new TextRun({text:$("#title").value||"Chemistry Notes",bold:true,size:36,font:FONT,color:C.tealInk})],{spacing:{after:60}}));
  children.push(P([new TextRun({text:key?"Teacher Key":"Guided Notes",bold:true,size:22,font:FONT,color:key?C.key:C.tealDeep})],{spacing:{after:120}}));
  if(key) children.push(P([new TextRun({text:"The words students fill in are highlighted.",italics:true,font:FONT,size:20,color:C.body})],{spacing:{after:200}}));
  if(!key) children.push(P([new TextRun({text:"Name: ________________________________     Date: ______________     Period: ______",font:FONT,size:22})],{spacing:{after:240}}));
  const lists={}; let inst=0; const listInst=id=>lists[id]||(lists[id]=++inst);
  const MM=model();
  for(let mi=0;mi<MM.length;mi++){ const m=MM[mi];
    if(m.k==="page"||m.k==="sec"||m.k==="lc") endKeep();
    if(m.k==="sec"){ let j=mi+1; while(j<MM.length&&!["sec","page","lc"].includes(MM[j].k)) j++; KEEP=estH(MM.slice(mi+1,j))<=8.2; }
    if(m.k==="page") children.push(P([new TextRun({text:m.text,bold:true,size:32,font:FONT,color:C.tealInk})],{spacing:{before:280,after:80},border:{bottom:{style:BorderStyle.SINGLE,size:12,color:C.teal,space:2}},keepNext:true}));
    else if(m.k==="sec") children.push(P([new TextRun({text:m.text+(m.enr?" (Enrichment)":""),bold:true,size:26,font:FONT,color:"FFFFFF"})],{shading:{type:ShadingType.CLEAR,fill:m.enr?C.purple:C.teal,color:"auto"},spacing:{before:220,after:100},keepNext:true}));
    else if(m.k==="p") children.push(P(runsOf(m.runs,m.style==="small"||m.style==="cap"?18:undefined),{alignment:m.style==="cap"?AlignmentType.CENTER:undefined}));
    else if(m.k==="li") children.push(P(runsOf(m.runs),{numbering:m.ordered?{reference:"num",level:Math.min(m.level,4),instance:listInst(m.list)}:{reference:"bul",level:Math.min(m.level,4)},spacing:{after:60}}));
    else if(m.k==="box"){ children.push(boxOf(m.runs,m.kind==="cause"?"E7F6F9":C.box)); children.push(P([],{spacing:{after:60}})); }
    else if(m.k==="table"){ children.push(tableOf(m.rows)); children.push(P([],{spacing:{after:60}})); }
    else if(m.k==="img") await addImg(m);
    else if(m.k==="lc"){ children.push(P([new TextRun({text:"Learning Check: "+m.page,bold:true,size:26,font:FONT,color:"FFFFFF"})],{shading:{type:ShadingType.CLEAR,fill:C.orange,color:"auto"},spacing:{before:240,after:100},keepNext:true}));
      for(const it of m.items){ if(it.t==="cluster"){ children.push(P(runsOf(it.head,20).map(x=>x),{spacing:{before:160,after:60}})); children.push(P(runsOf(it.title,24),{alignment:AlignmentType.CENTER}));
          if(it.intro) children.push(P(runsOf(it.intro))); else for(const x of imgs(m.pg,it.introImg)) await addImg(x);
          if(it.cap) children.push(P(runsOf(it.cap),{alignment:AlignmentType.CENTER})); if(it.table){ children.push(tableOf(it.table.rows)); children.push(P([])); } if(it.fig) for(const x of imgs(m.pg,it.fig)) await addImg(x,520,320);
          for(const q of it.qs) await question(q,m.pg); }
        else await question(it,m.pg); } } }
  endKeep();
  async function question(q,pg){ KEEP=true; children.push(P([new TextRun({text:q.n+".  ",bold:true,font:FONT})].concat(q.stem?runsOf(q.stem):[]),{spacing:{before:160,after:60},keepNext:true}));
    if(q.stemImg) for(const x of imgs(pg,q.stemImg)) await addImg(x,560,300);
    if(q.fig) for(const x of imgs(pg,q.fig)) await addImg(x,480,300);
    if(q.table){ children.push(tableOf(q.table.rows,18)); children.push(P([])); } if(q.tableImg) for(const x of imgs(pg,q.tableImg)) await addImg(x,520,300);
    if(q.choices){ for(let i=0;i<q.choices.length;i++){ const c=q.choices[i]; if(c.runs) children.push(P([new TextRun({text:`(${i+1})  `,font:FONT})].concat(runsOf(c.runs)),{indent:{left:440},spacing:{after:40}})); else { for(const x of imgs(pg,c.img)){ const data=await imgData(x.file); let w=x.w*0.62, h=x.h*0.62; const k=Math.min(1,300/w,120/h); children.push(P([new TextRun({text:`(${i+1})   `,font:FONT}),new ImageRun({type:"png",data,transformation:{width:Math.round(w*k),height:Math.round(h*k)}})],{alignment:AlignmentType.CENTER,spacing:{after:60,line:240}})); } } } }
    else if(q.parts) q.parts.forEach(p=>{ children.push(P([new TextRun({text:p+":",bold:true,font:FONT})],{indent:{left:440},spacing:{after:40}})); for(let i=0;i<3;i++) children.push(P([new TextRun({text:"\u00A0".repeat(120),underline:{},font:FONT})],{indent:{left:440},spacing:{after:120}})); });
    else children.push(P([new TextRun({text:"Answer: ",font:FONT}),new TextRun({text:"\u00A0".repeat(30),underline:{},font:FONT}),new TextRun({text:q.unit?" "+q.unit.replace(/<[^>]+>/g,""):"",font:FONT})],{indent:{left:440}}));
    if(showAns) children.push(P([new TextRun({text:"Answer: ",bold:true,font:FONT,color:C.key,size:20})].concat(runsOf(q.answer,20).map(x=>x)),{indent:{left:440},spacing:{after:120}}));
    endKeep(); }
  const lvl=(i,fmt,text)=>({level:i,format:fmt,text,alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:360+i*360,hanging:300}}}});
  const doc=new Document({creator:"Chemistry Notes",title:$("#title").value||"Chemistry Notes",styles:{default:{document:{run:{font:FONT,size:22,color:C.ink}}}},
    numbering:{config:[{reference:"bul",levels:["•","◦","▪","•","◦"].map((c,i)=>lvl(i,LevelFormat.BULLET,c))},{reference:"num",levels:[0,1,2,3,4].map(i=>lvl(i,LevelFormat.DECIMAL,`%${i+1}.`))}]},
    sections:[{properties:{page:{margin:{top:900,bottom:900,left:1000,right:1000}}},footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:($("#title").value||"")+"  ·  ",size:16,color:C.body,font:FONT}),new TextRun({children:[PageNumber.CURRENT],size:16,color:C.body,font:FONT})]})]})},children}]});
  save(await Packer.toBlob(doc),fname("docx",key?"Teacher Key":"Guided Notes")); }

/* ---------------- slides ---------------- */
async function makePptx(){ const pptx=new PptxGenJS(); pptx.layout="LAYOUT_WIDE"; pptx.title=$("#title").value||"Chemistry Notes";
  const SW=13.333, SH=7.5, M=0.5, TOP=1.25, BOT=6.95, FONT="Arial", FS=20;
  const title=$("#title").value||"Chemistry Notes";
  /* title slide */
  let s=pptx.addSlide(); s.background={color:C.teal}; s.addText(title,{x:M,y:2.3,w:SW-2*M,h:1.4,fontFace:FONT,fontSize:40,bold:true,color:"FFFFFF",align:"center"});
  s.addText(order().map(x=>x.page.title).filter((v,i,a)=>a.indexOf(v)===i).join("  ·  "),{x:M,y:3.8,w:SW-2*M,h:1,fontFace:FONT,fontSize:18,color:"FFFFFF",align:"center"});
  const bar=(sl,t,color,sub)=>{ sl.background={color:"FFFFFF"}; sl.addShape(pptx.ShapeType.rect,{x:0,y:0,w:SW,h:1.0,fill:{color},line:{color}}); sl.addText(t,{x:M,y:0.08,w:SW-2*M,h:0.84,fontFace:FONT,fontSize:28,bold:true,color:"FFFFFF",valign:"middle"}); if(sub) sl.addText(sub,{x:M,y:SH-0.42,w:SW-2*M,h:0.3,fontFace:FONT,fontSize:11,color:C.body}); };
  const pr=(r,extra)=>{ const st=r.st||r; return {text:r.x,options:Object.assign({fontFace:FONT,bold:!!(st.b||st.def),italic:!!st.i,underline:(st.u||st.def)?{style:"sng"}:undefined,strike:st.s?"sngStrike":undefined,subscript:!!st.sub,superscript:!!st.sup,color:st.c?st.c.replace("#",""):C.ink},extra||{})}; };
  const prRuns=(runs)=>{ const out=[]; (runs||[]).forEach(r=>{ if(r.blank){ out.push(pr(r,{bold:true,color:C.tealDeep,underline:{style:"sng"}})); return; } String(r.x).split("\n").forEach((part,i)=>{ if(i) out.push({text:"",options:{breakLine:true}}); if(part) out.push(pr(Object.assign({},r,{x:part}))); }); }); return out; };
  const cellRuns=r=>{ const a=prRuns(r); return a.length?a:[{text:" ",options:{fontFace:FONT}}]; };  /* an empty table cell still needs one run, or PowerPoint repairs the file */
  const plain=runs=>(runs||[]).map(r=>r.x).join("");
  /* paragraph list → pptx text array (one paragraph per entry) */
  const paraArr=(paras)=>{ const out=[]; paras.forEach((p,pi)=>{ const rr=prRuns(p.runs); if(!rr.length) rr.push({text:" ",options:{}}); const fs=p.level>0?FS-2:FS;
      rr.forEach((x,i)=>{ x.options.fontSize=p.size||fs; if(i===0){ if(p.bullet) x.options.bullet=p.ordered?{type:"number",indent:22}:{indent:22}; if(p.bullet||p.level) x.options.indentLevel=p.level; if(p.italic) x.options.italic=true; if(p.color) x.options.color=p.color; x.options.paraSpaceAfter=4; } });
      rr[rr.length-1].options.breakLine=pi<paras.length-1; out.push(...rr); }); return out; };
  const lines=(p,wIn)=>{ const cpl=Math.max(20,Math.floor(wIn*72/((p.size||FS)*0.56))); const t=plain(p.runs); return Math.max(1,Math.ceil((t.length+(p.level||0)*3)/cpl))+((t.match(/\n/g)||[]).length); };
  const LH=(p)=>(p.size||FS)/72*1.32+0.06;
  /* lay out one section: stack text, boxes, tables and pictures; text next to a picture when it fits */
  const sectionSlides=async(head,color,items,foot)=>{ let sl=null, y=TOP; const fresh=()=>{ sl=pptx.addSlide(); bar(sl,head,color,foot); y=TOP; };
    const room=()=>BOT-y;
    const textBlock=(paras,x,w,yy,maxH)=>{ sl.addText(paraArr(paras),{x,y:yy,w,h:maxH,fontFace:FONT,valign:"top",margin:2}); };
    const heightOf=(paras,w)=>paras.reduce((a,p)=>a+lines(p,w)*LH(p),0.1);
    let i=0; fresh();
    while(i<items.length){ const it=items[i];
      if(it.k==="text"){ /* a run of paragraphs, maybe split over slides */
        let paras=it.paras; const nxt=items[i+1];
        if(nxt&&nxt.k==="img"&&nxt.w/nxt.h<1.7&&y===TOP){ const tw=6.0, h=heightOf(paras,tw); if(h<=BOT-TOP){ textBlock(paras,M,tw,y,h+0.2); await picture(nxt,M+tw+0.2,y,SW-M-(M+tw+0.2),BOT-y); i+=2; if(i<items.length) fresh(); continue; } }
        while(paras.length){ let n=0, h=0.1; while(n<paras.length&&h+lines(paras[n],SW-2*M)*LH(paras[n])<=room()) { h+=lines(paras[n],SW-2*M)*LH(paras[n]); n++; }
          if(n===0){ if(y>TOP){ fresh(); continue; } n=1; h=Math.min(BOT-TOP,lines(paras[0],SW-2*M)*LH(paras[0])); }
          textBlock(paras.slice(0,n),M,SW-2*M,y,h+0.1); y+=h+0.12; paras=paras.slice(n); if(paras.length) fresh(); }
        i++; continue; }
      if(it.k==="box"){ const h=heightOf([{runs:it.runs}],SW-2*M-0.4)+0.25; if(h>room()&&y>TOP) fresh();
        sl.addText(paraArr([{runs:it.runs}]),{x:M,y,w:SW-2*M,h:Math.min(h,BOT-y),fill:{color:it.kind==="cause"?C.row:C.box},line:{color:C.boxLine,width:1.25},rectRadius:0.08,shape:pptx.ShapeType.roundRect,margin:8,valign:"middle",fit:"shrink"}); y+=h+0.15; i++; continue; }
      if(it.k==="table"){ const rows=it.rows.length, h=rows*0.42+0.1; if(h>room()&&y>TOP) fresh();
        const fs=rows>9?12:rows>6?14:16; sl.addTable(it.rows.map(r=>r.map(c=>({text:cellRuns(c.runs).map(x=>{ x.options.fontSize=fs; return x; }),options:{fill:{color:c.th?C.row2:"FFFFFF"},colspan:c.cs,rowspan:c.rs,align:"center",valign:"middle",fontFace:FONT,fontSize:fs}}))),{x:M,y,w:SW-2*M,border:{type:"solid",pt:1,color:"9FC4D6"},fontFace:FONT,autoPage:false});
        y+=Math.min(h,room())+0.15; i++; continue; }
      if(it.k==="img"){ const need=Math.min(BOT-TOP,(it.h/it.w)*(SW-2*M)); if(room()<Math.min(2.6,need)&&y>TOP) fresh(); await picture(it,M,y,SW-2*M,room()); i++; continue; }
      i++; }
    async function picture(m,x,yy,w,h){ let labH=0; if(m.label){ sl.addText("Example: "+m.label,{x,y:yy,w,h:0.35,fontFace:FONT,fontSize:14,italic:true,color:C.body,align:"center"}); labH=0.4; }
      const data=await imgB64(m.file); let iw=m.w/96, ih=m.h/96; const k=Math.min(w/iw,(h-labH)/ih,1.6); iw*=k; ih*=k; sl.addImage({data,x:x+(w-iw)/2,y:yy+labH,w:iw,h:ih}); y=yy+labH+ih+0.15; } };
  /* group the document model into sections */
  const M2=model(); let cur=null, pageT="", unitT=""; const secs=[];
  M2.forEach(m=>{ if(m.k==="page"){ pageT=m.text; unitT=m.unit; return; }
    if(m.k==="sec"){ cur={head:m.text+(m.enr?" (Enrichment)":""),color:m.enr?C.purple:C.teal,items:[],foot:[unitT,pageT].filter(Boolean).join("  ·  ")}; secs.push(cur); return; }
    if(m.k==="lc"){ secs.push({lc:m}); cur=null; return; }
    if(!cur) return; const L=cur.items[cur.items.length-1];
    const para= m.k==="p"? {runs:m.runs,level:0,bullet:false,italic:m.style==="intro",color:m.style==="intro"?C.body:undefined,size:m.style==="small"||m.style==="cap"?15:undefined} : m.k==="li"? {runs:m.runs,level:m.level,bullet:true,ordered:m.ordered} : null;
    if(para){ if(L&&L.k==="text") L.paras.push(para); else cur.items.push({k:"text",paras:[para]}); return; }
    cur.items.push(m); });
  for(const sc of secs){
    if(sc.lc){ const m=sc.lc, pg=m.pg; const foot="Learning Check  ·  "+m.page;
      const qSlide=async(q,ctxTitle)=>{ const sl=pptx.addSlide(); bar(sl,ctxTitle||`Learning Check: ${m.page}`,C.orange,foot); let y=TOP;
        const stem=[{runs:[{x:q.n+". ",b:1}].concat(q.stem||[]),level:0}]; sl.addText(paraArr(stem),{x:M,y,w:SW-2*M,h:1.1,fontFace:FONT,valign:"top",fit:"shrink"}); y+=1.15;
        const pics=[].concat(q.stemImg?imgs(pg,q.stemImg):[],q.fig?imgs(pg,q.fig):[],q.tableImg?imgs(pg,q.tableImg):[]);
        const chH=q.choices?Math.min(3.2,q.choices.length*0.55):0.6;
        for(const p of pics){ const data=await imgB64(p.file); let iw=p.w/96, ih=p.h/96; const k=Math.min((SW-2*M)/iw,(BOT-y-chH-0.1)/ih,1.4); iw*=k; ih*=k; sl.addImage({data,x:(SW-iw)/2,y,w:iw,h:ih}); y+=ih+0.1; }
        if(q.table){ sl.addTable(q.table.rows.map(r=>r.map(c=>({text:cellRuns(c.runs).map(x=>{ x.options.fontSize=14; return x; }),options:{fill:{color:c.th?C.row2:"FFFFFF"},colspan:c.cs,align:"center",fontFace:FONT,fontSize:14}}))),{x:M+1,y,w:SW-2*M-2,border:{type:"solid",pt:1,color:"9FC4D6"}}); y+=q.table.rows.length*0.4+0.1; }
        if(q.choices){ for(let i=0;i<q.choices.length;i++){ const c=q.choices[i]; const cw=(SW-2*M)/2, cx=M+(i%2)*cw, cy=y+Math.floor(i/2)*((BOT-y)/2);
            if(c.runs) sl.addText(paraArr([{runs:[{x:`(${i+1})  `,b:1}].concat(c.runs),level:0}]),{x:cx,y:cy,w:cw-0.2,h:Math.min(1,(BOT-y)/2),fontFace:FONT,valign:"top",fit:"shrink"});
            else { sl.addText(`(${i+1})`,{x:cx,y:cy,w:0.6,h:0.4,fontFace:FONT,fontSize:FS,bold:true,color:C.ink}); for(const p of imgs(pg,c.img)){ const data=await imgB64(p.file); let iw=p.w/96, ih=p.h/96; const k=Math.min((cw-0.9)/iw,((BOT-y)/2-0.1)/ih,1.4); sl.addImage({data,x:cx+0.65,y:cy,w:iw*k,h:ih*k}); } } } }
        if(!q.choices){ const lab=q.parts? q.parts.map(p=>p+": ________________________________________________").join("\n\n") : "Answer: ______________________"+(q.unit?" "+String(q.unit).replace(/<[^>]+>/g,""):""); sl.addText(lab,{x:M,y:Math.min(y+0.2,BOT-1.5),w:SW-2*M,h:1.4,fontFace:FONT,fontSize:FS,color:C.ink,valign:"top"}); }
        sl.addNotes("Answer: "+plain(q.answer)); };
      for(const it of m.items){ if(it.t==="cluster"){ const sl=pptx.addSlide(); bar(sl,plain(it.title)||"Learning Check",C.orange,foot); let y=TOP;
          sl.addText(paraArr([{runs:it.head,level:0,italic:true,size:15}]),{x:M,y,w:SW-2*M,h:0.5,fontFace:FONT}); y+=0.55;
          if(it.intro){ const h=Math.min(2.6,lines({runs:it.intro},SW-2*M)*LH({})+0.1); sl.addText(paraArr([{runs:it.intro,level:0}]),{x:M,y,w:SW-2*M,h,fontFace:FONT,valign:"top",fit:"shrink"}); y+=h+0.1; }
          const pics=[].concat(it.introImg?imgs(pg,it.introImg):[],it.fig?imgs(pg,it.fig):[]); for(const p of pics){ const data=await imgB64(p.file); let iw=p.w/96, ih=p.h/96; const k=Math.min((SW-2*M)/iw,(BOT-y)/ih,1.4); iw*=k; ih*=k; sl.addImage({data,x:(SW-iw)/2,y,w:iw,h:ih}); y+=ih+0.1; }
          if(it.table&&y<BOT-0.8){ sl.addTable(it.table.rows.map(r=>r.map(c=>({text:cellRuns(c.runs).map(x=>{ x.options.fontSize=14; return x; }),options:{fill:{color:c.th?C.row2:"FFFFFF"},colspan:c.cs,align:"center",fontFace:FONT,fontSize:14}}))),{x:M+1,y,w:SW-2*M-2,border:{type:"solid",pt:1,color:"9FC4D6"}}); }
          for(const q of it.qs) await qSlide(q,plain(it.title)); }
        else await qSlide(it); } continue; }
    await sectionSlides(sc.head,sc.color,sc.items,sc.foot); }
  /* pptxgenjs writes an extra paragraph-property tag between runs of one paragraph; remove it so PowerPoint opens the file cleanly */
  const ab=await pptx.write({outputType:"arraybuffer"}); const z=await JSZip.loadAsync(ab);
  for(const f of Object.keys(z.files).filter(f=>/ppt\/slides\/slide\d+\.xml$/.test(f))){ let x=await z.file(f).async("string"); x=x.replace(/(<\/a:r>)<a:pPr\b[^>]*?(?:\/>|>.*?<\/a:pPr>)/g,"$1"); z.file(f,x); }
  save(await z.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}),fname("pptx","Slides")); }

window.__pb={sel,O,MAN,PRE,model,makeDocx,makePptx,setPage:(s,on)=>{ setPage(s,on); syncTree(); render(); },render,start};
})();
