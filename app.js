(function(){
const S={data:JSON.parse(document.getElementById('seed').textContent),source:'no data loaded yet — use Load Excel or Link data file',dirty:false};
const TABS=[['overview','Overview'],['projects','Alternate Development'],['mqis','MQIS & Single Phase Parts'],['ds6','6DS Project Parts'],['fan','DC Fan Comparison'],['teknic','Teknic Parts (Domestic)']];
const SHEETS={mqis:'MQIS AND SINGLE PHASE PART',ds6:'6DS PROJECT PARTS',projects:'ALTERNATE DEVELOPNMENT',fan:'DC Fan Comparison',teknic:'Teknic Parts (Domestic)'};
const $=s=>document.querySelector(s);
const esc=v=>v==null?'':String(v).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const today=new Date();today.setHours(0,0,0,0);
const inr=v=>typeof v==='number'&&isFinite(v)?'₹ '+v.toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2}):'';
const toDate=v=>{if(v==null||v==='')return null;if(v instanceof Date)return v;const d=new Date(String(v).slice(0,10)+'T00:00:00');return isNaN(d)?null:d};
const fmtD=d=>d?d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'2-digit'}):'';
const projKey=p=>{p=String(p||'');if(/single/i.test(p))return 'Single';if(/p1r5/i.test(p))return 'P1R5';if(/6ds/i.test(p))return '6DS';return 'Common'};
const stCls=s=>{s=String(s||'').trim();if(s==='Completed')return 's-Completed';if(s==='Ongoing')return 's-Ongoing';if(/hold/i.test(s))return 's-Hold';if(/drop/i.test(s))return 's-Dropped';return 's-none'};
const num=v=>{if(v===''||v==null)return null;const n=Number(String(v).replace(/[₹,\s]/g,''));return isFinite(n)?n:null};
const blank=v=>v==null||String(v).trim()==='';
function toast(m){const t=$('#toast');t.textContent=m;t.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>t.hidden=true,4500)}

/* ---------- normalise data model ---------- */
function teknicFromRaw(raw){
  const rows=raw.filter(r=>r&&r.some(v=>v!=null&&v!==''));
  const title=rows[0]?rows[0][0]:'Teknic Parts (Domestic)';
  const parts=rows.filter(r=>typeof r[0]==='number'&&r[1]).map(r=>({pn:r[0],name:r[1],make:r[2],mpn:r[3],bom:r[4],exs:r[5],alts:r[6],exp:r[7],qp:r[8]}));
  const st=rows.find(r=>/Stations per year/i.test(String(r[0]||'')));
  return {title,parts,stations:st&&typeof st[1]==='number'?st[1]:null};
}
function norm(d){
  if(!d.tk)d.tk=teknicFromRaw(d.teknic||[]);
  if(!d.fanExtra)d.fanExtra=[];
  return d;
}
norm(S.data);

/* ---------- save: real data file (primary) + browser storage (fallback) ---------- */
const KEY='npdSourcingTracker.v1';let canSave=true,lastSaved=null,lastBackup=null;
let fh=null,fhOK=false,fileErr='';const FSA=!!(window.showSaveFilePicker&&window.showOpenFilePicker&&window.indexedDB);
try{const raw=localStorage.getItem(KEY);if(raw){const o=JSON.parse(raw);if(o&&o.data){S.data=norm(o.data);lastSaved=o.savedAt;lastBackup=o.backupAt||null;S.source='saved in this browser'}}}catch(e){canSave=false}
function idb(mode,fn){return new Promise((res,rej)=>{const r=indexedDB.open('npdTracker',1);r.onupgradeneeded=()=>r.result.createObjectStore('h');r.onerror=()=>rej(r.error);r.onsuccess=()=>{const tx=r.result.transaction('h',mode);const st=tx.objectStore('h');let q;try{q=fn(st)}catch(e){rej(e);return}tx.oncomplete=()=>res(q&&q.result);tx.onerror=()=>rej(tx.error)}})}
const fileBlob=()=>JSON.stringify({app:'NPD Sourcing Tracker',savedAt:new Date().toISOString(),data:S.data},null,1);
async function writeFile(){if(!fh||!fhOK)return false;try{const w=await fh.createWritable();await w.write(fileBlob());await w.close();fileErr='';return true}catch(e){fileErr=e.message||String(e);fhOK=false;return false}}
async function readFile(){const f=await fh.getFile();const t=await f.text();if(!t.trim())return false;const o=JSON.parse(t);if(!o||!o.data||!o.data.projects)throw new Error(f.name+' is not a tracker data file.');S.data=norm(o.data);S.source=f.name;lastSaved=o.savedAt||null;return true}
async function linkFile(createNew){
  try{
    if(createNew){fh=await window.showSaveFilePicker({suggestedName:'NPD_Tracker_Data.json',types:[{description:'Tracker data',accept:{'application/json':['.json']}}]});fhOK=true;await writeFile();}
    else{[fh]=await window.showOpenFilePicker({types:[{description:'Tracker data',accept:{'application/json':['.json']}}]});fhOK=(await fh.requestPermission({mode:'readwrite'}))==='granted';if(fhOK)await readFile();}
    await idb('readwrite',st=>st.put(fh,'file')).catch(()=>{});
    lastSaved=new Date().toISOString();try{localStorage.setItem(KEY,JSON.stringify({data:S.data,savedAt:lastSaved,backupAt:lastBackup}))}catch(e){}
    render();toast(fhOK?'Linked to '+fh.name+'. Every change now saves to this file.':'Permission was not given, so the file is not linked.');
  }catch(e){if(e.name!=='AbortError')toast('Could not link the file: '+e.message)}
}
async function reconnect(){try{fhOK=(await fh.requestPermission({mode:'readwrite'}))==='granted';if(fhOK){await readFile();lastSaved=lastSaved||new Date().toISOString();render();toast('Connected to '+fh.name+'.')}else toast('Permission not given. Click Connect again and choose Allow.')}catch(e){toast('Could not open '+(fh&&fh.name)+': '+e.message)}}
async function initFile(){if(!FSA)return;try{fh=await idb('readonly',st=>st.get('file'))||null;if(fh){fhOK=(await fh.queryPermission({mode:'readwrite'}))==='granted';if(fhOK)await readFile()}}catch(e){fh=null}render()}
async function persist(){let okLS=false;try{localStorage.setItem(KEY,JSON.stringify({data:S.data,savedAt:new Date().toISOString(),backupAt:lastBackup}));okLS=true}catch(e){canSave=false}
  const okF=await writeFile();if(okF||okLS)lastSaved=new Date().toISOString();return {okF,okLS}}
async function save(msg){const r=await persist();S.dirty=!(r.okF||r.okLS);render();
  if(r.okF)toast(msg+' — saved to '+fh.name+'.');
  else if(fh&&!fhOK)toast(msg+' — NOT saved to the data file. Click "Connect '+fh.name+'" at the top, then save again.');
  else if(r.okLS)toast(msg+' — saved in this browser only. Link a data file (top right) to keep it safe.');
  else toast(msg+' — NOT saved. This window blocks saving. Link a data file or use Backup now.')}
function ago(iso){if(!iso)return 'never';const d=(Date.now()-new Date(iso))/864e5;return d<1?'today '+new Date(iso).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'}):Math.floor(d)+' day'+(d>=2?'s':'')+' ago'}
function paintSave(){const el=$('#saveState'),fb=$('#fileBtn');if(!el)return;
  if(fb){if(!FSA){fb.hidden=true}else{fb.hidden=false;if(fh&&fhOK){fb.textContent='Data file: '+fh.name;fb.className='btn filebtn ok';fb.dataset.mode='change'}else if(fh){fb.textContent='Connect '+fh.name;fb.className='btn filebtn warn';fb.dataset.mode='reconnect'}else{fb.textContent='Link data file';fb.className='btn filebtn warn';fb.dataset.mode='link'}}}
  if(fh&&fhOK){el.textContent='Saved '+ago(lastSaved);el.className='savestate';return}
  if(fh&&!fhOK){el.textContent='Not connected — changes not saved to file';el.className='savestate warn';return}
  if(!canSave){el.textContent='Saving blocked here — link a data file';el.className='savestate warn';return}
  el.textContent='Browser only · saved '+ago(lastSaved);el.className='savestate warn'}
function download(name,blob){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(u);a.remove()},500)}
function fileMenu(){const mode=$('#fileBtn').dataset.mode;if(mode==='reconnect')return reconnect();
  openForm({title:mode==='change'?'Data file':'Link a data file',fields:[{k:'how',label:'Choose',type:'select',opts:['Create a new data file (saves current data)','Open an existing data file']}],
    submit:v=>{linkFile(v.how.startsWith('Create'))}})}
/* ---------- modal form ---------- */
function openForm({title,fields,values={},submit,del}){
  const m=$('#modal');
  m.innerHTML=`<form class="mbox" id="mform" novalidate><h2>${esc(title)}</h2><div class="fgrid">${fields.map((f,i)=>{
    const v=values[f.k]??(f.def??'');const id='f_'+f.k;
    let input;
    if(f.type==='select')input=`<select id="${id}" name="${f.k}">${f.opts.map(o=>`<option ${String(o)===String(v)?'selected':''}>${esc(o)}</option>`).join('')}</select>`;
    else if(f.type==='textarea')input=`<textarea id="${id}" name="${f.k}" rows="3">${esc(v)}</textarea>`;
    else if(f.type==='check')input=`<label class="chk"><input type="checkbox" id="${id}" name="${f.k}" ${v?'checked':''}> ${esc(f.hint||'')}</label>`;
    else input=`<input id="${id}" name="${f.k}" type="${f.type||'text'}" ${f.type==='number'?'step="any"':''} value="${esc(f.type==='date'&&v?String(v).slice(0,10):v)}" ${f.list?`list="dl_${f.k}"`:''}>${f.list?`<datalist id="dl_${f.k}">${f.list.map(o=>`<option value="${esc(o)}">`).join('')}</datalist>`:''}`;
    return `<div class="fld ${f.wide?'wide':''}">${f.type==='check'?'':`<label for="${id}">${esc(f.label)}${f.req?' <span class="req">*</span>':''}</label>`}${input}${f.note?`<small>${esc(f.note)}</small>`:''}</div>`}).join('')}</div>
    <p class="ferr" id="ferr" hidden></p>
    <div class="mact">${del?'<button type="button" class="btn danger" id="mdel">Delete</button>':''}<span style="flex:1"></span><button type="button" class="btn" id="mcancel">Cancel</button><button type="submit" class="btn primary">Save</button></div></form>`;
  m.hidden=false;
  const first=m.querySelector('input,select,textarea');if(first)first.focus();
  $('#mcancel').onclick=()=>{m.hidden=true};
  if(del){const b=$('#mdel');b.onclick=()=>{if(b.dataset.armed){m.hidden=true;del()}else{b.dataset.armed=1;b.textContent='Click again to delete'}}}
  $('#mform').onsubmit=e=>{e.preventDefault();
    const out={};fields.forEach(f=>{const el=document.getElementById('f_'+f.k);out[f.k]=f.type==='check'?el.checked:f.type==='number'?num(el.value):el.value.trim()});
    const miss=fields.filter(f=>f.req&&blank(out[f.k]));
    if(miss.length){showErr('Fill in: '+miss.map(f=>f.label).join(', '));return}
    const err=submit(out);if(err){showErr(err);return}
    m.hidden=true};
  function showErr(t){const p=$('#ferr');p.textContent=t;p.hidden=false}
}
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('#modal').hidden)$('#modal').hidden=true});

/* ---------- stats ---------- */
function projStats(p){
  let total=0,done=0,late=[],soon=0;
  p.acts.forEach(a=>{
    if(typeof a[0]!=='number'||/trigger/i.test(a[1]||''))return;
    const pl=toDate(a[2]),ac=toDate(a[3]);
    if(!pl&&!ac)return;total++;
    if(ac)done++;else if(pl&&pl<today)late.push({act:a[1],plan:pl,days:Math.round((today-pl)/864e5),rem:a[4]});else if(pl&&(pl-today)/864e5<=7)soon++;
  });
  return {total,done,late,soon,pct:total?done/total:0};
}
const PCFG={
  mqis:{num:[4,11],c:[0,5],mpn:[8,9,10],price:11,proj:r=>r[7]||''},
  ds6:{num:[5,6,9],c:[0,3,4],mpn:[8],price:9,proj:()=>'6DS'}
};
const hasMpn=(key,r)=>PCFG[key].mpn.some(i=>!blank(r[i]));
function fanCalc(f){
  const pINR=(f.pUSD||0)*(f.fx||0),fr=f.freight||0,ins=(pINR+fr)*(f.insPct||0)/100,cif=pINR+fr+ins,bcd=cif*(f.bcdPct||0)/100,sws=bcd*(f.swsPct||0)/100,cha=f.cha||0;
  const landed=cif+bcd+sws+cha,alt=(f.altPrice||0)+(f.altFreight||0),save=landed-alt;
  return {pINR,ins,cif,bcd,sws,landed,alt,save,pct:landed?save/landed:0,annual:save*(f.vol||0)};
}

/* ---------- views ---------- */
const V={};
V.overview=()=>{
  const d=S.data,ps=d.projects.map(p=>({p,st:projStats(p)}));
  const comp=ps.filter(x=>x.p.status==='Completed').length,ong=ps.filter(x=>x.p.status==='Ongoing').length;
  const late=ps.flatMap(x=>x.st.late.map(l=>({...l,proj:x.p.name}))).sort((a,b)=>b.days-a.days);
  const tot=ps.reduce((s,x)=>s+x.st.total,0),dn=ps.reduce((s,x)=>s+x.st.done,0);
  const all=[...d.mqis[1].map(r=>({r,k:'mqis'})),...d.ds6[1].map(r=>({r,k:'ds6'}))];
  const parts=all.length,withMpn=all.filter(x=>hasMpn(x.k,x.r)).length,withPrice=all.filter(x=>typeof x.r[PCFG[x.k].price]==='number').length;
  const byProj={};d.mqis[1].forEach(r=>{const k=projKey(r[7]);byProj[k]=(byProj[k]||0)+1});byProj['6DS']=d.ds6[1].length;
  const pct=(a,b)=>b?Math.round(a/b*100):0;
  return `
  <p class="note">Data: ${esc(S.source)} · Today ${fmtD(today)} · Delay = plan date passed with no actual date.</p>
  <section><div class="kpis">
    <div class="kpi"><span class="lab">Projects</span><span class="val">${d.projects.length}</span><span class="sub"><span style="color:var(--ok)">${comp} completed</span> · <span style="color:var(--warn)">${ong} ongoing</span></span></div>
    <div class="kpi"><span class="lab">Activities done</span><span class="val">${dn}<span class="sm"> / ${tot}</span></span><div class="meter"><i style="width:${pct(dn,tot)}%"></i></div><span class="sub">${pct(dn,tot)}% across all projects</span></div>
    <div class="kpi"><span class="lab">Delayed activities</span><span class="val" style="color:var(--bad)">${late.length}</span><span class="sub">plan date passed, not closed</span></div>
    <div class="kpi"><span class="lab">Parts tracked</span><span class="val">${parts}</span><span class="sub">${Object.entries(byProj).map(([k,v])=>`<span class="proj proj-${k}">${k==='Single'?'Single Phs':k} ${v}</span>`).join(' ')}</span></div>
    <div class="kpi"><span class="lab">MPN coverage</span><span class="val">${pct(withMpn,parts)}%</span><div class="meter"><i style="width:${pct(withMpn,parts)}%;background:var(--ok)"></i></div><span class="sub">${parts-withMpn} parts without MAKE/MPN</span></div>
    <div class="kpi"><span class="lab">Price coverage</span><span class="val">${pct(withPrice,parts)}%</span><div class="meter"><i style="width:${pct(withPrice,parts)}%;background:var(--warn)"></i></div><span class="sub">${parts-withPrice} parts without price</span></div>
  </div></section>
  <section><div class="h2row"><h2>Project status</h2><button class="btn sm-btn" type="button" data-act="addProject">＋ Add project</button></div><div class="cards">${ps.map((x,i)=>`
    <button class="card" type="button" data-go="${i}">
      <div class="row"><h3>${esc(x.p.name)}</h3><span class="pill ${stCls(x.p.status)}">${esc(x.p.status||'Not set')}</span></div>
      <div class="meter"><i style="width:${Math.round(x.st.pct*100)}%;background:${x.p.status==='Completed'?'var(--ok)':x.st.late.length?'var(--bad)':'var(--warn)'}"></i></div>
      <div class="row meta"><span>${x.st.done} of ${x.st.total} activities done</span><span>${x.st.late.length?`<span class="flag late">${x.st.late.length} delayed</span>`:x.st.soon?`<span class="flag soon">${x.st.soon} due ≤7d</span>`:'<span class="flag done">on track</span>'}</span></div>
    </button>`).join('')}</div></section>
  <section class="split">
    <div class="panel"><h2>Delayed activities</h2>${late.length?`<ul class="alerts">${late.slice(0,12).map(l=>`<li><span class="d">${l.days} d</span><span><b>${esc(l.proj)}</b> — ${esc(l.act)}</span><span class="r">Plan ${fmtD(l.plan)}${l.rem?' · '+esc(l.rem):''}</span></li>`).join('')}</ul>${late.length>12?`<p class="sm">+${late.length-12} more in Alternate Development</p>`:''}`:'<p class="sm">No delayed activities.</p>'}</div>
    <div class="panel"><h2>Parts missing data</h2>${missingList()}</div>
  </section>`;
};
function missingList(){
  const rows=[];
  [['mqis',r=>projKey(r[7])],['ds6',()=>'6DS']].forEach(([k,pk])=>S.data[k][1].forEach(r=>{const noM=!hasMpn(k,r),noP=typeof r[PCFG[k].price]!=='number';if(noM||noP)rows.push({pn:r[1],name:r[2],proj:pk(r),noM,noP})}));
  if(!rows.length)return '<p class="sm">All parts have MAKE/MPN and price.</p>';
  return `<div class="tablebox" style="max-height:360px;overflow:auto"><table><thead><tr><th>Part No</th><th>Part Name</th><th>Project</th><th>Missing</th></tr></thead><tbody>${rows.map(r=>`<tr><td class="mono">${esc(r.pn)}</td><td>${esc(r.name)}</td><td><span class="proj proj-${r.proj}">${r.proj==='Single'?'Single Phs':r.proj}</span></td><td>${r.noM?'<span class="miss">MPN</span> ':''}${r.noP?'<span class="miss">Price</span>':''}</td></tr>`).join('')}</tbody></table></div>`;
}
V.projects=(open)=>`<div class="tools"><p class="note" style="margin:0;flex:1">Red date = plan passed with no actual · amber = due within 7 days · green = closed. Click any activity row to update dates or remarks.</p><button class="btn primary" type="button" data-act="addProject">＋ Add project</button></div>`+
  S.data.projects.map((p,i)=>{const st=projStats(p);
    return `<details class="pj" id="pj${i}" ${open===i?'open':''}><summary><h3>${esc(p.name)}</h3><span class="pill ${stCls(p.status)}">${esc(p.status||'Not set')}</span>
      <div class="meter"><i style="width:${Math.round(st.pct*100)}%;background:${p.status==='Completed'?'var(--ok)':st.late.length?'var(--bad)':'var(--warn)'}"></i></div>
      <span class="sm">${st.done} of ${st.total} activities done${st.late.length?` · <span style="color:var(--bad)">${st.late.length} delayed</span>`:''}</span></summary>
      <div class="pjbar"><button class="btn sm-btn" type="button" data-act="addAct" data-p="${i}">＋ Add activity</button><button class="btn sm-btn" type="button" data-act="editProject" data-p="${i}">Edit project / status</button></div>
      <div class="tablebox" style="border:0;border-radius:0 0 8px 8px"><table class="steps"><thead><tr><th class="c">#</th><th>Stage-gate activity</th><th>Plan</th><th>Actual</th><th>Flag</th><th>Remarks / delay reason</th></tr></thead><tbody>${p.acts.map((a,j)=>{
        const pl=toDate(a[2]),ac=toDate(a[3]),trig=/trigger/i.test(a[1]||'');
        let f='';if(ac)f='<span class="flag done">Done</span>';else if(pl&&!trig&&pl<today)f=`<span class="flag late">${Math.round((today-pl)/864e5)} d late</span>`;else if(pl&&!trig&&(pl-today)/864e5<=7)f='<span class="flag soon">Due soon</span>';
        return `<tr class="click" data-act="editAct" data-p="${i}" data-a="${j}" tabindex="0"><td>${esc(a[0])}</td><td>${esc(a[1])}</td><td class="dt ${!ac&&pl&&!trig&&pl<today?'late':''}">${pl?fmtD(pl):esc(a[2]||'')}</td><td class="dt ${ac?'done':''}">${ac?fmtD(ac):esc(a[3]||'')}</td><td>${f}</td><td style="min-width:260px">${esc(a[4])}</td></tr>`}).join('')}</tbody></table></div></details>`}).join('');
function partsView(key){
  const [h,rows]=S.data[key],cfg=PCFG[key];
  const projOpts=[...new Set(rows.map(r=>cfg.proj(r)))].filter(Boolean);
  return `<div class="tools"><input id="q" type="search" placeholder="Search part no, name, supplier, MPN…" aria-label="Search parts">
    ${projOpts.length>1?`<select id="pf" aria-label="Filter by project"><option value="">All projects</option>${projOpts.map(p=>`<option>${esc(p)}</option>`).join('')}</select>`:''}
    <label class="chk"><input type="checkbox" id="mf"> Only missing MPN/price</label><span class="sm" id="cnt"></span><span style="flex:1"></span>
    <button class="btn primary" type="button" data-act="addPart" data-k="${key}">＋ Add part</button></div>
    <p class="note">Click a row to edit it.</p>
    <div class="tablebox"><table><thead><tr>${h.map((x,i)=>`<th class="${cfg.num.includes(i)?'num':cfg.c.includes(i)?'c':''}">${esc(x)}</th>`).join('')}</tr></thead><tbody id="tb"></tbody></table></div>`;
}
function fillParts(key){
  const cfg=PCFG[key],[h,rows]=S.data[key];
  const q=($('#q').value||'').toLowerCase(),pf=$('#pf')?$('#pf').value:'',mf=$('#mf').checked;
  const out=rows.map((r,i)=>({r,i})).filter(({r})=>{
    if(pf&&cfg.proj(r)!==pf)return false;
    if(mf&&hasMpn(key,r)&&typeof r[cfg.price]==='number')return false;
    return !q||r.some(v=>v!=null&&String(v).toLowerCase().includes(q));
  });
  $('#cnt').textContent=`${out.length} of ${rows.length} parts`;
  $('#tb').innerHTML=out.map(({r,i:ri})=>`<tr class="click" tabindex="0" data-act="editPart" data-k="${key}" data-i="${ri}">${h.map((_,i)=>{
    const v=r[i];
    if(i===1)return `<td class="mono">${esc(v)}</td>`;
    if(key==='mqis'&&i===7)return `<td>${v?`<span class="proj proj-${projKey(v)}">${esc(v)}</span>`:''}</td>`;
    if(i===cfg.price)return `<td class="num">${typeof v==='number'?inr(v):'<span class="miss">No price</span>'}</td>`;
    if(cfg.mpn[0]===i&&!hasMpn(key,r))return `<td><span class="miss">No MPN</span></td>`;
    return `<td class="${cfg.num.includes(i)?'num':cfg.c.includes(i)?'c':''}">${esc(v)}</td>`}).join('')}</tr>`).join('');
}
function rawTable(rows,opt){
  rows=rows.filter(r=>r&&r.some(v=>v!=null&&v!==''));
  const width=Math.max(...rows.map(r=>{let n=r.length;while(n&&(r[n-1]==null||r[n-1]===''))n--;return n}));
  let html=`<div class="tablebox"><table class="cmp"><tbody>`;
  rows.slice(1).forEach(r=>{
    const filled=r.slice(0,width).filter(v=>v!=null&&v!=='').length,lab=String(r[0]||'');
    if(filled===1&&lab&&!opt.keep1.test(lab)){html+=`<tr class="sec"><td colspan="${width}">${esc(lab)}</td></tr>`;return}
    const isHead=opt.head.test(lab),isTot=/^(LANDED COST|TOTAL|Annual saving|Saving)/i.test(lab);
    html+=`<tr class="${isTot?'tot':''}">`+r.slice(0,width).map((v,i)=>isHead?`<th>${esc(v)}</th>`:`<td class="${typeof v==='number'?'num':''}">${opt.fmt(v,i,lab)}</td>`).join('')+'</tr>';
  });
  return html+'</tbody></table></div>';
}
const fanFmt=(v,i,lab)=>{if(typeof v!=='number')return esc(v);if(/%/.test(lab))return (v*100).toFixed(2)+'%';if(/Saving per unit|Annual saving/i.test(lab))return `<span class="save">${inr(v)}</span>`;if(/US\$|Exchange|volume|Lead time/i.test(lab))return v.toLocaleString('en-IN');return inr(v)};
V.fan=()=>{
  const d=S.data;
  let h=`<div class="tools"><p class="note" style="margin:0;flex:1">Landed cost per unit, existing (import) vs alternate supplier. Click ＋ to add a comparison for another part; landed cost and savings calculate automatically.</p><button class="btn primary" type="button" data-act="addFan">＋ Add comparison</button></div>`;
  const hasRaw=d.fan&&d.fan.filter(r=>r&&r.some(v=>v!=null&&v!=='')).length>1;
  if(hasRaw)h+=`<section><h2>1 · ${esc((d.fan.find(r=>r&&r[0])||[''])[0])}</h2>${rawTable(d.fan,{keep1:/^(Engineering|MOQ)/i,head:/^(Cost Element|Item)$/,fmt:fanFmt})}<p class="note" style="margin-top:8px">From the Excel sheet. To change it, update the Excel and use Load Excel.</p></section>`;
  else if(!d.fanExtra.length)h+='<div class="panel"><p class="sm">No comparisons yet. Click ＋ Add comparison, or Load Excel.</p></div>';
  d.fanExtra.forEach((f,i)=>{const c=fanCalc(f);
    const R=(l,a,b,n,cls)=>`<tr class="${cls||''}"><td>${l}</td><td class="num">${a}</td><td class="num">${b}</td><td>${n||''}</td></tr>`;
    h+=`<section><div class="h2row"><h2>${i+(hasRaw?2:1)} · ${esc(f.part)}</h2><button class="btn sm-btn" type="button" data-act="editFan" data-i="${i}">Edit</button></div><div class="tablebox"><table class="cmp"><tbody>
      <tr><th>Cost element</th><th>Existing supplier (import)${f.exName?' — '+esc(f.exName):''}</th><th>Alternate supplier (domestic)${f.altName?' — '+esc(f.altName):''}</th><th>Basis / note</th></tr>
      ${R('Unit price (US$)',f.pUSD??'','–')}${R('Exchange rate (₹ per US$)',f.fx??'','–')}${R('Unit price (₹)',inr(c.pINR),inr(f.altPrice))}
      ${R('Freight',inr(f.freight||0),inr(f.altFreight||0))}${R('Insurance',inr(c.ins),'–',(f.insPct||0)+'% of price + freight')}${R('CIF value',inr(c.cif),'–')}
      ${R('BCD',inr(c.bcd),'–',(f.bcdPct||0)+'%')}${R('SWS (on BCD)',inr(c.sws),'–',(f.swsPct||0)+'%')}${R('CHA / clearance / local freight',inr(f.cha||0),'–')}
      ${R('LANDED COST PER UNIT (₹)',inr(c.landed),inr(c.alt),'IGST excluded','tot')}
      <tr class="sec"><td colspan="4">Comparison</td></tr>
      ${R('Saving per unit with alternate',`<span class="${c.save>=0?'save':'loss'}">${inr(c.save)}</span>`,'')}${R('Saving %',(c.pct*100).toFixed(2)+'%','')}
      ${R('Annual volume (nos)',(f.vol||0).toLocaleString('en-IN'),'')}${R('Annual saving',`<span class="${c.annual>=0?'save':'loss'}">${inr(c.annual)}</span>`,'','','tot')}
      <tr class="sec"><td colspan="4">Before switching — confirm</td></tr>
      ${[['Make / MPN','exMpn','altMpn'],['Lead time (weeks)','exLt','altLt'],['MOQ','exMoq','altMoq'],['Warranty','exWar','altWar'],['Payment terms','exPay','altPay']].map(([l,a,b])=>`<tr><td>${l}</td><td>${esc(f[a])}</td><td>${esc(f[b])}</td><td></td></tr>`).join('')}
      ${f.note?`<tr><td>Note</td><td colspan="3">${esc(f.note)}</td></tr>`:''}
    </tbody></table></div></section>`});
  return h;
};
V.teknic=()=>{
  const t=S.data.tk;let te=0,tq=0,ts=0;
  const rows=t.parts.map((p,i)=>{const s=(p.exp||0)-(p.qp||0),pct=p.exp?s/p.exp:0,ss=s*(p.bom||0);te+=(p.exp||0)*(p.bom||0);tq+=(p.qp||0)*(p.bom||0);ts+=ss;
    return `<tr class="click" tabindex="0" data-act="editTk" data-i="${i}"><td class="mono">${esc(p.pn)}</td><td>${esc(p.name)}</td><td>${esc(p.make)}</td><td>${esc(p.mpn)}</td><td class="num">${esc(p.bom)}</td><td>${esc(p.exs)}</td><td>${esc(p.alts)}</td><td class="num">${inr(p.exp)}</td><td class="num">${inr(p.qp)}</td><td class="num"><span class="${s>=0?'save':'loss'}">${inr(s)}</span></td><td class="num">${(pct*100).toFixed(2)}%</td><td class="num"><span class="${ss>=0?'save':'loss'}">${inr(ss)}</span></td></tr>`}).join('');
  return `<div class="tools"><p class="note" style="margin:0;flex:1">${esc(t.title)}. Click a row to edit. Savings and totals calculate automatically.</p><button class="btn" type="button" data-act="editStations">Stations / year</button><button class="btn primary" type="button" data-act="addTk">＋ Add part</button></div>
  <div class="tablebox"><table><thead><tr><th>Part No.</th><th>Part Name</th><th>Make</th><th>MPN</th><th class="num">BOM / Stn</th><th>Existing supplier</th><th>Alternate supplier</th><th class="num">Existing ₹</th><th class="num">Quoted ₹</th><th class="num">Saving / unit</th><th class="num">Saving %</th><th class="num">Saving / stn</th></tr></thead>
  <tbody>${rows}<tr class="tot"><td colspan="7">TOTAL PER STATION</td><td class="num">${inr(te)}</td><td class="num">${inr(tq)}</td><td></td><td class="num">${te?((te-tq)/te*100).toFixed(2)+'%':''}</td><td class="num">${inr(ts)}</td></tr>
  <tr><td colspan="11">Stations per year</td><td class="num">${t.stations!=null?t.stations.toLocaleString('en-IN'):'—'}</td></tr>
  <tr class="tot"><td colspan="11">Annual saving</td><td class="num"><span class="save">${t.stations!=null?inr(ts*t.stations):'Enter stations'}</span></td></tr></tbody></table></div>`;
};

/* ---------- actions ---------- */
const STATUSES=['Ongoing','Completed','On Hold','Dropped'];
const stdActs=()=>{const p=S.data.projects.find(x=>x.name==='DLS')||S.data.projects.reduce((a,b)=>b.acts.length>a.acts.length?b:a,{acts:[]});return p.acts.length?p.acts.map(a=>a[1]):(S.data.stdActs||[])};
const A={
  addProject(){openForm({title:'Add project',fields:[
      {k:'name',label:'Project / part name',req:true,wide:true},{k:'status',label:'Status',type:'select',opts:STATUSES},
      {k:'start',label:'Start (trigger) date',type:'date'},{k:'std',type:'check',hint:'Add the standard stage-gate activity list ('+stdActs().length+' steps)',def:true,wide:true},
      {k:'remark',label:'Trigger remark',type:'textarea',wide:true}],
    submit:v=>{if(S.data.projects.some(p=>p.name.toLowerCase()===v.name.toLowerCase()))return 'A project with this name already exists.';
      const acts=v.std?stdActs().map((n,i)=>[i+1,n,null,null,null]):[[1,'Alternate change (trigger / start)',null,null,null]];
      acts[0][2]=v.start||null;acts[0][4]=v.remark||null;
      S.data.projects.push({name:v.name,status:v.status,acts});location.hash='projects-'+(S.data.projects.length-1);save('Project added')}})},
  editProject(e){const i=+e.dataset.p,p=S.data.projects[i];openForm({title:'Edit project',values:p,fields:[{k:'name',label:'Project / part name',req:true,wide:true},{k:'status',label:'Status',type:'select',opts:STATUSES}],
    submit:v=>{if(S.data.projects.some((x,j)=>j!==i&&x.name.toLowerCase()===v.name.toLowerCase()))return 'Another project already has this name.';p.name=v.name;p.status=v.status;save('Project updated')},
    del:()=>{S.data.projects.splice(i,1);location.hash='projects';save('Project deleted')}})},
  addAct(e){const p=S.data.projects[+e.dataset.p];openForm({title:'Add activity — '+p.name,fields:actFields(),
    submit:v=>{const n=Math.max(0,...p.acts.map(a=>typeof a[0]==='number'?a[0]:0))+1;p.acts.push([n,v.name,v.plan||null,v.actual||null,v.rem||null]);save('Activity added')}})},
  editAct(e){const p=S.data.projects[+e.dataset.p],j=+e.dataset.a,a=p.acts[j];
    openForm({title:'Update activity — '+p.name,values:{name:a[1],plan:a[2],actual:a[3],rem:a[4]},fields:actFields(),
      submit:v=>{a[1]=v.name;a[2]=v.plan||null;a[3]=v.actual||null;a[4]=v.rem||null;save('Activity updated')},del:()=>{p.acts.splice(j,1);save('Activity deleted')}})},
  addPart(e){partForm(e.dataset.k,null)},
  editPart(e){partForm(e.dataset.k,+e.dataset.i)},
  addFan(){fanForm(null)},editFan(e){fanForm(+e.dataset.i)},
  addTk(){tkForm(null)},editTk(e){tkForm(+e.dataset.i)},
  editStations(){openForm({title:'Stations per year',values:{s:S.data.tk.stations},fields:[{k:'s',label:'Annual station volume',type:'number'}],submit:v=>{S.data.tk.stations=v.s;save('Station volume updated')}})}
};
function actFields(){return [{k:'name',label:'Activity',req:true,wide:true},{k:'plan',label:'Plan date',type:'date'},{k:'actual',label:'Actual date',type:'date'},{k:'rem',label:'Remarks / delay reason',type:'textarea',wide:true}]}
function partForm(key,idx){
  const [h,rows]=S.data[key],cfg=PCFG[key],r=idx==null?null:rows[idx];
  const sugg=i=>[...new Set(rows.map(x=>x[i]).filter(x=>!blank(x)))].slice(0,40);
  const fields=h.map((lab,i)=>{
    if(i===0)return null;
    const f={k:'c'+i,label:lab,wide:/Part Name|Supplier|Vendor|MPN/i.test(lab)&&i!==1};
    if(i===1){f.req=true;f.label=lab+' (number)'}
    if(i===2)f.req=true;
    if(cfg.num.includes(i))f.type='number';
    if(key==='mqis'&&i===7)f.list=['Common','P1R5','Single Phs',...sugg(7)];
    else if(/UOM|Buyer|Supplier|Vendor|Rev/i.test(lab))f.list=sugg(i);
    if(cfg.mpn.includes(i))f.note='Format: MAKE/MPN';
    return f}).filter(Boolean);
  const values={};if(r)h.forEach((_,i)=>values['c'+i]=r[i]);
  else{values.c5=key==='mqis'?'Nos':'';if(key==='ds6')values.c5='Nos';}
  openForm({title:(r?'Edit part':'Add part')+' — '+(key==='mqis'?'MQIS & Single Phase':'6DS Project'),values,fields,
    submit:v=>{
      const pnRaw=String(v.c1).trim(),pn=/^\d+$/.test(pnRaw)?Number(pnRaw):pnRaw;
      const dup=rows.findIndex((x,j)=>j!==idx&&String(x[1]).trim()===String(pn));
      if(dup>=0)return `Part No ${pnRaw} already exists (row ${dup+1}: ${rows[dup][2]}). Edit that row instead.`;
      const row=h.map((_,i)=>i===0?null:i===1?pn:(v['c'+i]===''?null:v['c'+i]));
      if(r){row[0]=r[0];rows[idx]=row;save('Part updated')}
      else{row[0]=Math.max(0,...rows.map(x=>typeof x[0]==='number'?x[0]:0))+1;rows.push(row);save('Part added')}},
    del:r?()=>{rows.splice(idx,1);save('Part deleted')}:null});
}
function fanForm(idx){
  const f=idx==null?{fx:S.data.fan.find(r=>/Exchange/.test(String(r[0])))?.[1]??null,insPct:0.13,bcdPct:10,swsPct:10,freight:0,altFreight:0,cha:0}:S.data.fanExtra[idx];
  openForm({title:idx==null?'Add landed-cost comparison':'Edit comparison',values:f,fields:[
    {k:'part',label:'Part / description',req:true,wide:true},{k:'exName',label:'Existing supplier (import)'},{k:'altName',label:'Alternate supplier (domestic)'},
    {k:'pUSD',label:'Existing unit price (US$)',type:'number',req:true},{k:'fx',label:'Exchange rate (₹ per US$)',type:'number',req:true},
    {k:'freight',label:'Existing freight (₹/unit)',type:'number'},{k:'insPct',label:'Insurance %',type:'number'},
    {k:'bcdPct',label:'BCD %',type:'number'},{k:'swsPct',label:'SWS % (on BCD)',type:'number'},{k:'cha',label:'CHA / clearance (₹/unit)',type:'number'},
    {k:'altPrice',label:'Alternate unit price (₹)',type:'number',req:true},{k:'altFreight',label:'Alternate freight (₹/unit)',type:'number'},{k:'vol',label:'Annual volume (nos)',type:'number'},
    {k:'exMpn',label:'Existing make / MPN'},{k:'altMpn',label:'Alternate make / MPN'},{k:'exLt',label:'Existing lead time (wks)'},{k:'altLt',label:'Alternate lead time (wks)'},
    {k:'exMoq',label:'Existing MOQ'},{k:'altMoq',label:'Alternate MOQ'},{k:'exWar',label:'Existing warranty'},{k:'altWar',label:'Alternate warranty'},
    {k:'exPay',label:'Existing payment terms',wide:true},{k:'altPay',label:'Alternate payment terms',wide:true},{k:'note',label:'Note',type:'textarea',wide:true}],
    submit:v=>{if(S.data.fanExtra.some((x,j)=>j!==idx&&x.part.toLowerCase()===v.part.toLowerCase()))return 'A comparison for this part already exists.';
      if(idx==null){S.data.fanExtra.push(v);save('Comparison added')}else{S.data.fanExtra[idx]=v;save('Comparison updated')}},
    del:idx==null?null:()=>{S.data.fanExtra.splice(idx,1);save('Comparison deleted')}});
}
function tkForm(idx){
  const t=S.data.tk,p=idx==null?{bom:1}:t.parts[idx];
  openForm({title:idx==null?'Add part — Teknic (domestic)':'Edit part — Teknic (domestic)',values:p,fields:[
    {k:'pn',label:'Part No.',req:true},{k:'name',label:'Part name',req:true,wide:true},{k:'make',label:'Make',list:[...new Set(t.parts.map(x=>x.make).filter(Boolean))]},{k:'mpn',label:'MPN',wide:true},
    {k:'bom',label:'BOM / station',type:'number'},{k:'exs',label:'Existing supplier',list:[...new Set(t.parts.map(x=>x.exs).filter(Boolean))]},{k:'alts',label:'Alternate supplier',list:[...new Set(t.parts.map(x=>x.alts).filter(Boolean))]},
    {k:'exp',label:'Existing unit price (₹)',type:'number',req:true},{k:'qp',label:'Quoted unit price (₹)',type:'number',req:true}],
    submit:v=>{const pn=/^\d+$/.test(v.pn)?Number(v.pn):v.pn;if(t.parts.some((x,j)=>j!==idx&&String(x.pn)===String(pn)))return 'Part No '+v.pn+' is already in this table.';
      v.pn=pn;if(idx==null){t.parts.push(v);save('Part added')}else{t.parts[idx]=v;save('Part updated')}},
    del:idx==null?null:()=>{t.parts.splice(idx,1);save('Part deleted')}});
}

/* ---------- routing ---------- */
function render(){
  const id=(location.hash||'#overview').slice(1).split('-')[0];
  const cur=TABS.some(t=>t[0]===id)?id:'overview';
  $('#tabs').innerHTML=TABS.map(([k,l])=>`<button class="tab" role="tab" type="button" aria-selected="${k===cur}" data-tab="${k}">${l}</button>`).join('');
  const m=/^#projects-(\d+)$/.exec(location.hash);
  const keep={q:$('#q')?.value,pf:$('#pf')?.value,mf:$('#mf')?.checked};
  if(cur==='mqis'||cur==='ds6'){$('#view').innerHTML=partsView(cur);if(keep.q)$('#q').value=keep.q;if(keep.pf&&$('#pf'))$('#pf').value=keep.pf;if(keep.mf)$('#mf').checked=true;fillParts(cur);['#q','#pf','#mf'].forEach(s=>{const e=$(s);if(e)e.addEventListener('input',()=>fillParts(cur))})}
  else $('#view').innerHTML=V[cur](m?+m[1]:undefined);
  if(m&&render.scrollTo!==location.hash){render.scrollTo=location.hash;const el=document.getElementById('pj'+m[1]);if(el)el.scrollIntoView({block:'start'})}
  paintSave();
}
document.addEventListener('click',e=>{
  if(e.target.closest('#modal'))return;
  const t=e.target.closest('[data-tab]');if(t){location.hash=t.dataset.tab;return}
  const a=e.target.closest('[data-act]');if(a){e.preventDefault();A[a.dataset.act](a);return}
  const g=e.target.closest('[data-go]');if(g){location.hash='projects-'+g.dataset.go}
});
document.addEventListener('keydown',e=>{if(e.key==='Enter'){const a=e.target.closest&&e.target.closest('tr[data-act]');if(a)A[a.dataset.act](a)}});
window.addEventListener('hashchange',()=>{render.scrollTo=null;render()});

/* ---------- Excel load ---------- */
$('#loadBtn').onclick=()=>$('#file').click();
$('#file').onchange=async e=>{
  const f=e.target.files[0];if(!f)return;
  try{
    if(!window.XLSX)throw new Error('Excel reader did not load. Check your connection and reload the page.');
    const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true});
    const sh=(name,idx)=>wb.Sheets[name]||wb.Sheets[wb.SheetNames[idx]];
    const grid=ws=>XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true});
    const parts=(ws,hdrRe,ncol)=>{const g=grid(ws);const hi=g.findIndex(r=>r.some(v=>hdrRe.test(String(v||'').trim())));if(hi<0)throw new Error('Could not find the header row in a parts sheet.');
      const c0=g[hi].findIndex(v=>v!=null&&String(v).trim()!=='');const h=g[hi].slice(c0,c0+ncol).map(v=>String(v||'').replace(/ /g,' ').trim());
      return [h,g.slice(hi+1).map(r=>r.slice(c0,c0+ncol)).filter(r=>r.some(v=>v!=null&&v!==''))]};
    const nd={fanExtra:S.data.fanExtra};
    nd.mqis=parts(sh(SHEETS.mqis,0),/^Part No$/i,12);
    nd.ds6=parts(sh(SHEETS.ds6,1),/^Part No\.?$/i,10);
    const g=grid(sh(SHEETS.projects,2));const si=g.findIndex(r=>String(r[0]||'').trim()==='Stage');
    if(si<0)throw new Error('Could not find the "Stage" header in the Alternate Development sheet.');
    nd.projects=[];let cur=null;
    const fx=v=>v instanceof Date?new Date(v.getTime()-v.getTimezoneOffset()*6e4).toISOString().slice(0,10):v;
    g.slice(si+1).forEach(r=>{const a=r[0],b=r[1];
      if(typeof a==='string'&&a.trim()&&(b==null||b==='')){cur={name:a.trim(),status:String(r[4]||'').trim(),acts:[]};nd.projects.push(cur)}
      else if(cur&&b)cur.acts.push([a,b,fx(r[2]),fx(r[3]),r[5]])});
    const fws=sh(SHEETS.fan,3),tws=sh(SHEETS.teknic,4);
    nd.fan=fws?grid(fws).map(r=>r.slice(1,5)):S.data.fan;
    nd.teknic=tws?grid(tws).map(r=>r.slice(1,13)):S.data.teknic;
    nd.tk=tws?teknicFromRaw(nd.teknic):S.data.tk;
    S.data=nd;S.source=f.name;S.dirty=true;
    $('#srcline').textContent='Loaded '+f.name+' — not saved yet';render();
    $('#keepBtn').hidden=false;toast('Loaded '+f.name+'. Click "Save loaded data" to keep it.');
  }catch(err){toast('Could not read that file: '+err.message)}
  e.target.value='';
};
$('#keepBtn').onclick=()=>{$('#keepBtn').hidden=true;$('#srcline').textContent='E&E BOP · Alternate development · Sun Mobility CPMO';save('Excel data stored')};

/* ---------- Excel export ---------- */
$('#exportBtn').onclick=async()=>{
  try{
    if(!window.XLSX)throw new Error('Excel writer did not load.');
    const d=S.data,wb=XLSX.utils.book_new();
    const add=(name,aoa,cols)=>{const ws=XLSX.utils.aoa_to_sheet(aoa);if(cols)ws['!cols']=cols.map(w=>({wch:w}));XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31))};
    const st=d.projects.map(projStats);
    add('Project Status',[['#','Project Name','Status','Activities Done','Total','Progress %','Delayed']].concat(d.projects.map((p,i)=>[i+1,p.name,p.status,st[i].done,st[i].total,Math.round(st[i].pct*100),st[i].late.length])),[5,32,12,14,8,11,9]);
    add(SHEETS.mqis,[d.mqis[0]].concat(d.mqis[1]),[6,12,38,10,6,10,32,10,34,28,16,12]);
    add(SHEETS.ds6,[d.ds6[0]].concat(d.ds6[1]),[6,12,38,6,6,6,12,26,26,12]);
    const det=[['Project','#','Stage-Gate Activity','Plan Date','Actual Date','Status','Remarks']];
    d.projects.forEach(p=>p.acts.forEach(a=>det.push([p.name,a[0],a[1],a[2],a[3],p.status,a[4]])));
    add(SHEETS.projects,det,[26,5,46,12,12,12,70]);
    add(SHEETS.fan,d.fan,[38,26,28,60]);
    d.fanExtra.forEach((f,i)=>{const c=fanCalc(f);add('DC Fan '+(i+2)+' '+f.part.replace(/[\\\/?*\[\]:]/g,' '),[[f.part],['Cost element','Existing (import)','Alternate (domestic)'],['Supplier',f.exName,f.altName],['Unit price (US$)',f.pUSD],['Exchange rate',f.fx],['Unit price (₹)',c.pINR,f.altPrice],['Freight',f.freight||0,f.altFreight||0],['Insurance',c.ins],['CIF',c.cif],['BCD',c.bcd],['SWS',c.sws],['CHA',f.cha||0],['LANDED COST PER UNIT',c.landed,c.alt],['Saving per unit',c.save],['Saving %',c.pct],['Annual volume',f.vol],['Annual saving',c.annual],['Make / MPN',f.exMpn,f.altMpn],['Lead time (wks)',f.exLt,f.altLt],['MOQ',f.exMoq,f.altMoq],['Warranty',f.exWar,f.altWar],['Payment terms',f.exPay,f.altPay],['Note',f.note]],[34,22,22])});
    const t=d.tk,tk=[[t.title],['Part No.','Part Name','Make','MPN','BOM / Station','Existing Supplier','Alternate Supplier','Existing Unit Price (₹)','Quoted Unit Price (₹)','Saving / Unit (₹)','Saving %','Saving / Station (₹)']];
    let te=0,tq=0,ts=0;t.parts.forEach(p=>{const s=(p.exp||0)-(p.qp||0);te+=(p.exp||0)*(p.bom||0);tq+=(p.qp||0)*(p.bom||0);ts+=s*(p.bom||0);tk.push([p.pn,p.name,p.make,p.mpn,p.bom,p.exs,p.alts,p.exp,p.qp,s,p.exp?s/p.exp:0,s*(p.bom||0)])});
    tk.push(['TOTAL PER STATION',null,null,null,null,null,null,te,tq,null,te?(te-tq)/te:0,ts],[],['Stations per year',t.stations],['Annual saving',t.stations!=null?ts*t.stations:null]);
    add(SHEETS.teknic,tk,[14,28,10,30,8,16,14,14,14,12,9,14]);
    const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    download('NPD_Tracker_Export_'+new Date().toISOString().slice(0,10)+'.xlsx',new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    toast('Excel exported to your Downloads folder.');
  }catch(err){toast(err&&err.code==='declined'?'Export cancelled.':'Export failed: '+(err.message||err.code))}
};
$('#fileBtn').onclick=fileMenu;
$('#backupBtn').onclick=()=>{lastBackup=new Date().toISOString();persist();download('NPD_Tracker_Backup_'+lastBackup.slice(0,10)+'.json',new Blob([JSON.stringify({app:'NPD Sourcing Tracker',savedAt:lastBackup,data:S.data})],{type:'application/json'}));paintSave();toast('Backup saved to Downloads. Keep it on OneDrive or a shared drive.')};
$('#restoreBtn').onclick=()=>$('#rfile').click();
$('#rfile').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{const o=JSON.parse(await f.text());if(!o||!o.data||!o.data.projects)throw new Error('This is not a tracker backup file.');S.data=norm(o.data);S.source=f.name;save('Backup restored')}catch(err){toast('Could not restore: '+err.message)}e.target.value=''};
render();
initFile().then(()=>{if(!FSA)setTimeout(()=>toast('This browser cannot save to a file. Use Chrome or Edge, or take a Backup after each session.'),800);else if(!fh)setTimeout(()=>toast('First time: click "Link data file" (top right) and create NPD_Tracker_Data.json on OneDrive.'),800);else if(!fhOK)setTimeout(()=>toast('Click "Connect '+fh.name+'" (top right) to continue saving.'),800)});
})();
