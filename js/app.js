/* BI Analytics Platform v7 - robust import wizard */
const INITIAL_DATA = window.INITIAL_DATA || { analyses: [], records: [] };
const STORE_KEY = 'mtcit_bi_analytics_v7';
const STATE_KEY = STORE_KEY + '_state_v13';
const IDB_NAME = 'mtcit_bi_analytics_storage_v13';
const IDB_STORE = 'state';
const SEED_IDS = new Set((INITIAL_DATA.analyses||[]).map((a,i)=> a.id || ('seed_'+i+'_'+hashString([a.name,a.file,a.rows].join('|')))));
let DB = loadDB();
let charts = {};
let PENDING_IMPORT = null;

const PROJECT_TYPES = {
  survey: { label:'استطلاعات الرأي - إجادة/تجاوب', source:'استطلاعات', icon:'fa-square-poll-vertical', primary:'CSAT' },
  website: { label:'تقييم الموقع الإلكتروني/التطبيق', source:'تقييم الموقع', icon:'fa-globe', primary:'رضا الاستخدام' },
  training: { label:'تقييم الدورات التدريبية وورش العمل', source:'تقييم الدورات', icon:'fa-chalkboard-user', primary:'رضا المتدربين' },
  complaints: { label:'تحليل الشكاوى', source:'الشكاوى', icon:'fa-triangle-exclamation', primary:'نسبة الإغلاق' },
  inquiries: { label:'تحليل الاستفسارات', source:'الاستفسارات', icon:'fa-circle-question', primary:'نسبة الرد' },
  suggestions: { label:'تحليل المقترحات', source:'المقترحات', icon:'fa-lightbulb', primary:'نسبة القبول/التنفيذ' },
  reports: { label:'تحليل البلاغات', source:'البلاغات', icon:'fa-bullhorn', primary:'نسبة المعالجة' },
  generic: { label:'تحليل عام بربط أعمدة مرن', source:'تحليل عام', icon:'fa-table', primary:'المؤشر العام' }
};


function baseDB(){ return migrateDB(JSON.parse(JSON.stringify(INITIAL_DATA))); }
function loadDB(){
  const base = baseDB();
  try{
    const small = localStorage.getItem(STATE_KEY);
    if(small) return applyStoredState(base, JSON.parse(small));
  }catch(e){}
  // Backward compatibility: if an older full localStorage copy exists and is readable, use it once.
  try{
    const old = localStorage.getItem(STORE_KEY);
    if(old){
      const parsed = migrateDB(JSON.parse(old));
      if((old.length||0) < 4500000) return parsed;
    }
  }catch(e){}
  return base;
}
function migrateDB(db){
  db=db||{};
  db.analyses=Array.isArray(db.analyses)?db.analyses:[];
  db.records=Array.isArray(db.records)?db.records:[];
  db.deletedSeedIds=Array.isArray(db.deletedSeedIds)?db.deletedSeedIds:[];
  db.snapshots=Array.isArray(db.snapshots)?db.snapshots:[];
  db.analyses.forEach((a,i)=>{
    if(!a.id) a.id='seed_'+i+'_'+hashString([a.name,a.file,a.rows].join('|'));
    if(!a.projectType) a.projectType=(a.source==='إجادة'||a.source==='تجاوب')?'survey':'generic';
    a._seed = SEED_IDS.has(a.id);
  });
  db.records.forEach(r=>{
    if(!r.projectType) r.projectType=(r.source==='إجادة'||r.source==='تجاوب')?'survey':'generic';
    if(!r.fileId){ const a=db.analyses.find(x=>x.name===r.analysis && (!r.file || x.file===r.file)) || db.analyses.find(x=>x.name===r.analysis); if(a) r.fileId=a.id; }
    r._seed = SEED_IDS.has(r.fileId);
    if(r.score!==null&&r.score!==undefined&&r.score!=='') r.score=Number(r.score);
    if(r.scale!==null&&r.scale!==undefined&&r.scale!=='') r.scale=Number(r.scale);
  });
  return db;
}
function applyStoredState(base,state){
  state=state||{};
  const deleted = new Set(state.deletedSeedIds||[]);
  base.analyses = base.analyses.filter(a=>!deleted.has(a.id));
  base.records = base.records.filter(r=>!deleted.has(r.fileId));
  const userAnalyses = Array.isArray(state.userAnalyses)?state.userAnalyses:[];
  const userRecords = Array.isArray(state.userRecords)?state.userRecords:[];
  const existing = new Set(base.analyses.map(a=>a.id));
  userAnalyses.forEach(a=>{ if(a && a.id && !existing.has(a.id)){ a._seed=false; base.analyses.push(a); existing.add(a.id); } });
  userRecords.forEach(r=>{ if(r){ r._seed=false; base.records.push(r); } });
  base.deletedSeedIds = [...deleted];
  base.snapshots = Array.isArray(state.snapshots)?state.snapshots:[];
  base.lastSavedAt = state.lastSavedAt || base.lastSavedAt;
  return migrateDB(base);
}
function buildStoredState(){
  const deletedSeedIds = Array.isArray(DB.deletedSeedIds)?DB.deletedSeedIds:[];
  return {
    version:13,
    savedAt:new Date().toISOString(),
    lastSavedAt:DB.lastSavedAt || new Date().toISOString(),
    deletedSeedIds,
    // Save only imported/user data. The original 7 Excel files remain in data/initial-data.js and are not duplicated in browser storage.
    userAnalyses: DB.analyses.filter(a=>!SEED_IDS.has(a.id)).map(a=>({...a,_seed:false})),
    userRecords: DB.records.filter(r=>!SEED_IDS.has(r.fileId)).map(r=>({...r,_seed:false})),
    snapshots: DB.snapshots || []
  };
}
function openBIStore(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(IDB_NAME,1);
    req.onupgradeneeded=()=>{ const db=req.result; if(!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE); };
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
async function idbSet(key,val){ const db=await openBIStore(); return new Promise((resolve,reject)=>{ const tx=db.transaction(IDB_STORE,'readwrite'); tx.objectStore(IDB_STORE).put(val,key); tx.oncomplete=()=>{db.close(); resolve();}; tx.onerror=()=>{db.close(); reject(tx.error);}; }); }
async function idbGet(key){ const db=await openBIStore(); return new Promise((resolve,reject)=>{ const tx=db.transaction(IDB_STORE,'readonly'); const req=tx.objectStore(IDB_STORE).get(key); req.onsuccess=()=>{db.close(); resolve(req.result);}; req.onerror=()=>{db.close(); reject(req.error);}; }); }
async function loadIndexedState(){
  try{
    const state=await idbGet('dbState');
    if(state){ DB = applyStoredState(baseDB(), state); refreshAll(false); setSaveStatus('saved','تم تحميل البيانات المحفوظة'); }
  }catch(e){ console.warn('IndexedDB load failed', e); }
}
let persistTimer=null;
function persist(showMessage=false){
  try{
    DB.lastSavedAt = DB.lastSavedAt || new Date().toISOString();
    const state=buildStoredState();
    // Keep localStorage tiny to avoid quota errors, and save large imported data in IndexedDB.
    try{ localStorage.removeItem(STORE_KEY); }catch(_e){}
    try{
      const summary={version:state.version,lastSavedAt:state.lastSavedAt,savedAt:state.savedAt,deletedSeedIds:state.deletedSeedIds,snapshots:state.snapshots, userAnalyses: state.userAnalyses.map(a=>({...a, rows:a.rows||0})), userRecords: []};
      localStorage.setItem(STATE_KEY, JSON.stringify(summary));
    }catch(_e){}
    clearTimeout(persistTimer);
    persistTimer=setTimeout(()=>{
      idbSet('dbState', state).then(()=>{ setSaveStatus('saved','محفوظ'); if(showMessage) showToast('تم حفظ التحليل بنجاح في مساحة التخزين الموسعة.','success'); })
      .catch(e=>{ setSaveStatus('dirty','تعذر الحفظ'); showToast('تعذر حفظ البيانات في التخزين الموسع: '+(e?.message||e),'danger'); });
    },80);
  }catch(e){ showToast('تعذر تجهيز البيانات للحفظ: '+(e?.message||e),'danger'); setSaveStatus('dirty','تعذر الحفظ'); }
}

function fmt(n,d=0){ return (Number(n)||0).toLocaleString('ar-OM',{maximumFractionDigits:d,minimumFractionDigits:d}); }
function pct(n){ return fmt(n,1)+'%'; }
function esc(s){ return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m])); }
function hashString(str){ let h=2166136261; for(let i=0;i<String(str).length;i++){ h^=String(str).charCodeAt(i); h=Math.imul(h,16777619); } return (h>>>0).toString(36); }
async function bufferHash(buf){ const arr=new Uint8Array(buf); let s=''; for(let i=0;i<arr.length;i+=Math.max(1,Math.floor(arr.length/5000))) s+=String.fromCharCode(arr[i]); return hashString(arr.length+'|'+s); }
function contentHash(rows){ return hashString(JSON.stringify(rows.slice(0,2000).map(r=>[r.service,r.date,r.score,r.comment,r.status,r.channel]))+'|'+rows.length); }
function groupBy(arr,fn){ return arr.reduce((m,x)=>{ const k=typeof fn==='function'?fn(x):x[fn]; (m[k||'غير محدد']||(m[k||'غير محدد']=[])).push(x); return m; },{}); }
function avg(arr,fn){ const vals=arr.map(fn).filter(v=>v!==null&&v!==undefined&&!isNaN(v)); return vals.length?vals.reduce((a,b)=>a+Number(b),0)/vals.length:0; }
function uniq(arr){ return [...new Set(arr.filter(Boolean))]; }
function getProjectTypeLabel(t){ return (PROJECT_TYPES[t]||PROJECT_TYPES.generic).label; }
function getProjectSource(t){ return (PROJECT_TYPES[t]||PROJECT_TYPES.generic).source; }

function showToast(msg,type='success'){
  let panel=document.getElementById('toastPanel');
  if(!panel){ panel=document.createElement('div'); panel.id='toastPanel'; panel.className='toast-panel'; document.body.appendChild(panel); }
  const el=document.createElement('div'); el.className='toast '+(type==='warning'?'warn':type==='danger'?'danger':'ok'); el.innerHTML=msg; panel.appendChild(el); setTimeout(()=>{el.remove();},6500);
}

function setupNav(){ document.querySelectorAll('.navbtn').forEach(b=>b.onclick=()=>{ document.querySelectorAll('.navbtn').forEach(x=>x.classList.remove('active')); b.classList.add('active'); document.querySelectorAll('.section').forEach(s=>s.classList.remove('active')); document.getElementById('tab-'+b.dataset.tab)?.classList.add('active'); if(b.dataset.tab==='report') renderReport(); if(b.dataset.tab==='imports') updateImportModeUI(); }); }
function sessionNames(){ return uniq(DB.analyses.map(a=>a.name)); }
function fillFilters(){
  const names=sessionNames();
  const af=document.getElementById('analysisFilter'); if(af){ const cur=af.value||'all'; af.innerHTML='<option value="all">كل الجلسات</option>'+names.map(n=>`<option>${esc(n)}</option>`).join(''); af.value=[...af.options].some(o=>o.value===cur)?cur:'all'; }
  const app=document.getElementById('appendTargetAnalysis'); if(app){ const cur=app.value; app.innerHTML=names.map(n=>`<option>${esc(n)}</option>`).join(''); app.value=names.includes(cur)?cur:(names[0]||''); }
  const ptf=document.getElementById('projectTypeFilter'); if(ptf){ const cur=ptf.value||'all'; const types=uniq(DB.analyses.map(a=>a.projectType)); ptf.innerHTML='<option value="all">كل الأنواع</option>'+Object.keys(PROJECT_TYPES).filter(k=>types.includes(k)).map(k=>`<option value="${k}">${PROJECT_TYPES[k].label}</option>`).join(''); ptf.value=[...ptf.options].some(o=>o.value===cur)?cur:'all'; }
  const side=document.getElementById('sideFiles'); if(side){ side.innerHTML=DB.analyses.slice(-12).reverse().map(a=>`• <b>${esc(a.name)}</b><br><span style="margin-right:10px">${esc(a.file)}</span>`).join('<br>') || 'لا توجد ملفات محفوظة'; }
  renderReportFilters();
}
function updateImportModeUI(){ const mode=document.getElementById('importMode')?.value||'new'; const nb=document.getElementById('newAnalysisBox'), ab=document.getElementById('appendAnalysisBox'); if(nb)nb.style.display=mode==='new'?'block':'none'; if(ab)ab.style.display=mode==='append'?'block':'none'; fillFilters(); updateTemplateHint(); }
function updateTemplateHint(){
  const type=document.getElementById('projectType')?.value||'survey'; const el=document.getElementById('templateHint'); if(!el)return;
  const examples={ survey:'يمكن ربط أعمدة الخدمة/السؤال/التقييم/الملاحظة. ملفات إجادة وتجاوب تُقبل أيضاً كاستطلاعات عامة عند اختلاف الفورمات.', website:'مثال: الصفحة، الخدمة الرقمية، سهولة الاستخدام، سرعة الموقع، التقييم العام، الملاحظات.', training:'مثال: اسم الدورة، المدرب، رضا المدرب، جودة المحتوى، التنظيم، التقييم العام، الملاحظات.', complaints:'مثال: رقم الشكوى، الخدمة، الحالة، الأولوية، تاريخ الإنشاء، تاريخ الإغلاق، وصف الشكوى.', inquiries:'مثال: موضوع الاستفسار، القناة، الحالة، تاريخ الاستفسار، نص الاستفسار.', suggestions:'مثال: المقترح، الجهة، الحالة، التاريخ، تفاصيل المقترح، قرار اللجنة.', reports:'مثال: نوع البلاغ، الحالة، الخطورة، الموقع/القناة، التاريخ، وصف البلاغ.', generic:'أي ملف عام: اختر الأعمدة المناسبة من شاشة المعاينة قبل الحفظ.' };
  el.innerHTML=`<b>${esc(getProjectTypeLabel(type))}</b><br>${esc(examples[type]||examples.generic)}<br><span class="small">V7: لا يلزم كتابة أسماء الأعمدة مسبقاً؛ اختر الملف وستظهر كل الأعمدة في شاشة الربط قبل الحفظ.</span>`;
}

function getFiltered(){ const af=document.getElementById('analysisFilter')?.value||'all'; const sf=document.getElementById('sourceFilter')?.value||'all'; return DB.records.filter(r=>(af==='all'||r.analysis===af)&&(sf==='all'||r.source===sf)); }
function csat(arr){ let ok=0,total=0; arr.forEach(r=>{ if(r.score==null || isNaN(r.score)) return; total++; const sc=Number(r.score), scale=Number(r.scale||5); if((scale<=5&&sc>=4)||(scale>5&&sc>=Math.ceil(scale*0.7))) ok++; }); return total?ok/total*100:0; }
function npsScore(arr){ const vals=arr.map(r=>r.nps).filter(v=>v!==null&&v!==undefined&&!isNaN(v)); if(!vals.length)return null; const p=vals.filter(x=>x>=9).length, d=vals.filter(x=>x<=6).length; return (p-d)/vals.length*100; }
function serviceStats(records){ return Object.entries(groupBy(records,'service')).map(([service,rows])=>({ service, source:uniq(rows.map(r=>r.source)).join(' / '), count:rows.length, avg:avg(rows,r=>r.score!=null?(Number(r.score)/(Number(r.scale)||5)*100):null), csat:csat(rows), nps:npsScore(rows) })).sort((a,b)=>b.count-a.count); }
function serviceStatsBySource(records){ return Object.entries(groupBy(records,r=>(r.source||'غير محدد')+'|||'+(r.service||'غير محدد'))).map(([key,rows])=>{ const [source,...rest]=key.split('|||'); return { source, service:rest.join('|||'), count:rows.length, avg:avg(rows,r=>r.score!=null?(Number(r.score)/(Number(r.scale)||5)*100):null), csat:csat(rows), nps:npsScore(rows) }; }).sort((a,b)=>b.count-a.count); }
function sourceList(records){ return uniq(records.map(r=>r.source||'غير محدد')); }
function qualityClass(v){ if(v>=85)return ['b-good','متميز']; if(v>=70)return ['b-mid','جيد']; return ['b-bad','يحتاج تحسين']; }
function isNoValueComment(c){ const t=String(c||'').trim().toLowerCase().replace(/[؟?!.،,؛;]+/g,' '); return !t || t==='.' || t==='-' || /(لا\s*يوجد|لايوجد|لا\s*توجد|لا\s*شيء|لا\s*ملاحظات|لا\s*مقترحات|nil|none|no\s*comments?|na|n\/a)/i.test(t); }

function renderKPIs(){
  const rec=getFiltered(), srcs=sourceList(rec); const target=document.getElementById('kpiCards'); if(!target)return;
  if(!rec.length){ target.innerHTML='<div class="empty-state">لا توجد بيانات في الفلتر الحالي.</div>'; document.getElementById('sourceCompareRows').innerHTML=''; return; }
  function card(icon,val,lbl,color,tip=''){ return `<div class="kpi" title="${esc(tip)}"><div class="ico" style="background:${color}"><i class="fa-solid ${icon}"></i></div><div><div class="val">${val}</div><div class="lbl">${lbl}</div></div></div>`; }
  let html='<div class="source-grid">';
  srcs.forEach(src=>{ const rows=rec.filter(r=>r.source===src); const pt=uniq(rows.map(r=>r.projectType))[0]||'generic'; const color=src==='تجاوب'?'var(--green)':src==='إجادة'?'var(--blue)':'var(--cyan)'; const comments=rows.filter(r=>!isNoValueComment(r.comment)).length; const low=serviceStats(rows).filter(s=>s.csat<70).length; html+=`<div class="source-panel ${src==='تجاوب'?'tajawb':'egada'}"><div class="source-title"><div><h3>${esc(src)}</h3><div class="source-note">${esc(getProjectTypeLabel(pt))}</div></div><span class="tag ${src==='تجاوب'?'tag-green':'tag-blue'}">${fmt(rows.length)} سجل</span></div><div class="mini-kpis">${card('fa-users',fmt(rows.length),'حجم التفاعل',color)}${card('fa-face-smile',pct(csat(rows)),'CSAT - رضا المستفيد','var(--green)','Customer Satisfaction Score')}${card('fa-bullhorn',npsScore(rows)==null?'—':fmt(npsScore(rows),1),'NPS - صافي المروجين','var(--cyan)','Net Promoter Score')}${card('fa-briefcase',fmt(new Set(rows.map(r=>r.service)).size),'الخدمات/المواضيع','var(--amber)')}${card('fa-comments',fmt(comments),'ملاحظات مفيدة','var(--navy)')}${card('fa-triangle-exclamation',fmt(low),'أولوية التحسين','var(--red)')}</div></div>`; });
  html+='</div>'; target.innerHTML=html;
  const tbody=document.getElementById('sourceCompareRows'); if(tbody){ tbody.innerHTML=srcs.map(src=>{ const rows=rec.filter(r=>r.source===src); return `<tr><td>${esc(src)}</td><td colspan="2">${fmt(rows.length)} سجل | ${fmt(new Set(rows.map(r=>r.service)).size)} خدمة/موضوع | CSAT ${pct(csat(rows))} | NPS ${npsScore(rows)==null?'—':fmt(npsScore(rows),1)}</td><td>${esc(getProjectTypeLabel(uniq(rows.map(r=>r.projectType))[0]||'generic'))}</td></tr>`; }).join(''); }
}
function makeChart(id,type,data,options={}){ const c=document.getElementById(id); if(!c || !window.Chart)return; if(charts[id])charts[id].destroy(); charts[id]=new Chart(c,{type,data,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom',labels:{font:{family:'Cairo'}}}},...options}}); }
function renderCharts(){ const rec=getFiltered(); if(!rec.length)return; const by=groupBy(rec,r=>(r.date||'غير محدد').slice(0,7)); const labels=Object.keys(by).sort(); const srcs=sourceList(rec); const colors=['#0b5cab','#16a34a','#00a6b2','#f59e0b','#8b5cf6','#dc2626']; makeChart('csatChart','bar',{labels,datasets:srcs.map((src,i)=>({label:src,data:labels.map(k=>csat((by[k]||[]).filter(r=>r.source===src))),backgroundColor:colors[i%colors.length]}))},{scales:{y:{beginAtZero:true,max:100}}}); makeChart('npsChart','bar',{labels:srcs,datasets:[{label:'NPS',data:srcs.map(src=>npsScore(rec.filter(r=>r.source===src))||0),backgroundColor:srcs.map((_,i)=>colors[i%colors.length])}]},{scales:{y:{beginAtZero:true}}}); }
function priorityAction(s){ if(s.csat<60)return 'تحسين عاجل لتجربة المستفيد، ومراجعة خطوات الخدمة وزمن الإنجاز وقنوات الدعم.'; if(s.csat<75)return 'تحليل أسباب انخفاض الرضا ووضع خطة تحسين قصيرة المدى للخدمة.'; return 'متابعة دورية ومعالجة الملاحظات المتكررة للحفاظ على التحسن.'; }
function renderTopLow(){ const rec=getFiltered(); const top=document.getElementById('topServices'), low=document.getElementById('lowServices'); if(!top||!low)return; const stats=serviceStatsBySource(rec); const topRows=stats.filter(s=>s.count>0).sort((a,b)=>b.csat-a.csat).slice(0,10); top.innerHTML=topRows.map(s=>`<tr><td>${esc(s.service)}</td><td>${esc(s.source)}</td><td>${fmt(s.count)}</td><td>${pct(s.csat)}</td><td>${s.nps==null?'—':fmt(s.nps,1)}</td></tr>`).join('')||'<tr><td colspan="5">لا توجد بيانات</td></tr>'; const lows=stats.filter(s=>s.count>0).sort((a,b)=>a.csat-b.csat).slice(0,8); low.innerHTML=lows.map(s=>`<div class="priority-card"><div><b>${esc(s.service)}</b><div class="small">${esc(s.source)} | ${fmt(s.count)} سجل</div></div><div class="priority-metrics"><span>CSAT: <b>${pct(s.csat)}</b></span><span>NPS: <b>${s.nps==null?'—':fmt(s.nps,1)}</b></span></div><div class="priority-action"><b>الإجراء المقترح:</b> ${esc(priorityAction(s))}</div></div>`).join('')||'<div class="empty-state">لا توجد بيانات كافية.</div>'; }
function renderServices(){ const q=(document.getElementById('serviceSearch')?.value||'').trim().toLowerCase(); const rows=serviceStatsBySource(getFiltered()).filter(s=>!q||s.service.toLowerCase().includes(q)||s.source.toLowerCase().includes(q)); const tb=document.getElementById('servicesBody'); if(!tb)return; tb.innerHTML=rows.map((s,i)=>{ const [cls,label]=qualityClass(s.csat); return `<tr><td>${i+1}</td><td>${esc(s.service)}</td><td>${esc(s.source)}</td><td>${fmt(s.count)}</td><td>${pct(s.avg)}</td><td>${pct(s.csat)}</td><td>${s.nps==null?'—':fmt(s.nps,1)}</td><td><span class="badge ${cls}">${label}</span></td></tr>`; }).join('')||'<tr><td colspan="8">لا توجد بيانات</td></tr>'; }

function uniqueComments(rows){
  const seen=new Set(), out=[];
  rows.forEach(r=>{ if(isNoValueComment(r.comment))return; const key=(String(r.comment||'').trim().replace(/\s+/g,' ')+'|||'+(r.service||'')).toLowerCase(); if(seen.has(key))return; seen.add(key); out.push(r); });
  return out;
}
function isNegativeComment(t){ return /(تأخير|بطي|صعب|مشكلة|شكوى|تحسين|زمن|رفض|صفر|سيئ|ضعيف|تعطل|انتظار|تطوير|slow|bad|issue|problem|complaint|delay)/i.test(t||''); }
function isPositiveComment(t){ return /(شكر|ممتاز|رائع|سعيد|تقدير|ابتسام|تعاون|راقي|good|excellent|perfect|thanks|thank)/i.test(t||''); }
function renderQual(){
  const rec=getFiltered(); const comments=uniqueComments(rec); const neg=comments.filter(r=>isNegativeComment(r.comment)); const pos=comments.filter(r=>isPositiveComment(r.comment)&&!isNegativeComment(r.comment)); const neutral=comments.filter(r=>!isPositiveComment(r.comment)&&!isNegativeComment(r.comment));
  const cards=document.getElementById('qualCards'); if(cards)cards.innerHTML=`<div class="kpi"><div class="ico"><i class="fa-solid fa-comments"></i></div><div><div class="val">${fmt(comments.length)}</div><div class="lbl">ملاحظات مفيدة غير مكررة</div></div></div><div class="kpi"><div class="ico" style="background:var(--green)"><i class="fa-solid fa-heart"></i></div><div><div class="val">${fmt(pos.length)}</div><div class="lbl">ملاحظات الثناء والتجارب الإيجابية</div></div></div><div class="kpi"><div class="ico" style="background:var(--red)"><i class="fa-solid fa-tools"></i></div><div><div class="val">${fmt(neg.length)}</div><div class="lbl">ملاحظات التحسين والشكاوى</div></div></div>`;
  const pc=document.getElementById('positiveComments'), nc=document.getElementById('negativeComments');
  const card=(r,cls='')=>`<div class="comment ${cls}"><b>${esc(r.service)}</b><div class="small">${esc(r.source||'')} | ${esc(r.date||'')}</div><p>${esc(r.comment)}</p></div>`;
  if(pc)pc.innerHTML=pos.slice(0,40).map(r=>card(r,'positive')).join('') || (neutral.length?neutral.slice(0,10).map(r=>card(r,'neutral')).join(''):'<div class="empty-state">لا توجد ملاحظات ثناء واضحة بعد إزالة التكرار. قد تكون بعض الردود محايدة أو مكتوبة بشكل غير واضح.</div>');
  if(nc)nc.innerHTML=neg.slice(0,40).map(r=>card(r,'negative')).join('') || '<div class="empty-state">لا توجد ملاحظات تحسين واضحة بعد استبعاد “لا يوجد” وإزالة التكرار.</div>';
}
function renderImports(){ const tb=document.getElementById('analysisRows'); if(!tb)return; tb.innerHTML=DB.analyses.map(a=>`<tr><td>${esc(a.name)}<div class="small">${esc(getProjectTypeLabel(a.projectType||'generic'))}</div></td><td><span class="badge b-info">${esc(a.source)}</span></td><td>${esc(a.date)}</td><td>${esc(a.file)}</td><td>${fmt(a.rows)}</td><td><button class="btn btn-red btn-mini" onclick="deleteImportedAnalysis('${esc(a.id)}')"><i class="fa-solid fa-trash"></i> حذف الملف ومحتوياته</button></td></tr>`).join('')||'<tr><td colspan="6">لا توجد ملفات محفوظة.</td></tr>'; }
function deleteImportedAnalysis(id){ const a=DB.analyses.find(x=>x.id===id); if(!a)return; if(!confirm('سيتم حذف الملف ومحتوياته وإعادة احتساب النتائج. هل تريد المتابعة؟'))return; if(SEED_IDS.has(id)){ DB.deletedSeedIds=Array.isArray(DB.deletedSeedIds)?DB.deletedSeedIds:[]; if(!DB.deletedSeedIds.includes(id)) DB.deletedSeedIds.push(id); } DB.analyses=DB.analyses.filter(x=>x.id!==id); DB.records=DB.records.filter(r=>r.fileId!==id); refreshAll(); showToast('تم حذف الملف ومحتوياته وتحديث المؤشرات.','warning'); }
function renderSources(){ const tb=document.getElementById('sourcesBody'); if(!tb)return; tb.innerHTML=DB.analyses.map(a=>`<tr><td>${esc(a.file)}</td><td>${esc(getProjectTypeLabel(a.projectType||'generic'))}</td><td>${esc(a.date)}</td><td>${fmt(a.rows)}</td><td>${esc(a.name)}</td></tr>`).join('')||'<tr><td colspan="5">لا توجد مصادر</td></tr>'; }


function projectStats(projectType, rows){
  const total=rows.length;
  const closed=rows.filter(r=>/(مغلق|منجز|مكتمل|تم|closed|done|resolved|accepted|approved|replied|answered)/i.test(r.status||'')).length;
  const high=rows.filter(r=>/(عاجل|حرج|مرتفع|high|critical|urgent)/i.test((r.priority||'')+' '+(r.status||''))).length;
  const usefulComments=rows.filter(r=>!isNoValueComment(r.comment)).length;
  const files=new Set(rows.map(r=>r.fileId||r.file).filter(Boolean)).size;
  const sessions=new Set(rows.map(r=>r.analysis).filter(Boolean)).size;
  const cs=csat(rows), ns=npsScore(rows), closedRate=total?closed/total*100:0;
  return { total, services:new Set(rows.map(r=>r.service)).size, cs, ns, closedRate, high, usefulComments, files, sessions, avgScore:avg(rows,r=>r.score!=null?(Number(r.score)/(Number(r.scale)||5)*100):null) };
}
function isOpsType(t){ return ['complaints','reports','inquiries','suggestions'].includes(t); }
function primaryMetricLabel(t){ return isOpsType(t)?'نسبة المعالجة/الإغلاق':'مؤشر الرضا CSAT'; }
function primaryMetricValue(t,st){ return isOpsType(t)?pct(st.closedRate):pct(st.cs); }
function projectInsight(t,st){
  if(!st.total) return 'لا توجد بيانات محفوظة لهذا النوع بعد.';
  if(isOpsType(t)) return `تم رصد ${fmt(st.high)} حالة عالية الأولوية، مع ${fmt(st.usefulComments)} وصف/ملاحظة قابلة للتحليل.`;
  const nps=st.ns==null?'لا توجد بيانات كافية لـ NPS':`NPS ${fmt(st.ns,1)}`;
  return `${nps}، و${fmt(st.usefulComments)} ملاحظة مفيدة لتحليل تجربة المستفيد.`;
}
function renderProjectCards(){
  const box=document.getElementById('analysisTypeCards'); if(!box)return;
  const groups=groupBy(DB.records,'projectType');
  box.innerHTML=Object.keys(PROJECT_TYPES).map(k=>{
    const rows=groups[k]||[]; const st=projectStats(k,rows);
    const hasData = st.total>0;
    const metric = primaryMetricValue(k,st);
    const metricLabel = primaryMetricLabel(k);
    const insight = projectInsight(k,st);
    const subtitle = PROJECT_TYPES[k].primary || metricLabel;
    return `<article class="analysis-type-card ${k} ${hasData?'has-data':'no-data'}" onclick="setProjectTypeFilter('${k}')">
      <div class="type-head">
        <div class="type-icon" aria-hidden="true"><i class="fa-solid ${PROJECT_TYPES[k].icon}"></i></div>
        <div class="type-title">
          <h3>${esc(PROJECT_TYPES[k].label)}</h3>
          <p>${esc(subtitle)}</p>
        </div>
      </div>
      <div class="type-metrics">
        <div class="type-metric primary"><span>الجلسات</span><b>${fmt(st.sessions)}</b><small>عدد جلسات التحليل</small></div>
        <div class="type-metric"><span>السجلات</span><b>${fmt(st.total)}</b><small>إجمالي الصفوف المحللة</small></div>
        <div class="type-metric"><span>${esc(metricLabel)}</span><b>${metric}</b><small>${isOpsType(k)?'نسبة الحالات المعالجة':'نسبة الرضا المحسوبة'}</small></div>
      </div>
      <div class="type-footer ${hasData?'':'muted'}">
        <i class="fa-solid ${hasData?'fa-circle-info':'fa-database'}"></i>
        <span>${esc(insight)}</span>
      </div>
    </article>`;
  }).join('');
}
function setProjectTypeFilter(k){ const sel=document.getElementById('projectTypeFilter'); if(sel){ sel.value=k; renderProjectDashboard(); } }
function renderProjectSessions(){
  const tb=document.getElementById('projectSessionsRows'); if(!tb)return;
  const grouped=groupBy(DB.analyses,a=>(a.projectType||'generic')+'|||'+a.name);
  tb.innerHTML=Object.entries(grouped).map(([k,files])=>{
    const [pt,name]=k.split('|||'); const rows=DB.records.filter(r=>r.analysis===name && r.projectType===pt); const st=projectStats(pt,rows);
    return `<tr><td>${esc(getProjectTypeLabel(pt))}</td><td><b>${esc(name)}</b><div class="small">${esc(files.map(f=>f.date).filter(Boolean).slice(-1)[0]||'بدون تاريخ')}</div></td><td>${fmt(files.length)}<span class="metric-help">عدد ملفات Excel داخل الجلسة</span></td><td>${fmt(rows.length)}<span class="metric-help">إجمالي السجلات بعد الدمج</span></td><td><b>${primaryMetricValue(pt,st)}</b><span class="metric-help">${primaryMetricLabel(pt)}</span></td></tr>`;
  }).join('')||'<tr><td colspan="5">لا توجد جلسات</td></tr>';
}
function renderProjectDashboard(){
  const t=document.getElementById('projectTypeFilter')?.value||'all'; const rows=DB.records.filter(r=>t==='all'||r.projectType===t); const box=document.getElementById('projectDashboard'); if(!box)return;
  const type=t==='all'?'generic':t; const st=projectStats(type,rows); const top=serviceStatsBySource(rows).slice(0,5); const bySession=Object.entries(groupBy(rows,'analysis')).map(([name,rs])=>({name,count:rs.length,cs:csat(rs)})).sort((a,b)=>b.count-a.count).slice(0,6);
  box.innerHTML=`<div class="project-kpi-grid"><div class="project-kpi"><b>${fmt(st.total)}</b><span>إجمالي السجلات</span><em>عدد الصفوف المحللة في النوع المحدد</em></div><div class="project-kpi"><b>${fmt(st.sessions)}</b><span>جلسات نشطة</span><em>عدد جلسات التحليل المرتبطة بهذا النوع</em></div><div class="project-kpi"><b>${fmt(st.services)}</b><span>خدمات/مواضيع</span><em>عدد الموضوعات أو الخدمات المميزة</em></div><div class="project-kpi"><b>${primaryMetricValue(type,st)}</b><span>${primaryMetricLabel(type)}</span><em>${isOpsType(type)?'نسبة السجلات المغلقة أو المعالجة':'نسبة الرضا المحسوبة من التقييمات'}</em></div><div class="project-kpi"><b>${fmt(st.usefulComments)}</b><span>ملاحظات مفيدة</span><em>بعد حذف العبارات غير المفيدة والتكرارات الأولية</em></div><div class="project-kpi"><b>${fmt(st.high)}</b><span>أولوية عالية</span><em>حالات حرجة أو عاجلة حسب الحالة/الأولوية</em></div></div><div class="generic-chart"><h3>أكبر الجلسات حسب حجم البيانات</h3>${bySession.map(x=>`<div class="generic-row"><span title="${esc(x.name)}">${esc(x.name)}</span><div class="track"><div class="fill" style="width:${Math.min(100,(x.count/Math.max(...bySession.map(b=>b.count),1))*100)}%"></div></div><b>${fmt(x.count)}</b></div>`).join('')||'<div class="empty-state">لا توجد جلسات لهذا النوع.</div>'}</div><div class="generic-chart"><h3>أبرز الخدمات/المواضيع</h3>${top.map(x=>`<div class="generic-row"><span title="${esc(x.service)}">${esc(x.service)}</span><div class="track"><div class="fill" style="width:${Math.min(100,x.csat)}%"></div></div><b>${pct(x.csat)}</b></div>`).join('')||'<div class="empty-state">لا توجد خدمات كافية.</div>'}</div>`;
}

// Import Wizard V7
function findHeaderRow(rows){ let best=0,score=-1; for(let i=0;i<Math.min(rows.length,20);i++){ const vals=(rows[i]||[]).map(x=>String(x??'').trim()).filter(Boolean); const sc=vals.length + vals.filter(v=>/[ء-يa-zA-Z]/.test(v)).length; if(sc>score){score=sc;best=i;} } return best; }
function normalizeHeader(h){ return String(h||'').trim().replace(/\s+/g,' '); }
function findColumn(headers,patterns){ const hs=headers.map(h=>normalizeHeader(h).toLowerCase()); for(const p of patterns){ const idx=hs.findIndex(h=>h.includes(p.toLowerCase())); if(idx>=0)return headers[idx]; } return ''; }
function detectMapping(headers,projectType){
  return {
    service: findColumn(headers, projectType==='training'?['الدورة','اسم الدورة','البرنامج','الورشة','الخدمة','الموضوع']:projectType==='website'?['الصفحة','الموقع','الخدمة','الرابط','الموضوع']:['الخدمة','الموضوع','التقسيم','الدائرة','الإدارة','اسم الخدمة','البلاغ','الشكوى','المقترح','الاستفسار']),
    score: findColumn(headers, ['التقييم العام','التقييم','الرضا','الدرجة','score','rating','راضي','سهولة','جودة','مدى رضاك']),
    comment: findColumn(headers, ['الملاحظات','ملاحظات','تعليق','التعليق','الوصف','نص','الشكوى','المقترح','الاستفسار','comment','description']),
    status: findColumn(headers, ['الحالة','status','وضع','الإجراء','نتيجة']),
    date: findColumn(headers, ['التاريخ','تاريخ','date','created','وقت','اليوم']),
    channel: findColumn(headers, ['القناة','الجهة','الإدارة','الدائرة','المحافظة','channel','department']),
    priority: findColumn(headers, ['الأولوية','الخطورة','priority','severity','حرج'])
  };
}
function excelDateToISO(v){ if(v instanceof Date) return v.toISOString().slice(0,10); const s=String(v||'').trim(); if(/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) return s.slice(0,10); if(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(s)) return s; return s; }
function parseScore(val,status,projectType){
  const t=String(val??'').trim(); const st=String(status??'').trim();
  if(t!=='' && !isNaN(Number(t))){ let n=Number(t); let scale=n<=5?5:n<=7?7:10; return {score:n,scale,nps:scale===10?n:Math.round(n/scale*10)}; }
  const text=(t+' '+st).toLowerCase();
  if(/راضي جدا|ممتاز|excellent|very good|مكتمل|مغلق|منجز|تم|accepted|approved|closed|resolved/.test(text)) return {score:5,scale:5,nps:10};
  if(/راضي|جيد|good|تحت الدراسة|قيد المعالجة|in progress/.test(text)) return {score:4,scale:5,nps:8};
  if(/محايد|متوسط|neutral|medium/.test(text)) return {score:3,scale:5,nps:6};
  if(/غير راضي|ضعيف|سيئ|مرفوض|مفتوح|متأخر|rejected|open|bad|poor/.test(text)) return {score:2,scale:5,nps:3};
  if(projectType==='complaints'||projectType==='reports'||projectType==='inquiries'||projectType==='suggestions'){
    if(/مغلق|منجز|تم|resolved|closed|accepted|approved/.test(st.toLowerCase())) return {score:5,scale:5,nps:null};
    if(/مفتوح|جديد|open|new/.test(st.toLowerCase())) return {score:2,scale:5,nps:null};
  }
  return {score:null,scale:5,nps:null};
}
function cellByHeader(row,headers,headerName){ const idx=headers.indexOf(headerName); return idx>=0?row[idx]:''; }
function rowsFromSheet(ws){ return XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false}); }
function buildRecordsFromPrepared(prep,mapping){
  const out=[]; const projectType=prep.projectType; const source=getProjectSource(projectType); const session=prep.target; const dateDefault=prep.date;
  prep.files.forEach(f=>f.sheets.forEach(sh=>{
    const headers=sh.headers; const start=sh.headerRow+1;
    for(let i=start;i<sh.rows.length;i++){
      const row=sh.rows[i]||[]; if(!row.some(x=>String(x??'').trim())) continue;
      const service=String(cellByHeader(row,headers,mapping.service)||cellByHeader(row,headers,mapping.comment)||`${getProjectTypeLabel(projectType)} - سجل ${i-start+1}`).trim();
      const status=String(cellByHeader(row,headers,mapping.status)||'').trim();
      const scoreObj=parseScore(cellByHeader(row,headers,mapping.score),status,projectType);
      const comment=String(cellByHeader(row,headers,mapping.comment)||'').trim();
      const date=excelDateToISO(cellByHeader(row,headers,mapping.date)) || dateDefault;
      const channel=String(cellByHeader(row,headers,mapping.channel)||source).trim();
      const priority=String(cellByHeader(row,headers,mapping.priority)||'').trim();
      if(!service && !comment && !status && scoreObj.score==null) continue;
      out.push({analysis:session,date:date||dateDefault,source,platform:channel||source,service:service||'غير محدد',question:getProjectTypeLabel(projectType),score:scoreObj.score,scale:scoreObj.scale,nps:scoreObj.nps,comment,status,priority,channel,projectType,file:f.name,sheet:sh.name});
    }
  }));
  return out;
}
async function handleFiles(fileList){
  const files=[...fileList]; if(!files.length)return;
  const projectType=document.getElementById('projectType')?.value||'generic'; const mode=document.getElementById('importMode')?.value||'new'; const date=document.getElementById('newAnalysisDate')?.value||new Date().toISOString().slice(0,10);
  let target=mode==='append'?(document.getElementById('appendTargetAnalysis')?.value||''):(document.getElementById('newAnalysisName')?.value.trim()||(`${getProjectTypeLabel(projectType)} - ${new Date().toLocaleString('ar-OM',{dateStyle:'short',timeStyle:'short'})}`));
  if(mode==='append'&&!target){ showToast('اختر جلسة موجودة قبل إضافة الملفات.','warning'); return; }
  const prepared={projectType,mode,date,target,files:[],skipped:[],mapping:{}};
  for(const file of files){
    try{
      const buf=await file.arrayBuffer(); const fileHash=await bufferHash(buf); const wb=XLSX.read(buf,{type:'array'}); const f={name:file.name,fileHash,sheets:[]};
      wb.SheetNames.forEach(sn=>{ const rows=rowsFromSheet(wb.Sheets[sn]); const headerRow=findHeaderRow(rows); const headers=(rows[headerRow]||[]).map(normalizeHeader).filter(Boolean); if(headers.length) f.sheets.push({name:sn,rows,headerRow,headers}); });
      if(!f.sheets.length){ prepared.skipped.push({file:file.name,reason:'لم يتم العثور على ورقة تحتوي عناوين أعمدة.'}); continue; }
      prepared.files.push(f);
    }catch(err){ prepared.skipped.push({file:file.name,reason:'تعذر قراءة الملف: '+(err?.message||err)}); }
  }
  if(!prepared.files.length){ renderImportWizard(prepared); showToast('لم يتم تجهيز أي ملف للمعاينة.','warning'); return; }
  const firstHeaders=prepared.files[0].sheets[0].headers; prepared.mapping=detectMapping(firstHeaders,projectType); PENDING_IMPORT=prepared; renderImportWizard(prepared); showToast('تمت قراءة الملف/الملفات بنجاح. راجع ربط الأعمدة ثم اضغط حفظ.','success'); const input=document.getElementById('fileInput'); if(input)input.value='';
}
function allHeaders(prep){ return uniq(prep.files.flatMap(f=>f.sheets.flatMap(s=>s.headers))); }
function selectHeader(name,label,headers,value){ return `<div><label>${label}</label><select id="wiz_${name}"><option value="">— غير مستخدم —</option>${headers.map(h=>`<option ${h===value?'selected':''}>${esc(h)}</option>`).join('')}</select></div>`; }
function renderImportWizard(prep){
  const box=document.getElementById('importSummary'); if(!box)return; box.style.display='block';
  if(!prep.files.length){ box.innerHTML=`<div class="import-file-skip"><b>لم تنجح قراءة الملفات</b><br>${prep.skipped.map(s=>esc(s.file)+': '+esc(s.reason)).join('<br>')}</div>`; return; }
  const headers=allHeaders(prep); const m=prep.mapping||{};
  const previewRecords=buildRecordsFromPrepared(prep,m).slice(0,20);
  const totalRows=prep.files.reduce((s,f)=>s+f.sheets.reduce((a,sh)=>a+Math.max(0,sh.rows.length-sh.headerRow-1),0),0);
  const fileCards=prep.files.map(f=>`<div class="import-file-ok"><b>${esc(f.name)}</b><br>${f.sheets.map(sh=>`${esc(sh.name)}: ${fmt(Math.max(0,sh.rows.length-sh.headerRow-1))} صف | الأعمدة: ${fmt(sh.headers.length)}`).join('<br>')}</div>`).join('');
  const skipped=prep.skipped.map(s=>`<div class="import-file-skip"><b>${esc(s.file)}</b><br>${esc(s.reason)}</div>`).join('');
  box.innerHTML=`<div class="import-confirm-box"><h3><i class="fa-solid fa-list-check"></i> معالج الاستيراد V7 - معاينة قبل الحفظ</h3><div class="wizard-steps"><span class="done">1 قراءة الملف</span><span class="done">2 اكتشاف الأعمدة</span><span class="active">3 ربط الأعمدة</span><span>4 معاينة</span><span>5 حفظ</span></div>${fileCards}${skipped}<div class="mapping-preview"><h3>ربط الأعمدة</h3><div class="mapping-grid">${selectHeader('service','الخدمة/الموضوع *',headers,m.service)}${selectHeader('score','التقييم/الرضا',headers,m.score)}${selectHeader('comment','الملاحظة/الوصف',headers,m.comment)}${selectHeader('status','الحالة',headers,m.status)}${selectHeader('date','تاريخ السجل',headers,m.date)}${selectHeader('channel','الجهة/القناة',headers,m.channel)}${selectHeader('priority','الأولوية/الخطورة',headers,m.priority)}</div><div class="actions" style="margin-top:10px"><button class="btn btn-soft" onclick="updateWizardPreview()"><i class="fa-solid fa-rotate"></i> تحديث المعاينة حسب الربط</button></div></div><div id="wizardPreviewArea">${wizardPreviewHTML(previewRecords,totalRows,prep)}</div><div class="actions" style="margin-top:12px"><button class="btn btn-primary" onclick="confirmWizardImport()"><i class="fa-solid fa-floppy-disk"></i> حفظ الجلسة وتحديث لوحة المؤشرات</button><button class="btn btn-soft" onclick="cancelWizardImport()"><i class="fa-solid fa-xmark"></i> إلغاء</button></div></div>`;
}
function getWizardMapping(){ return { service:document.getElementById('wiz_service')?.value||'', score:document.getElementById('wiz_score')?.value||'', comment:document.getElementById('wiz_comment')?.value||'', status:document.getElementById('wiz_status')?.value||'', date:document.getElementById('wiz_date')?.value||'', channel:document.getElementById('wiz_channel')?.value||'', priority:document.getElementById('wiz_priority')?.value||'' }; }
function wizardPreviewHTML(records,totalRows,prep){ const services=new Set(records.map(r=>r.service)).size; return `<div class="preview-summary"><b>ملخص المعاينة</b><br>الجلسة: ${esc(prep.target)} | النوع: ${esc(getProjectTypeLabel(prep.projectType))} | الصفوف المقروءة تقريباً: ${fmt(totalRows)} | أول سجلات معاينة: ${fmt(records.length)} | الخدمات/المواضيع في المعاينة: ${fmt(services)}</div><div class="table-wrap" style="max-height:260px"><table><thead><tr><th>الخدمة/الموضوع</th><th>التقييم</th><th>الحالة</th><th>التاريخ</th><th>الجهة</th><th>الملاحظة</th></tr></thead><tbody>${records.map(r=>`<tr><td>${esc(r.service)}</td><td>${r.score==null?'—':esc(r.score+'/'+r.scale)}</td><td>${esc(r.status)}</td><td>${esc(r.date)}</td><td>${esc(r.channel)}</td><td>${esc(String(r.comment||'').slice(0,120))}</td></tr>`).join('')||'<tr><td colspan="6">لم ينتج الربط الحالي أي سجلات. غيّر عمود الخدمة أو الملاحظة ثم اضغط تحديث المعاينة.</td></tr>'}</tbody></table></div>`; }
function updateWizardPreview(){ if(!PENDING_IMPORT)return; PENDING_IMPORT.mapping=getWizardMapping(); const records=buildRecordsFromPrepared(PENDING_IMPORT,PENDING_IMPORT.mapping).slice(0,20); const total=PENDING_IMPORT.files.reduce((s,f)=>s+f.sheets.reduce((a,sh)=>a+Math.max(0,sh.rows.length-sh.headerRow-1),0),0); document.getElementById('wizardPreviewArea').innerHTML=wizardPreviewHTML(records,total,PENDING_IMPORT); }
function isDuplicate(fileHash,cHash){ return DB.analyses.find(a=>a.fileHash===fileHash || a.contentHash===cHash); }
function confirmWizardImport(){
  if(!PENDING_IMPORT){ showToast('لا توجد عملية استيراد جاهزة.','warning'); return; }
  PENDING_IMPORT.mapping=getWizardMapping(); if(!PENDING_IMPORT.mapping.service && !PENDING_IMPORT.mapping.comment){ showToast('يجب ربط عمود الخدمة/الموضوع أو عمود الملاحظة على الأقل.','warning'); return; }
  let saved=[], skipped=[];
  PENDING_IMPORT.files.forEach(f=>{ const one={...PENDING_IMPORT, files:[f]}; const recs=buildRecordsFromPrepared(one,PENDING_IMPORT.mapping); const cHash=contentHash(recs); const dup=isDuplicate(f.fileHash,cHash); if(dup){ skipped.push(`${f.name}: تم استيراد ملف مطابق سابقاً (${dup.name})`); return; } if(!recs.length){ skipped.push(`${f.name}: لم ينتج الربط أي سجلات`); return; } const id='imp_'+Date.now()+'_'+Math.random().toString(36).slice(2,9); const rows=recs.map(r=>({...r,fileId:id,file:f.name,analysis:PENDING_IMPORT.target,importedAt:new Date().toISOString()})); DB.records.push(...rows); DB.analyses.push({id,name:PENDING_IMPORT.target,source:getProjectSource(PENDING_IMPORT.projectType),date:PENDING_IMPORT.date,file:f.name,rows:rows.length,fileHash:f.fileHash,contentHash:cHash,projectType:PENDING_IMPORT.projectType,mode:PENDING_IMPORT.mode==='new'?'جلسة جديدة':'إضافة إلى جلسة موجودة',importedAt:new Date().toISOString(),mapping:PENDING_IMPORT.mapping}); saved.push(`${f.name}: ${fmt(rows.length)} سجل`); });
  const box=document.getElementById('importSummary'); PENDING_IMPORT=null; refreshAll(); if(box){ box.style.display='block'; box.innerHTML=`<div class="import-saved"><b><i class="fa-solid fa-check"></i> تم حفظ البيانات وتحديث المؤشرات</b><br>${saved.join('<br>')||''}${skipped.length?'<hr>'+skipped.map(esc).join('<br>'):''}</div>`; } showToast(`<b>نتيجة الحفظ</b><br>${saved.join('<br>')||'لم يتم حفظ ملفات جديدة'}${skipped.length?'<br><b>تنبيهات:</b><br>'+skipped.map(esc).join('<br>'):''}`,'success'); }
function cancelWizardImport(){ PENDING_IMPORT=null; const box=document.getElementById('importSummary'); if(box){ box.style.display='none'; box.innerHTML=''; } showToast('تم إلغاء الاستيراد.','warning'); }

function exportCSV(){ const rows=DB.records.map((r,i)=>({م:i+1,التحليل:r.analysis,النوع:getProjectTypeLabel(r.projectType),المصدر:r.source,التاريخ:r.date,الخدمة:r.service,التقييم:r.score,المقياس:r.scale,NPS:r.nps,الحالة:r.status,الأولوية:r.priority,الجهة:r.channel,الملاحظة:r.comment,الملف:r.file})); if(!rows.length){showToast('لا توجد بيانات للتصدير.','warning');return;} const headers=Object.keys(rows[0]); const csv=[headers.join(',')].concat(rows.map(r=>headers.map(h=>'"'+String(r[h]??'').replace(/"/g,'""')+'"').join(','))).join('\n'); const blob=new Blob(['\ufeff',csv],{type:'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='bi_analytics_records.csv'; a.click(); URL.revokeObjectURL(a.href); }
function exportXLSX(){ if(!window.XLSX){showToast('مكتبة Excel غير متوفرة.','warning');return;} const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(DB.records),'Records'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(DB.analyses),'Files'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(serviceStatsBySource(DB.records)),'KPIs'); XLSX.writeFile(wb,'bi_analytics_export.xlsx'); }
function exportJSON(){ const blob=new Blob([JSON.stringify(DB,null,2)],{type:'application/json'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='bi_analytics_data.json'; a.click(); URL.revokeObjectURL(a.href); }


function renderReportFilters(){
  const box=document.getElementById('reportFilterChecks'); if(!box)return;
  const initialized=box.dataset.initialized==='1';
  const current=[...box.querySelectorAll('input:checked')].map(x=>x.value);
  const rows=DB.analyses.map(a=>{ const fileOnly=String(a.file||a.name||'سجل بدون ملف').split(/[\\/]/).pop(); return {id:a.id,file:fileOnly,type:getProjectTypeLabel(a.projectType||'generic'),rows:a.rows,date:a.date}; });
  box.dataset.initialized='1';
  box.innerHTML=rows.map(a=>`<label class="check-item" title="${esc(a.file)}"><input type="checkbox" value="${esc(a.id)}" ${(!initialized||current.includes(a.id))?'checked':''} onchange="renderReport()"><span><span class="file-name">${esc(a.file)}</span><em>${esc(a.type)} | ${esc(a.date)} | ${fmt(a.rows)} سجل</em></span></label>`).join('')||'<div class="empty-state">لا توجد جلسات لاختيارها.</div>';
}
function selectAllReportItems(flag){ document.querySelectorAll('#reportFilterChecks input[type=checkbox]').forEach(x=>x.checked=flag); renderReport(); }
function getReportRecords(){
  const checks=[...document.querySelectorAll('#reportFilterChecks input[type=checkbox]')];
  if(!checks.length) return DB.records;
  const ids=checks.filter(x=>x.checked).map(x=>x.value);
  if(!ids.length) return [];
  return DB.records.filter(r=>ids.includes(r.fileId));
}
function reportBarRows(items,valueFn,labelFn,maxVal){ return items.map(x=>{ const v=valueFn(x)||0; const w=Math.max(3,Math.min(100,(v/(maxVal||100))*100)); return `<div class="bar-row"><span title="${esc(labelFn(x))}">${esc(labelFn(x))}</span><div class="bar-track"><div class="bar-fill" style="width:${w}%"></div></div><b>${fmt(v,0)}</b></div>`; }).join(''); }
function statusInterpretation(st,type){ if(!st.total)return 'لا توجد بيانات كافية لإصدار حكم تنفيذي.'; if(isOpsType(type)) return st.closedRate>=85?'الأداء التشغيلي جيد مع التزام مرتفع بمعالجة السجلات.':st.closedRate>=60?'الأداء متوسط ويحتاج متابعة السجلات المفتوحة والمتأخرة.':'توجد فجوة تشغيلية واضحة تستدعي خطة معالجة عاجلة.'; return st.cs>=85?'رضا المستفيدين مرتفع ويعكس تجربة إيجابية عامة.':st.cs>=70?'الرضا جيد لكنه يحتاج تحسينات موجهة في الخدمات الأقل أداء.':'مستوى الرضا منخفض نسبياً ويتطلب خطة تحسين عاجلة.'; }

function renderReport(){
  const d=getReportRecords(); const doc=document.getElementById('reportDoc'); if(!doc)return;
  const selectedAnalyses=DB.analyses.filter(a=>d.some(r=>r.fileId===a.id));
  const byType=Object.entries(groupBy(d,'projectType')).map(([t,rows])=>({t,rows,st:projectStats(t,rows)}));
  const serviceStatsTop=serviceStatsBySource(d).sort((a,b)=>b.csat-a.csat).slice(0,10);
  const lowStats=serviceStatsBySource(d).sort((a,b)=>a.csat-b.csat).slice(0,8);
  const comments=uniqueComments(d); const neg=comments.filter(r=>isNegativeComment(r.comment)).slice(0,10); const pos=comments.filter(r=>isPositiveComment(r.comment)&&!isNegativeComment(r.comment)).slice(0,10);
  const total=d.length, services=new Set(d.map(r=>r.service)).size, sources=new Set(d.map(r=>r.source)).size, overallCsat=csat(d), overallNps=npsScore(d);
  const custom=document.getElementById('customActions')?.value||'';
  if(!d.length){ doc.innerHTML='<h1>تقرير تنفيذي</h1><div class="empty-state">لم يتم اختيار أي سجل ضمن نطاق التقرير. اختر جلسة واحدة أو أكثر من القائمة أعلاه.</div>'; return; }
  const maxCount=Math.max(...serviceStatsTop.map(x=>x.count),1);
  doc.innerHTML=`<h1>التقرير التنفيذي لتحليل البيانات المؤسسية</h1><p>يعرض هذا التقرير نتائج تحليل <b>${fmt(total)}</b> سجل ضمن <b>${fmt(selectedAnalyses.length)}</b> ملف/جلسة مختارة، تغطي <b>${fmt(services)}</b> خدمة/موضوع و<b>${fmt(sources)}</b> مصدر بيانات.</p>
  <h2>1. الملخص التنفيذي</h2><div class="report-chart-grid"><div class="report-chart"><h3>المؤشرات العامة</h3><div class="donut-mini"><span class="donut-pill">CSAT: ${pct(overallCsat)}</span><span class="donut-pill">NPS: ${overallNps==null?'—':fmt(overallNps,1)}</span><span class="donut-pill">السجلات: ${fmt(total)}</span><span class="donut-pill">الملاحظات المفيدة: ${fmt(comments.length)}</span></div><p>${esc(statusInterpretation(projectStats('survey',d),'survey'))}</p></div><div class="report-chart"><h3>حجم البيانات حسب نوع التحليل</h3>${reportBarRows(byType.map(x=>({label:getProjectTypeLabel(x.t),count:x.rows.length})),x=>x.count,x=>x.label,Math.max(...byType.map(x=>x.rows.length),1))}</div></div>
  <h2>2. تعريف المؤشرات ومنهجية القراءة</h2><table><thead><tr><th>المؤشر</th><th>الاسم العربي</th><th>ماذا يقيس</th></tr></thead><tbody><tr><td>CSAT</td><td>مؤشر رضا المستفيد</td><td>نسبة التقييمات المصنفة كراضية أو عالية الرضا.</td></tr><tr><td>NPS</td><td>مؤشر صافي المروجين</td><td>مدى استعداد المستفيد للتوصية بالخدمة، ويتراوح عادة من -100 إلى +100.</td></tr><tr><td>نسبة المعالجة/الإغلاق</td><td>مؤشر الإنجاز التشغيلي</td><td>نسبة الشكاوى/البلاغات/الاستفسارات التي تم إغلاقها أو معالجتها.</td></tr></tbody></table>
  <h2>3. ملخص الأداء حسب نوع التحليل</h2><table><thead><tr><th>نوع التحليل</th><th>السجلات</th><th>الخدمات/المواضيع</th><th>CSAT</th><th>NPS</th><th>الإغلاق/المعالجة</th><th>قراءة تنفيذية</th></tr></thead><tbody>${byType.map(x=>`<tr><td>${esc(getProjectTypeLabel(x.t))}</td><td>${fmt(x.rows.length)}</td><td>${fmt(x.st.services)}</td><td>${pct(x.st.cs)}</td><td>${x.st.ns==null?'—':fmt(x.st.ns,1)}</td><td>${pct(x.st.closedRate)}</td><td>${esc(statusInterpretation(x.st,x.t))}</td></tr>`).join('')}</tbody></table>
  <h2>4. أفضل الخدمات/المواضيع حسب الرضا</h2><div class="report-chart"><h3>أعلى الخدمات حسب CSAT</h3>${reportBarRows(serviceStatsTop.map(x=>({label:x.service,value:x.csat})),x=>x.value,x=>x.label,100)}</div><table><thead><tr><th>الخدمة/الموضوع</th><th>المصدر</th><th>السجلات</th><th>CSAT</th><th>NPS</th></tr></thead><tbody>${serviceStatsTop.map(s=>`<tr><td>${esc(s.service)}</td><td>${esc(s.source)}</td><td>${fmt(s.count)}</td><td>${pct(s.csat)}</td><td>${s.nps==null?'—':fmt(s.nps,1)}</td></tr>`).join('')}</tbody></table>
  <h2>5. الخدمات ذات أولوية التحسين</h2><table><thead><tr><th>الخدمة/الموضوع</th><th>المصدر</th><th>السجلات</th><th>CSAT</th><th>الإجراء المقترح</th></tr></thead><tbody>${lowStats.map(s=>`<tr><td>${esc(s.service)}</td><td>${esc(s.source)}</td><td>${fmt(s.count)}</td><td>${pct(s.csat)}</td><td class="action-cell">${esc(priorityAction(s))}</td></tr>`).join('')}</tbody></table>
  <h2>6. التحليل النوعي للملاحظات</h2><div class="report-chart-grid"><div class="report-chart"><h3>ملاحظات الثناء والتجارب الإيجابية</h3>${pos.map(r=>`<p>• <b>${esc(r.service)}:</b> ${esc(r.comment)}</p>`).join('')||'<p>لا توجد ملاحظات ثناء واضحة.</p>'}</div><div class="report-chart"><h3>ملاحظات التحسين والشكاوى</h3>${neg.map(r=>`<p>• <b>${esc(r.service)}:</b> ${esc(r.comment)}</p>`).join('')||'<p>لا توجد ملاحظات تحسين واضحة.</p>'}</div></div>
  <h2>7. التوصيات الاستراتيجية وخطة العمل</h2><table><thead><tr><th>المجال</th><th>الإجراء</th><th>الأولوية</th></tr></thead><tbody><tr><td>تحسين الخدمات الأقل رضا</td><td>تحليل رحلة المستفيد للخدمات ذات CSAT الأقل وتحديد أسباب الانخفاض.</td><td>عالية</td></tr><tr><td>إدارة الملاحظات</td><td>تحويل الملاحظات المتكررة إلى سجل إجراءات ومتابعتها شهرياً.</td><td>متوسطة</td></tr><tr><td>حوكمة المؤشرات</td><td>اعتماد لوحة متابعة شهرية لمؤشرات CSAT وNPS والإغلاق حسب نوع التحليل.</td><td>عالية</td></tr><tr><td>توصيات إضافية</td><td>${esc(custom)}</td><td>حسب اعتماد الإدارة</td></tr></tbody></table>
  <h2>8. مصادر البيانات ونطاق التقرير</h2><table><thead><tr><th>الجلسة</th><th>النوع</th><th>الملف</th><th>التاريخ</th><th>السجلات</th></tr></thead><tbody>${selectedAnalyses.map(a=>`<tr><td>${esc(a.name)}</td><td>${esc(getProjectTypeLabel(a.projectType||'generic'))}</td><td>${esc(a.file)}</td><td>${esc(a.date)}</td><td>${fmt(a.rows)}</td></tr>`).join('')}</tbody></table>`;
}
function refreshAll(shouldPersist=true){ fillFilters(); renderReportFilters(); renderProjectCards(); renderProjectSessions(); renderProjectDashboard(); renderKPIs(); renderCharts(); renderTopLow(); renderImports(); renderServices(); renderQual(); renderSources(); renderReport(); if(shouldPersist) persist(); }

function setupDropZone(){ const dz=document.getElementById('dropZone'); if(!dz)return; ['dragover','dragenter'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag');})); ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag');})); dz.addEventListener('drop',e=>handleFiles(e.dataTransfer.files)); }

document.addEventListener('DOMContentLoaded',()=>{ setupNav(); setupDropZone(); updateTemplateHint(); updateImportModeUI(); refreshAll(false); loadIndexedState(); });

/* ===== V12: Word export + real save snapshot + save status ===== */
function setSaveStatus(state, msg){
  const el=document.getElementById('saveStatus');
  if(!el) return;
  const icons={saved:'fa-circle-check',dirty:'fa-circle-exclamation',saving:'fa-spinner fa-spin'};
  el.className='save-status '+(state||'saved');
  el.innerHTML=`<i class="fa-solid ${icons[state]||icons.saved}"></i> ${esc(msg|| (state==='dirty'?'تعديلات غير محفوظة':state==='saving'?'جاري الحفظ...':'محفوظ'))}`;
}
// persist() is defined above using IndexedDB + lightweight localStorage to avoid quota errors.
function safeFileName(name){
  return String(name||'تقرير').replace(/[\\/:*?"<>|]+/g,'_').replace(/\s+/g,'_').slice(0,120);
}
function buildWordHTML(reportHTML){
  const generated=new Date().toLocaleString('ar-OM');
  const css=`
    @page WordSection1 { size: A4; margin: 1.5cm 1.3cm 1.5cm 1.3cm; }
    body{font-family:'Cairo','Arial',sans-serif;direction:rtl;text-align:right;color:#102033;line-height:1.8;font-size:12pt;}
    h1{font-size:22pt;color:#0f2742;border-bottom:3px solid #0b5cab;padding-bottom:8px;margin:0 0 18px;}
    h2{font-size:16pt;color:#0b5cab;border-bottom:1px solid #dbe7f3;padding-bottom:6px;margin-top:22px;}
    h3{font-size:13pt;color:#0f2742;margin:8px 0;}
    table{width:100%;border-collapse:collapse;margin:10px 0 16px;table-layout:auto;}
    th{background:#0f2742;color:#fff;padding:8px;border:1px solid #dbe7f3;font-weight:bold;}
    td{padding:7px;border:1px solid #dbe7f3;vertical-align:top;white-space:normal;}
    .report-chart-grid{display:block;}
    .report-chart{border:1px solid #dbe7f3;border-radius:10px;padding:12px;margin:10px 0;background:#f8fbff;}
    .bar-row{display:flex;align-items:center;gap:8px;margin:6px 0;}
    .bar-row span{display:inline-block;width:220px;font-weight:bold;color:#334155;}
    .bar-track{display:inline-block;width:260px;height:9px;background:#e2e8f0;border-radius:999px;overflow:hidden;vertical-align:middle;}
    .bar-fill{height:9px;background:#0b5cab;border-radius:999px;}
    .donut-pill{display:inline-block;background:#eef6ff;border:1px solid #d2e8fb;color:#0b477c;border-radius:999px;padding:5px 9px;margin:3px;font-weight:bold;}
    .action-cell{line-height:1.8;}
    .cover{background:#0f2742;color:#fff;padding:18px;border-radius:12px;margin-bottom:16px;}
    .cover .meta{color:#dbeafe;font-size:10pt;margin-top:6px;}
  `;
  return `<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><title>التقرير التنفيذي</title><style>${css}</style></head><body><div class="cover"><h1 style="color:#fff;border:0;margin:0">التقرير التنفيذي لتحليل البيانات المؤسسية</h1><div class="meta">وزارة النقل والاتصالات وتقنية المعلومات | تاريخ التصدير: ${generated}</div></div>${reportHTML}</body></html>`;
}
function exportWord(){
  try{
    renderReport();
    const doc=document.getElementById('reportDoc');
    if(!doc || !doc.innerHTML.trim()){
      showToast('لا يوجد تقرير جاهز للتصدير. افتح صفحة التقرير التنفيذي واختر نطاق التقرير أولاً.','warning');
      return;
    }
    const selectedRecords=getReportRecords();
    if(!selectedRecords.length){
      showToast('لم يتم اختيار أي بيانات ضمن نطاق التقرير. اختر سجل واحد أو أكثر ثم أعد التصدير.','warning');
      return;
    }
    const html=buildWordHTML(doc.innerHTML);
    const blob=new Blob(['\ufeff',html],{type:'application/msword;charset=utf-8'});
    const a=document.createElement('a');
    const today=new Date().toISOString().slice(0,10);
    a.href=URL.createObjectURL(blob);
    a.download=safeFileName('التقرير_التنفيذي_'+today)+'.doc';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(a.href);
    showToast('تم تصدير التقرير التنفيذي بصيغة Word بنجاح.','success');
  }catch(e){
    showToast('تعذر تصدير التقرير: '+(e?.message||e),'danger');
  }
}
function saveSnapshot(){
  try{
    setSaveStatus('saving','جاري الحفظ...');
    const name=(document.getElementById('analysisFilter')?.value && document.getElementById('analysisFilter').value!=='all')
      ? document.getElementById('analysisFilter').value
      : 'Snapshot - '+new Date().toLocaleString('ar-OM');
    DB.snapshots = Array.isArray(DB.snapshots) ? DB.snapshots : [];
    const snap={
      id:'snap_'+Date.now(),
      name,
      savedAt:new Date().toISOString(),
      analysesCount:DB.analyses.length,
      recordsCount:DB.records.length,
      servicesCount:new Set(DB.records.map(r=>r.service)).size,
      sourcesCount:new Set(DB.records.map(r=>r.source)).size,
      reportScope:[...document.querySelectorAll('#reportFilterChecks input:checked')].map(x=>x.value)
    };
    DB.snapshots.push(snap);
    DB.lastSavedAt=snap.savedAt;
    persist(true);
    setSaveStatus('saved','تم الحفظ');
    showToast(`<b>تم حفظ التحليل بنجاح</b><br>السجلات: ${fmt(snap.recordsCount)} | الملفات: ${fmt(snap.analysesCount)}<br>آخر تحديث: ${new Date(snap.savedAt).toLocaleString('ar-OM')}`,'success');
  }catch(e){
    setSaveStatus('dirty','تعذر الحفظ');
    showToast('تعذر حفظ التحليل: '+(e?.message||e),'danger');
  }
}
function markDirty(){ setSaveStatus('dirty','تعديلات غير محفوظة'); }
document.addEventListener('DOMContentLoaded',()=>{
  setSaveStatus('saved','محفوظ');
  document.addEventListener('input',e=>{
    if(e.target && (e.target.id==='customActions' || e.target.closest('#tab-imports'))) markDirty();
  }, true);
  document.addEventListener('change',e=>{
    if(e.target && (e.target.closest('#reportFilterChecks') || e.target.closest('#tab-imports'))) markDirty();
  }, true);
});
